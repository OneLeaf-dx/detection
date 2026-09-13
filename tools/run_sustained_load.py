# -*- coding: utf-8 -*-
r"""持續負載量測與判定（v12.3 計畫 §2 的條件 S）。

判準登記於 docs/v12.3_計畫_高階裝置即時辨識門檻.md（commit e15e916），
**早於任何平台 2 的量測**。常數寫死在這裡，與 judge_deploy_stoploss.py 同一個用意。

  協定  單一組合連續 BLOCKS=10 段。每段 CPU-only、4 threads、num_runs=25，
        另加 --min_secs=30 --max_secs=32，所以每段實際跑約 30 秒。
        段與段之間只有「重啟 benchmark activity」的間隔——
        本腳本輪詢 logcat 等到該段結束就立刻開下一段，不用固定等待，
        否則段間的閒置會讓機身降溫，量到的就不是持續負載。
  判定  最後 LAST_N=3 段的平均延遲，**每一段**都 ≤ 40 ms（end2end=False 另加 2 ms）。
        任一段沒有結果，就不能判「通過」。
  記錄  每段前後以唯讀方式讀 `dumpsys thermalservice` 的目前溫度與 `dumpsys battery`。

**為什麼不能直接用 run_mobile_benchmark.py --repeat 10**：
它的 `all_ms` 只收有結果的輪次，失敗的一段會被靜默跳過，
「最後 3 段」就會悄悄變成別的段；它也不記錄溫度。

用法（專案根目錄）：
    ANDROID_SERIAL=HA27D9E0 .venv/Scripts/python.exe tools/run_sustained_load.py \
        --model last__fp32__i416__e2e0 --tag v12.3_p2_sustained

    # 先冷卻到 CPU ≤ 47 °C（最多 10 分鐘）再開始
    ... --cool-to 47 --cool-max-min 10
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

import run_mobile_benchmark as RB  # noqa: E402  共用裝置檢查、推送、log 解析

# ── 登記的常數（v12.3 §2 的 S）──────────────────────────────────────
BLOCKS = 10
LAST_N = 3
LAT_LIMIT_MS = 40.0
E2E0_RESERVE_MS = 2.0
NUM_RUNS, THREADS, SETTING = 25, 4, "cpu"
BLOCK_FLAGS = "--min_secs=30 --max_secs=32"
POLL_S = 2
BLOCK_TIMEOUT_S = 90


def read_temps() -> dict:
    """`dumpsys thermalservice` 的「目前」溫度區塊，外加電池溫度。唯讀。"""
    out = RB.adb("shell", "dumpsys", "thermalservice").stdout
    status = re.search(r"Thermal Status:\s*(\d+)", out)
    cur: dict[str, float] = {}
    # 輸出裡有「Cached temperatures」與「Current temperatures from HAL」兩段，
    # 只取後者；找不到標頭時退而求其次取最後一段，並標記出來
    m = re.search(r"Current temperatures from HAL:(.*?)(?:\n\S|\Z)", out, re.S)
    block = m.group(1) if m else out
    for val, name in re.findall(r"mValue=([\d.]+), mType=\d+, mName=(\w+)", block):
        cur.setdefault(name, float(val))           # 同名感測器取第一個
    bat = re.search(r"temperature:\s*(\d+)", RB.adb("shell", "dumpsys", "battery").stdout)
    return {
        "thermal_status": int(status.group(1)) if status else None,
        "source": "current_from_hal" if m else "unlabeled",
        "cpu_c": cur.get("CPU"), "soc_c": cur.get("SOC"), "skin_c": cur.get("SKIN"),
        "battery_c": int(bat.group(1)) / 10 if bat else None,
    }


def run_block(remote: str) -> tuple[dict, str, float]:
    """跑一段並輪詢到結束。回傳（解析結果、log、實際耗時秒數）。"""
    args = (f"--graph={remote} --num_runs={NUM_RUNS} --num_threads={THREADS} "
            f"{RB.SETTINGS[SETTING]} {BLOCK_FLAGS}").strip()
    RB.adb("logcat", "-c")
    t0 = time.time()
    # 引號必須自己加：adb shell 不保留參數邊界（見 run_mobile_benchmark.py 檔頭）
    RB.adb("shell", "am", "start", "-S", "-n", RB.ACTIVITY, "--es", "args", f'"{args}"')
    log = ""
    while time.time() - t0 < BLOCK_TIMEOUT_S:
        time.sleep(POLL_S)
        log = RB.adb("logcat", "-d", "-s", "tflite").stdout
        if "Inference timings in us" in log:
            break
    return RB.parse_log(log), log, time.time() - t0


def cool_down(target_c: float, max_min: float) -> dict:
    t0 = time.time()
    while True:
        t = read_temps()
        waited = (time.time() - t0) / 60
        if t["cpu_c"] is not None and t["cpu_c"] <= target_c:
            print(f"▷ 冷卻完成：CPU {t['cpu_c']:.1f} °C ≤ {target_c}（{waited:.1f} 分）")
            return {**t, "waited_min": round(waited, 1), "reason": "target"}
        if waited >= max_min:
            print(f"▷ 冷卻達上限 {max_min} 分：CPU {t['cpu_c']} °C")
            return {**t, "waited_min": round(waited, 1), "reason": "max_time"}
        print(f"  冷卻中… CPU {t['cpu_c']} °C（{waited:.1f} 分）", flush=True)
        time.sleep(30)


def main() -> None:
    ap = argparse.ArgumentParser(description="持續負載量測（v12.3 §2 條件 S）")
    ap.add_argument("--model", required=True, help="Benchmark/Model/ 底下的檔名（可省略 .tflite）")
    ap.add_argument("--tag", default="sustained")
    ap.add_argument("--cool-to", type=float, default=None, help="開始前先冷卻到 CPU ≤ 此溫度（°C）")
    ap.add_argument("--cool-max-min", type=float, default=10.0)
    a = ap.parse_args()

    name = a.model if a.model.endswith(".tflite") else f"{a.model}.tflite"
    local = RB.MODEL_DIR / name
    if not local.is_file():
        raise SystemExit(f"✗ 找不到 {local}")
    reserve = E2E0_RESERVE_MS if "__e2e0" in name else 0.0

    print("═" * 78)
    device = RB.device_gate()
    RB.ensure_apk()
    remote = RB.push_model(local)
    print(f"▷ {name}：{BLOCKS} 段 × 約 30 秒，判最後 {LAST_N} 段"
          f"（門檻 {LAT_LIMIT_MS} ms，後處理保留 {reserve} ms）")
    print("═" * 78)

    cool = cool_down(a.cool_to, a.cool_max_min) if a.cool_to is not None else None

    blocks = []
    for i in range(1, BLOCKS + 1):
        before = read_temps()
        start = datetime.now().isoformat(timespec="seconds")
        res, log, secs = run_block(remote)
        after = read_temps()
        ok = "avg_us" in res
        count = None
        for c in re.findall(r"count=(\d+) first=", log):
            count = max(count or 0, int(c))
        b = {
            "block": i, "start": start, "seconds": round(secs, 1), "ok": ok,
            "avg_ms": round(res["avg_us"] / 1000, 2) if ok else None,
            "eff_ms": round(res["avg_us"] / 1000 + reserve, 2) if ok else None,
            "min_ms": round(res.get("min_us", 0) / 1000, 2) if ok else None,
            "max_ms": round(res.get("max_us", 0) / 1000, 2) if ok else None,
            "std_ms": round(res.get("std_us", 0) / 1000, 2) if ok else None,
            "inferences": count,
            "temps_before": before, "temps_after": after,
        }
        blocks.append(b)
        mark = f"{b['avg_ms']:.1f} ms（有效 {b['eff_ms']:.1f}）" if ok else "✗ 沒有結果"
        print(f"  段 {i:>2}/{BLOCKS}  {mark}  推論 {count} 次  {secs:.0f} s  "
              f"CPU {before['cpu_c']}→{after['cpu_c']} °C  SKIN {after['skin_c']} °C  "
              f"狀態 {after['thermal_status']}", flush=True)

    tail = blocks[-LAST_N:]
    complete = len(blocks) == BLOCKS and all(b["ok"] for b in tail)
    passed = complete and all(b["eff_ms"] <= LAT_LIMIT_MS for b in tail)
    worst = max((b["eff_ms"] for b in tail if b["ok"]), default=None)
    first = blocks[0]["eff_ms"] if blocks[0]["ok"] else None
    verdict = ("通過" if passed else ("未通過" if complete else "無法判定（最後 3 段有缺）"))

    print("═" * 78)
    print(f"  S 判定：{verdict}　最後 {LAST_N} 段最差 {worst} ms（門檻 {LAT_LIMIT_MS}）"
          + (f"　第 1 段 {first} ms → 衰減 {worst / first - 1:+.0%}" if first and worst else ""))
    print("═" * 78)

    RB.REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = RB.REPORT_DIR / f"benchmark_{date.today().isoformat()}_{a.tag}__{Path(name).stem}.json"
    out.write_text(json.dumps({
        "device": device, "model": name,
        "protocol": {"blocks": BLOCKS, "last_n": LAST_N, "limit_ms": LAT_LIMIT_MS,
                     "reserve_ms": reserve, "setting": SETTING, "threads": THREADS,
                     "num_runs": NUM_RUNS, "block_flags": BLOCK_FLAGS,
                     "registered_in": "docs/v12.3_計畫_高階裝置即時辨識門檻.md (e15e916)"},
        "cooldown": cool, "blocks": blocks,
        "verdict": verdict, "worst_last_n_ms": worst,
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  數據 {out.relative_to(RB.P.REPO)}")


if __name__ == "__main__":
    main()
