# -*- coding: utf-8 -*-
r"""把各版本的權重（YOLO `.pt`／ONNX／TFLite）集中複製到 `weight/`，並寫一份附 SHA-256 的清單。

**只複製，不搬移。** 原位置都有別的東西依賴：
`Train Code/*/Train_output/extracted/` 的版面由 AGENTS.md 固定，
`Benchmark/Model/` 是 benchmark 工具鏈的輸入目錄。`weight/` 只是方便取用的整理區，
整個刪掉也能用這支重建。**整個資料夾進版控**：`.pt` 與原位置已追蹤的檔案內容相同，
git 只存一份；新增的只有 ONNX／TFLite。

版面：

    weight/
      README.md          手寫
      manifest.json      本腳本產生
      <版本>/YOLO/<版本>_{best,last}.pt
      <版本>/ONNX/<版本>__fp32__i<imgsz>__e2e1.onnx
      <版本>/TFLite/<版本>__<variant>__i<imgsz>__e2e1.tflite
      best/<版本>/        三種格式各挑一顆，給 App 端直接用（另有手寫的 README.md）

ONNX／TFLite 只收「標準組合」（TFLite fp32@640、fp32@320、w8a32@640，ONNX 640／320），
v12.1 掃描用的其他組合留在 `Benchmark/Model/`。
只有 v11.5 與 v13 匯出過；其他版本只有 `.pt`，這支**不會**替它們補匯出。

**逐框驗收沒通過的 TFLite 不收**，改記在 manifest 的 `excluded`——
`export_tflite.py` 對這種檔案的指示是「不要拿去 benchmark」，放進整理區只會被誤拿去用。

**`best/` 只從交付版本裡挑，不跨版本比 mAP。** 各版本評估用的資料集不同
（v13 是 v5.7，`Thrips_Damage` 換了框定義），數字不能互比。
`.pt` 放對外報告的那一顆（v13 是 `last.pt`，即交付文件登記的 SHA-256）；
ONNX／TFLite 各挑完整 test mAP50 最高的一顆。

用法：
    .venv/Scripts/python.exe tools/collect_weights.py
"""

from __future__ import annotations

import hashlib
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import dataset_paths as P  # noqa: E402

OUT = P.REPO / "weight"
MODEL_DIR = P.REPO / "Benchmark" / "Model"
HISTORY = P.REPO / "docs" / "training_history.json"
DETECT = "Train Code/{v}/Train_output/extracted/runs/detect/{run}/weights"

# 版本代號 → (training_history.json 的鍵, 權重資料夾, 匯出檔的檔名 stem)
# 資料集與「對外報哪一顆」一律從 training_history.json 讀，不在這裡重抄一份。
VERSIONS = {
    "v8": ("v8", "Train Records/YOLO26n_P2_Citrus_MuSGD_v8/weights", None),
    "v9": ("v9_A0", "Train Code/v9/Train_output/Phase4", None),
    "v10": ("v10", "Train Code/v10/Train_output/extracted", None),
    "v11": ("v11", DETECT.format(v="v11", run="v5.6_v11"), None),
    "v11.5": ("v11.5", DETECT.format(v="v11.5", run="v5.6_v11_5"), "last"),
    "v12s": ("v12s", DETECT.format(v="v12s", run="v5.6_v12s"), None),
    "v13": ("v5.7_v11_5", DETECT.format(v="v11.5", run="v5.7_v11_5"), "v13"),
}

# 2026-09-15 週會定案的交付模型（docs/v13_報告_最終交付與週會決議.md）
BEST_VERSION = "v13"
# 同一份文件 §2.1 登記的 v13 權重雜湊
V13_SHA256 = "8d6a2fd928d9544e9b24dc5aff73ddd3a7336199ccb3de0be7bf817268e6825d"

TFLITE_STD = ("fp32__i640__e2e1", "fp32__i320__e2e1", "w8a32__i640__e2e1")
ONNX_STD = (640, 320)


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def rel(p: Path) -> str:
    return p.relative_to(P.REPO).as_posix()


def place(src: Path, dest: Path) -> dict:
    """複製（內容相同就跳過），回傳清單條目。"""
    if not src.is_file():
        raise SystemExit(f"✗ 找不到來源 {rel(src)}")
    digest = sha256(src)
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.is_file() and sha256(dest) == digest:
        state = "不變"
    else:
        shutil.copy2(src, dest)
        state = "複製"
    print(f"  {state}  {dest.relative_to(OUT).as_posix()}")
    return {"path": dest.relative_to(OUT).as_posix(), "source": rel(src),
            "bytes": src.stat().st_size, "sha256": digest}


def collect_version(ver: str, key: str, wdir: str, stem: str | None, h: dict) -> dict:
    reported = h["reported_weight"]
    print(f"\n▷ {ver}（{key}，資料集 {h['dataset']}，對外報 {reported}）")
    entry = {"history_key": key, "dataset": h["dataset"],
             "reported_weight": reported, "files": [], "excluded": []}

    for kind in ("best", "last"):
        src = P.REPO / wdir / f"{kind}.pt"
        if not src.is_file():
            continue
        f = place(src, OUT / ver / "YOLO" / f"{ver}_{kind}.pt")
        f.update(format="pytorch", reported=f"{kind}.pt" == reported)
        if ver == "v13" and kind == "last" and f["sha256"] != V13_SHA256:
            raise SystemExit(f"✗ v13 last.pt 的 SHA-256 與交付文件不符：{f['sha256']}")
        entry["files"].append(f)

    if not stem:
        return entry

    rp = json.loads((MODEL_DIR / f"{stem}__export_report.json").read_text(encoding="utf-8"))
    for tag in TFLITE_STD:
        r = rp["results"][tag]
        verdict = r.get("verify", {}).get("verdict")
        ev = {"dataset": rp["dataset"], "mAP50": r.get("val", {}).get("mAP50"), "verify": verdict}
        src = MODEL_DIR / f"{stem}__{tag}.tflite"
        if verdict != "通過":
            print(f"  略過  {src.name}（逐框驗收：{verdict}）")
            entry["excluded"].append({"source": rel(src), "reason": f"逐框驗收{verdict}",
                                      "verify_detail": r.get("verify"), "eval": ev})
            continue
        f = place(src, OUT / ver / "TFLite" / f"{ver}__{tag}.tflite")
        f.update(format="tflite", imgsz=r["imgsz"], eval=ev)
        entry["files"].append(f)

    ro = json.loads((MODEL_DIR / "other" / f"{stem}__other_formats_report.json")
                    .read_text(encoding="utf-8"))
    for s in ONNX_STD:
        r = ro["results"][f"onnx__i{s}"]
        f = place(MODEL_DIR / "other" / f"{stem}__onnx__i{s}.onnx",
                  OUT / ver / "ONNX" / f"{ver}__fp32__i{s}__e2e1.onnx")
        f.update(format="onnx", imgsz=s,
                 eval={"dataset": ro["dataset"], "mAP50": r.get("val", {}).get("mAP50")})
        entry["files"].append(f)
    return entry


def collect_best(ver: str, entry: dict) -> list[dict]:
    """`.pt` 取對外報告那一顆；ONNX／TFLite 取完整 test mAP50 最高的一顆。"""
    print(f"\n▷ best/{ver}（App 端交付）")
    picks = []
    for fmt in ("pytorch", "onnx", "tflite"):
        cands = [f for f in entry["files"] if f["format"] == fmt]
        if fmt == "pytorch":
            cands = [f for f in cands if f["reported"]]
        else:
            cands = [f for f in cands if f["eval"].get("mAP50") is not None]
        if not cands:
            raise SystemExit(f"✗ {ver} 沒有可選的 {fmt} 檔案（缺 mAP 或未通過驗收）")
        pick = max(cands, key=lambda f: (f.get("eval") or {}).get("mAP50") or 0)
        f = place(OUT / pick["path"], OUT / "best" / ver / Path(pick["path"]).name)
        f.update(format=fmt, source=pick["source"], picked_from=pick["path"],
                 **({"eval": pick["eval"], "imgsz": pick["imgsz"]} if "eval" in pick else {}))
        picks.append(f)
    return picks


def main() -> None:
    history = json.loads(HISTORY.read_text(encoding="utf-8"))["runs"]
    manifest = {"generated_by": "tools/collect_weights.py",
                "note": "weight/ 底下的檔案都是複製品，原位置見各條目的 source。"
                        "各版本評估的資料集不同，mAP 不可跨資料集比較。",
                "best": {}, "versions": {}}

    for ver, (key, wdir, stem) in VERSIONS.items():
        manifest["versions"][ver] = collect_version(ver, key, wdir, stem, history[key])
    manifest["best"][BEST_VERSION] = collect_best(BEST_VERSION, manifest["versions"][BEST_VERSION])

    mp = OUT / "manifest.json"
    mp.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")

    print("\n" + "═" * 60)
    print(f"  {'版本':<8}{'資料集':<10}{'YOLO':>6}{'ONNX':>6}{'TFLite':>8}{'略過':>6}")
    print("═" * 60)
    for ver, e in manifest["versions"].items():
        n = {k: sum(f["format"] == k for f in e["files"]) for k in ("pytorch", "onnx", "tflite")}
        print(f"  {ver:<8}{e['dataset']:<10}{n['pytorch']:>6}{n['onnx']:>6}{n['tflite']:>8}"
              f"{len(e['excluded']):>6}")
    print("═" * 60)
    for f in manifest["best"][BEST_VERSION]:
        m = (f.get("eval") or {}).get("mAP50")
        print(f"  best/{BEST_VERSION}  {Path(f['path']).name:<34}"
              + (f"mAP50 {m:.5f}" if m is not None else "對外報告的權重"))
    print(f"  清單 {rel(mp)}")


if __name__ == "__main__":
    main()
