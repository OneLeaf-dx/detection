# weight/：各版本權重整理區

這裡的權重檔**全部是複製品**，由 [`tools/collect_weights.py`](../tools/collect_weights.py) 從原位置複製而來；
原位置沒有移動（`Train Code/*/Train_output/extracted/` 的版面由 AGENTS.md 固定，`Benchmark/Model/` 是 benchmark 的輸入目錄）。
整個資料夾都進版控，clone 下來就拿得到所有版本；每個檔案的來源與 SHA-256 記在 [`manifest.json`](manifest.json)。
`.pt` 與原位置已追蹤的檔案內容相同，git 只存一份，不會讓 repo 變大。整個資料夾刪掉也能重建：

```bash
.venv/Scripts/python.exe tools/collect_weights.py
```

## 給 App 端：[`best/v13/`](best/v13/README.md)

交付模型 v13 的三種格式各一顆（`.pt`／ONNX／TFLite），前處理與輸出格式寫在該資料夾的 README。

## 版本一覽

| 版本 | 資料集 | 來源 run | 對外報告採用 | YOLO `.pt` | ONNX | TFLite |
| --- | --- | --- | --- | --- | --- | --- |
| v8 | v5 | `Train Records/YOLO26n_P2_Citrus_MuSGD_v8` | `best.pt` | best、last | — | — |
| v9 | v5r | v9 消融的 A0 基準臂（`Phase4`） | `best.pt` | best | — | — |
| v10 | v5.5 | `v55_v10` | `best.pt` | best | — | — |
| v11 | v5.6 | `v5.6_v11` | `best.pt` | best、last | — | — |
| v11.5 | v5.6 | `v5.6_v11_5` | `last.pt` | best、last | 640、320 | fp32@640、fp32@320、w8a32@640 |
| v12s | v5.6 | `v5.6_v12s`（s 尺度） | `last.pt` | best、last | — | — |
| **v13（交付）** | v5.7 | `v5.7_v11_5` | `last.pt` | best、last | 640、320 | fp32@640、fp32@320、w8a32@640 |

每個版本的資料夾是 `<版本>/YOLO/`、`<版本>/ONNX/`、`<版本>/TFLite/`。
檔名 `<版本>__<variant>__i<imgsz>__e2e1`：`e2e1` 表示保留 NMS-free 輸出（模型直接吐最終框）。

- **各版本的 mAP 不能互比。** 評估用的資料集不同；v13 在 v5.7 上的 0.861 與 v5.6 以前不可比，
  因為 `Thrips_Damage` 換了框定義（見 [docs/v13_報告_最終交付與週會決議.md](../docs/v13_報告_最終交付與週會決議.md) §2.2）。
  每個檔案的 mAP 與評估資料集記在 `manifest.json`
- 只有 v11.5 與 v13 匯出過 ONNX／TFLite。其他版本要用時，以 `tools/export_tflite.py`、`tools/export_other_formats.py` 在匯出容器裡轉（見 [Benchmark/export/README.md](../Benchmark/export/README.md)）
- ONNX／TFLite 只收標準組合；v12.1 掃描用的其他組合（int8、w8a16、224–512、`max_det` 等）留在 `Benchmark/Model/`
- 不收：v13 外部影像臂（`v5.7ext_v11_5`，判準未通過、不交付）、v9 其他消融臂（在 `Train Code/v9/Train_output/ablation_*.zip`）、逐 10 輪檢查點
