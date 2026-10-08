# pz-001

运行 `python -B scripts/generate_report.py` 生成 `output/report.md` 和 `output/data.xlsx`。依赖见 `requirements.txt`。

脚本使用此前核对的王强股东签字页清单，不重新解析Word文件。Excel在内存中打包，产物仅写入仓库的 `output/`。合并PDF一并纳入Git，便于从PR分支取回。
