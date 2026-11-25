开发与评测材料，不属于正式 Skill 包。
- 测试：`pip install -r requirements.txt && python -m unittest discover -s dev/tests -v`
- 案例：`dev/cases/`（c1 复盘、c2 储能；c2_plate.png 为程序生成的测试底图，非证据）。
- 渲染核验：另装 LibreOffice 与 `requirements-render.txt`，设 `SOFFICE`；字体家族名须与 plan 一致。
- 打包：`python dev/package.py` → dist/ppt-visual-art-director-<version>.zip + SHA256SUMS。
