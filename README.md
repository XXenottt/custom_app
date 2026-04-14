# Custom App (Tauri)

已实现内容：

- 桌面应用壳子（菜单、路由、状态存储）
- Excel 处理界面：输入目录 + 输出文件路径 + 执行按钮
- 按 `src/extract.py` 逻辑在 Rust 端处理文件并输出 `xlsx`

## 功能说明（对应 extract.py）

- 扫描输入目录中的文件
- 文件名包含 `INV` 的 Excel：按指定列重排并汇总到 `Sheet1`
- 文件名包含 `packing list` 的 Excel：提取 `Total` 与 `net weight(KG)` 到 `Sheet2`
- 生成输出 Excel 文件

## 开发运行

```bash
cd /home/zxh/custom_app
source ~/.cargo/env
npm install
npm run tauri dev
```

## 打包

```bash
npm run tauri build
```

产物目录：

- `src-tauri/target/release/custom_app` (Linux 可执行文件)
- `src-tauri/target/release/bundle/deb/`
- `src-tauri/target/release/bundle/rpm/`
