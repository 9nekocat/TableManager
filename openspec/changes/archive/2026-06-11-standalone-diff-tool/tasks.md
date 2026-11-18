## 1. Core

- [x] 1.1 在 `diff_engine.py` 新增 `diff_two_files(old_path, new_path, output_dir)`
- [x] 1.2 创建 `diff_standalone.py` CLI 入口，含路径推导

## 2. Build

- [x] 2.1 创建 `diff_standalone.spec` PyInstaller 配置
- [x] 2.2 打包 `dist\diff.exe`

## 3. Verify

- [x] 3.1 用两个 Effect.xlsx 版本手动测试 diff.exe
- [x] 3.2 确认 TortoiseSVN 配置可用
