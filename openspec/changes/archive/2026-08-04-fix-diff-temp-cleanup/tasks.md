## 1. 修复 diff_two_files

- [x] 1.1 临时文件名加入 `uuid.uuid4().hex[:6]` 后缀
- [x] 1.2 输出文件名加入 UUID 后缀（避免 Excel 锁）
- [x] 1.3 扩大 try/finally 范围包裹 recalc_batch
- [x] 1.4 回归测试 14/14 PASS
- [x] 1.5 重新打包 exe
