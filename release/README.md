# sow_merge_tool 使用指南

## 1. 适用范围
- 用于 Excel `.xlsx` 文件的对比与冲突合并（SVN/TortoiseSVN 工作流）。
- 当前仅支持 `.xlsx`，不支持 `.xlsm`。

## 2. 环境要求
- Windows 10/11
- 无需安装 Python 或其他运行库（已打包为单文件 EXE）

## 3. 直接运行
双击：
`dist\sow_merge_tool.exe`

## 4. 命令行参数
### 4.1 普通对比
```bat
sow_merge_tool.exe --base "A.xlsx" --mine "B.xlsx"
```

### 4.2 SVN 冲突合并（推荐）
```bat
sow_merge_tool.exe --base "BASE.xlsx" --mine "MINE.xlsx" --theirs "THEIRS.xlsx" --merged "MERGED.xlsx"
```

### 4.3 单文件自动识别冲突
如果传入的是冲突文件路径（同目录包含 `.mine` / `.rXXXX`），工具会自动识别：
```bat
sow_merge_tool.exe "C:\path\conflict.xlsx"
```

## 5. 冲突合并流程
1. 进入冲突界面后，点击行号箭头或“采用对方(B)”/“保留我的(A)”进行覆盖。
2. 完成后点击“保存Merged并退出”。
3. 如果提示 Excel 占用，请关闭 Excel 再保存。

## 6. 常见问题
### 6.1 保存失败（Permission denied）
- 通常是 Excel 或 SVN 正在占用目标文件。
- 关闭 Excel 再保存即可。

### 6.2 打不开 / 无反应
- 确保 TortoiseSVN 已正确注册 diff/merge 工具（见下一节）。

## 7. TortoiseSVN 注册表配置
（管理员不需要，仅限当前用户）

### 7.1 Diff
```bat
reg add "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsx /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --title \"%bname\"" /f
```

### 7.2 Merge
```bat
reg add "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsx /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --theirs \"%theirs\" --merged \"%merged\" --title \"%bname\"" /f
```

### 7.3 备用（XLSX 节点）
```bat
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --title %bname" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --theirs %theirs --merged %merged --title %bname" /f
```

## 8. 日志
日志路径：
```
%TEMP%\sow_merge_tool_debug.log
```

---

如需更新版本，请替换 `sow_merge_tool.exe` 后重新注册即可。
