# 学法用法速查助手

本地离线题库速查工具。输入或粘贴题干关键词，返回答案、法条依据与关联法规。
附带一个题库采集器，可从练习接口批量更新题库。

- 速查助手：纯标准库、零第三方依赖、纯离线，打包成单文件 exe
- 题库采集器：需要联网和已登录的浏览器会话，单独打包
- 内置题库 10196 道（单选 5187 / 多选 2553 / 判断 2456），附件3 法规清单 382 部

## 快速开始

从 Releases 下载两个 exe，放同一目录：

```
学法用法速查助手.exe        双击启动，离线查题
题库采集器.exe              双击更新题库，生成同目录 questions.json
题库采集器.exe --probe-only  只验证会话和接口
```

速查助手启动后优先读同目录 `questions.json`，没有就用内置题库。
三个页签：速查 / 考试范围 / 使用说明。支持手动粘贴、剪贴板监听、窗口置顶、Esc 最小化。

## 源码运行

需要带 tkinter 的 Python 3.10+（Windows 上推荐 Python 3.13）。

```bash
python -m pip install requests pyinstaller     # 只有采集和打包需要
python src/main.py                              # 启动速查助手
python src/main.py --selftest                   # 自检，输出 selftest.txt
python tools/fetch_practice_bank.py --probe-only
```

## 打包

```bash
python build.py           # dist/学法用法速查助手.exe
python build_fetcher.py   # dist/题库采集器.exe
```

两个脚本自动挑选带 tkinter 的解释器，只打包需要的数据文件，构建产物在 `dist/`。

## 更新题库

采集器从练习接口取题，返回对象自带答案与解析；自动读取 Edge/Chrome 本地存储里的会话凭证，用随机采样发现 ID 段后并发扫描，断点缓存在 `tools/cache/raw_questions.json`。

```bash
python tools/fetch_practice_bank.py --resume --sample 0 --skip-scan   # 从缓存离线重建
python tools/fetch_practice_bank.py --sample 600 --workers 6          # 联网增量采集
```

也可以导入 Excel 或 Word：

```bash
python tools/import_bank.py 题库导出.xlsx --dry-run
python tools/import_docx.py 我的题库.docx
```

`tools/import_bank.py` 兼容「广西普法考试题库录入」脚本导出的三张工作表。速查助手界面上也有「导入题库 Excel」按钮，导入前自动留 `.bak`。

## 目录结构

```
src/
  main.py            入口与 --selftest / --import
  ui.py              tkinter 界面
  bank.py            2-gram 倒排 + difflib 精排 + 选项一致性加权
  clipboard.py       GetClipboardSequenceNumber 变化检测
  paths.py           源码态 / _MEIPASS 态资源路径
  xlsx_import.py     纯标准库 xlsx 解析与合并去重
  data/
    questions.json   10196 道题
    laws_scope.json  附件3 的 382 部法规清单
tools/
  fetch_practice_bank.py   练习接口采集器
  import_bank.py           Excel / JSON 补库
  import_docx.py           Word 题库解析
  build_bank.py            初始题库生成
tests/                     21 项单元测试与夹具
build.py / build_fetcher.py
```

## 数据说明

题库来自练习接口返回数据、公开检索到的历年真题和本地整理资料，合并时按题干归一化去重。
接口未返回选项文本的题目保留答案与解析，并在界面标提示。
法规全文未内置，答案依据优先用题库自带解析；法规清单用于分类与关联统计。

## 测试

```bash
python -m unittest discover -s tests -v
```

当前 21/21 通过。索引构建约 0.8s，单次查询 1~11ms。

## License

MIT
