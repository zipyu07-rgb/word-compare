# Word 文本比对器

拖入两份 Word 文档，一键生成可逐条接受或拒绝的修订稿。适合合同、制度和其他需要核对版本的文档。

[下载 Windows 便携版](https://github.com/zipyu07-rgb/word-compare/raw/refs/heads/main/WordCompare.exe)。无需安装本工具；需要 Windows 和已安装、激活的 Microsoft Word 桌面版。

1. 左侧放原稿，右侧放修改稿；支持 .doc 和 .docx。
2. 填写修订人，按需选择是否比较格式，并可改输出文件名。
3. 点击「生成对比稿」。默认存到修改稿所在文件夹，重名会自动编号。

工具调用本机 Word 的原生比较功能。原稿和修改稿以只读方式打开，工具不会上传文档。设置保存在当前 Windows 账户的本地配置目录。加密文档不支持；已有修订按 Word 的比较规则处理。

## 从源码构建

在 Windows PowerShell 中执行 `.\build.ps1`。脚本使用 .NET Framework 的 C# 编译器，并自动生成程序图标。
