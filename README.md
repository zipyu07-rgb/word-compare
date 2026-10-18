# Word 对比

一个轻量的 Windows 桌面工具：拖入原稿和修改稿，一键调用本机 Microsoft Word，生成带原生修订标记的对比稿。

## 下载与使用

点击 [下载 WordCompare.exe](https://github.com/zipyu07-rgb/word-compare/raw/refs/heads/main/WordCompare.exe)，保存后双击运行。

需要 Windows 和已安装、激活的 Microsoft Word 桌面版。无需安装 Python，无需账号或联网。

1. 左侧选择或拖入原稿，右侧选择或拖入修改稿（支持 `.docx`、`.doc`）。
2. 设置修订人；默认忽略格式变化，可取消勾选。
3. 点击“生成对比稿”。默认保存到修改稿所在文件夹，也可指定保存位置。
4. 完成后可打开对比稿或所在文件夹。生成的 `.docx` 可以在 Word 中接受或拒绝修订。

文件名包含时间，重名时自动编号。输入文档以只读方式打开，所有文档在本机处理。

## 设置与限制

修订人、忽略格式、自动打开和保存位置设置会自动记住，保存在 `%LOCALAPPDATA%\FrankWordCompare\settings.xml`。

不支持加密文件。已有修订的文档按 Word 自身对比规则处理，建议先确认修订状态。首次使用如出现 Word 激活、许可或文件修复提示，请先在 Word 中完成处理。

## 从源码编译

在项目目录运行 PowerShell：

```powershell
.\build.ps1
```

脚本使用 Windows .NET Framework 4.x 自带的 C# 编译器，不需要额外下载依赖。源码入口为 `WordCompare.cs`。

## 验证

已在本机通过 Microsoft Word 生成测试对比稿，并检查输出文档的 XML，确认包含插入、删除及指定修订人。

开发用集成测试会在指定文件夹创建两份示例文档并生成对比稿：

```powershell
.\WordCompare.exe --self-test C:\Temp\WordCompareTest
```

请使用专门的空测试文件夹；测试结果写入 `test-result.txt`，错误写入 `test-error.txt`。
