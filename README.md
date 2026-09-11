# Trinomed Skills

AI 前沿使用的三个 skill 的私有版本源。首批版本完整保留 Trinomed 固定基线中的原件，
不包含 `resume`，不自动安装或自动启用任何 skill。

| 文件夹 | 用途 |
| --- | --- |
| [skills/skill-creator](skills/skill-creator) | 已适配 AI 前沿的创建、测试、评审与改进 skill 流程。 |
| [skills/ocr](skills/ocr) | 使用平台已有可信 OCR 工具识别图片和扫描 PDF。 |
| [skills/document-markdown](skills/document-markdown) | 使用任务镜像中的本地工具，将文档和表格转换为 Markdown。 |

管理员在有仓库权限的本地环境取得固定版本后，将上述三个文件夹分别完整放入本人网盘，
从 AI 前沿 Skills 的“从我的工作区选一个文件夹”入口逐个安装，目标选择组织层。
每个所选文件夹的根必须有 `SKILL.md`；支持文件及其相对目录一并保留。
私有仓库只作版本源，不向任务容器或安装容器传入 GitHub 凭据。

这些 skill 依赖 AI 前沿现有工具与任务镜像，不是独立模型或完整运行环境。
本次迁出没有重写正文；原文中的工具名、工作路径与使用说明均保持基线字节。

来源、许可范围与固定源 SHA 见 [SOURCES.md](SOURCES.md)。
[SOURCE-MANIFEST.json](SOURCE-MANIFEST.json) 和 [SHA256SUMS](SHA256SUMS)
记录全部 15 份 skill 原件；在仓库根可执行 `sha256sum -c SHA256SUMS` 核对。

## 技能专属测试

[tests/ai-frontier-document-markdown.test.py](tests/ai-frontier-document-markdown.test.py)
保留文档转换的 10 个测试，仅将 converter 定位改为本仓库目录。
在已具备任务镜像 Python 依赖与 LibreOffice Calc 的离线环境中执行：

```sh
python3 tests/ai-frontier-document-markdown.test.py
```

确认 10 项全部通过且无跳过；缺少依赖造成的 skip 不算完整通过。
测试源码的来源与调整独立记录在 [TEST-SOURCE-MANIFEST.json](TEST-SOURCE-MANIFEST.json)，
摘要见 [TEST-SHA256SUMS](TEST-SHA256SUMS)，不与未修改的 skill 原件摘要混记。
