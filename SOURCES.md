# 固定来源与许可证范围

首批来源仓库：`https://github.com/Springbudder/trinomed`。
固定来源提交：`987271c64e529274e1573226ce5107c23d4c9d06`。
来源目录：`scripts/ai-frontier-sandbox/skills/`。
提取时直接读取该提交的 Git blobs，不取主检出的未提交修改；原件逐字节一致。

## skill-creator

- 本仓库的 `skills/skill-creator/` 完整保留 Trinomed 基线中的裁剪适配版及全部支持文件。
- 原上游：`https://github.com/anthropics/skills`。
- 原上游固定提交：`41bbe19d1a1a7eaab5e7bb9050a417e5c6cffc8f`。
- 原上游目录：`skills/skill-creator`。
- 原许可证：[skills/skill-creator/LICENSE.txt](skills/skill-creator/LICENSE.txt)，Apache License 2.0，字节未改。
- 原适配记录：[skills/skill-creator/PROVENANCE.json](skills/skill-creator/PROVENANCE.json)，字节未改。
  其中逐项列出 Trinomed 既有裁剪、未纳入的上游文件，以及保留文件的上游和适配后摘要。
- 本次迁出没有追加裁剪，也没有更改 `SKILL.md`、脚本、参考资料、grader 或 eval-viewer。

## ocr、document-markdown 与测试

`skills/ocr/`、`skills/document-markdown/` 来自上述 Trinomed 固定提交。
在这两个目录、迁出的测试及该提交的仓库根未发现独立 LICENSE 或明确的开源许可声明；
因此这里只记录公司内部版本源，不为它们新增、推定或套用 Apache、OFL 等公开开源许可。
`skill-creator/LICENSE.txt` 的存在不表示它适用于整个仓库。

`tests/ai-frontier-document-markdown.test.py` 来自同一提交的
`__tests__/ai-frontier-document-markdown.test.py`，仅调整 converter 文件定位，测试语义不改。
运行时依赖仍由 AI 前沿任务镜像提供，本仓库未复制或重新许可这些第三方依赖。

## 摘要清单

- `SOURCE-MANIFEST.json`：15 份未修改的 skill 原件，含原路径、源 Git blob、长度及 SHA-256。
- `SHA256SUMS`：仅这 15 份原件，路径以本仓库根为基准。
- `TEST-SOURCE-MANIFEST.json`：测试原位置、原摘要、迁出后摘要及唯一定位调整。
- `TEST-SHA256SUMS`：迁出后的测试源码摘要。

以上源提交是内容来源锚点；本仓库提交 SHA 在实际提交后由 Git 记录，不在文件中自造。
