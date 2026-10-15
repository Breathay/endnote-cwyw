# EndNote 建库与 Word 引文技能

[English](README.md) · [技能入口](SKILL.md) · [MIT 许可](LICENSE)

把论文参考文献导入 EndNote，并通过 Word 的 Cite While You Write 插件生成
**真正可编辑、可更新的引文域和参考文献表**。

这套技能把实际操作中容易反复踩坑的步骤固定下来：先核对已有文献、只导入缺失
条目，再读取实际记录号、制作新版本论文、调用原生插件，最后核对引文与版式。
它不会伪造 EndNote 字段，也不会直接修改文献库数据库。

## 包含什么

- 可由 Codex 调用的完整技能，入口是 `SKILL.md`。
- Python 工具：检查 Word、生成导入 XML、匹配 EndNote 真实记录号、精确替换
  手工引用、保留已有引文域、核对字段及编号。
- PowerShell 工具：只读检查插件连接，调用 `ENFormatBibliography`，等待文献表
  写入完成后保存，避免“宏已返回但文件尚未生成完整”的问题。
- 通用数据示例、自动测试、Windows/Linux CI，以及故障处理说明。

仓库不包含真实论文、个人文献库、凭据、出版方 PDF 或用户自定义 `.ens` 样式。

## 环境

Python 3.10 及以上；现代 DOCX 的安全编辑需要 `lxml`。最终生成原生引文域还需要
Windows、桌面版 Microsoft Word、EndNote、已启用的 CWYW 插件和已安装的输出样式。

原生流程验证基线为 EndNote 20 与 64 位 Word。其它版本先探测连接，再用副本做
小规模测试。离线测试通过不代表所有 Office/EndNote 组合都已验证。

## 安装为个人技能

```powershell
git clone https://github.com/Breathay/endnote-cwyw.git
cd endnote-cwyw
python -m pip install -r requirements.txt
python scripts/install_skill.py
```

默认安装到 `$CODEX_HOME/skills/endnote-cwyw`；未设置 `CODEX_HOME` 时使用
`~/.codex/skills/endnote-cwyw`。更新已有安装时：

```powershell
python scripts/install_skill.py --replace
```

更新会在技能发现目录之外备份旧版，不修改全局设置，不删除已有文件。

下次可以直接提出请求：

> 使用 $endnote-cwyw，把这篇论文的参考文献导入 EndNote，并创建下一个 Word
> 版本，用我已安装的 Nature 自定义样式插入可更新引文。图注中直接插入引用。

## 标准流程

1. **检查来源**：确认源论文、新版本文件名、目标文献库和样式，读取已有引文域。
2. **整理元数据**：优先使用已核对的参考文献与库中条目，只补齐缺失的信息。
3. **核对与导入**：先匹配目标库，只导入缺失条目，随后读取实际记录号。
4. **制作新版本**：用精确文本范围和引用 ID 生成计划，保留原论文、图片和已有域。
5. **原生格式化**：通过 CWYW 生成引文和文献表，等待完整、稳定后保存。
6. **验收**：检查引用对象、显示编号、原文与图片完整性，并渲染检查全部页面。

完整命令与 JSON 格式见 [workflow.md](references/workflow.md)。例如：

```powershell
python scripts/endnote_workflow.py inspect manuscript.docx --output work/inspection.json
python scripts/endnote_workflow.py map work/records.json --library Library.enl --output work/mapping.json --missing-xml work/missing.xml
```

缺失条目通过 EndNote 的 **File → Import → File → EndNote generated XML**
导入一次，再重新执行匹配。`missing` 与 `conflicts` 均为空后才准备论文。

```powershell
python scripts/endnote_workflow.py prepare manuscript.docx --output manuscript-next.docx --mapping work/mapping.json --plan work/plan.json --audit work/prepare-audit.json
pwsh -NoProfile -File scripts/endnote_word.ps1 -Action Format -Document manuscript-next.docx -StylePath 'C:/path/to/installed/Style.ens' -ExpectedCitations 25 -ExpectedReferences 18 -PdfPath work/final.pdf
python scripts/endnote_workflow.py verify manuscript-next.docx --audit work/prepare-audit.json --expected-citations 25 --expected-references 18
```

示例中的数量必须替换为实际数量。引文组数与被引文献数是两个不同指标。
`-Action Probe` 仅检查连接，不会调用宏或修改文档。PowerShell 7 会自动转入
Windows PowerShell 5.1 完成 Word COM 调用，不更改系统执行策略。

## 已固化的关键经验

- Journal Article 的 EndNote **XML 类型是 17**，不是本地数据库类型 0。
  当前 XML 导出器仅支持这个已经验证的类型；其它类型用已验证的原生导出/RIS。
- 论文编号不等于库中记录号；非空库、重复条目和多个库尤其不能假定 `#1..#N`。
- 作者字段必须分开记录完整姓名；`et al.`、Unicode 乱码和姓氏拆分错误要先处理。
- 内部追踪编号不能写入 `Label`、`Short Title` 等可能被自定义样式打印的字段。
- 预印本编号不能当页码处理；已核对的版本年份不应被首次提交年份覆盖。
- 宏返回后仍需等文献表完整稳定，不能只看 Word 总字段数。
- 只读匹配库；不修改 SQLite 记录、触发器、自定义排序函数或排序缓存。
- 图注按用户要求合并引用；不能因为移除 Sources 段就丢失面板与文献的对应关系。

更多故障与恢复办法见 [troubleshooting.md](references/troubleshooting.md)。

## 测试

```powershell
python -m unittest discover -s tests -v
```

测试使用虚构记录、临时库和解析夹具，不会修改打开的 Word 或真实 EndNote 库。
原生格式化与页面布局仍需另行验证。真实论文及库文件已列入 `.gitignore`，发布前
仍应检查实际提交清单。

## 许可

本项目采用 [MIT 许可](LICENSE)，版权归 Breathay。EndNote、Microsoft Word
及自定义输出样式不包含在本项目中，其许可不因本仓库开源而改变。
