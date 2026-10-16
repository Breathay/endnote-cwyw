# EndNote 建库与 Word 引文

[English](README.md) · [技能说明](SKILL.md) · [MIT 许可](LICENSE)

把参考文献导入 EndNote，再将论文里手写的引用替换成可更新的 EndNote 引用。
处理完成后，你可以继续在 Word 中添加文献、调整引用顺序或更换期刊格式，
文末的参考文献表也会随之更新。

这是一个供 Codex 使用的技能，适合整理整篇论文的引用，也可以只补几处引用。
已有的 EndNote 引用会保留；修改后的论文另存为新版本，原稿留在原处。

## 开始之前

请在 Windows 上安装桌面版 Microsoft Word 和 EndNote，并确认 Word 中能看到
EndNote 的 Cite While You Write（CWYW）插件。准备好论文、要使用的文献库，
以及已安装的期刊格式；自定义 `.ens` 格式也可以使用。

安装技能需要 Git 和 [uv](https://docs.astral.sh/uv/getting-started/installation/)。
项目使用 `pyproject.toml` 管理依赖，`uv.lock` 记录依赖版本。
下面的命令会准备所需的 Python 和依赖。

## 安装

在终端中运行：

```powershell
git clone https://github.com/Breathay/endnote-cwyw.git
cd endnote-cwyw
uv sync --locked
uv run scripts/install_skill.py
```

默认安装到 `~/.codex/skills/endnote-cwyw`。如果设置了 `CODEX_HOME`，
则安装到该目录下的 `skills/endnote-cwyw`。

更新时，在仓库目录中运行以下命令。安装程序会先备份已安装的旧版技能。

```powershell
git pull --ff-only
uv sync --locked
uv run scripts/install_skill.py --replace
```

## 怎么使用

在 Codex 中说明论文位置、目标文献库和想用的格式，例如：

> 使用 $endnote-cwyw，把 manuscript-v12.docx 的参考文献导入我的 EndNote
> 文献库，另存为 v13，用已安装的 Nature 自定义格式插入引用。
> 图注中的引用直接放在对应描述后面，不另列 Sources 段。

也可以只处理一部分：

> 使用 $endnote-cwyw，在这份已有 EndNote 引用的论文里补上这三篇文献。

技能会先核对库中已有的文献，再导入缺少的条目；随后在论文副本中插入引用，
由 Word 的 EndNote 插件生成编号和参考文献表，最后检查引用是否对应、图片和
正文是否完整。遇到同名文献、重复记录或信息不一致时，会先核实再继续。

## 用 uv 在终端中运行

通常直接让 Codex 使用技能即可。需要自己检查文档或运行某一步时，
也可以在仓库目录中使用以下命令。

查看命令和参数：

```powershell
uv run scripts/endnote_workflow.py --help
```

检查 Word 中的插件连接；这一步不会修改文档：

```powershell
uv run scripts/endnote_workflow.py probe
```

检查一份论文中已有的引用：

```powershell
uv run scripts/endnote_workflow.py inspect manuscript.docx --output work/inspection.json
```

对已经完成文献匹配、写入临时引文的新版本，调用插件生成正式引用：

```powershell
uv run scripts/endnote_workflow.py format manuscript-next.docx --style 'C:/path/to/Style.ens'
uv run scripts/endnote_workflow.py verify manuscript-next.docx
```

`format` 需要 Windows、Word、EndNote 和已打开的目标文献库。
它处理的是已经准备好的论文，不会自动判断原稿中每一处手写编号的含义。
从整理文献到生成引用的完整步骤和文件示例见[详细使用说明](references/workflow.md)。

`.python-version` 默认选择 Python 3.12，也支持 Python 3.10 及以上版本。
已有 Python 环境时，可以用 `uv sync --locked --python <Python路径>` 选择它。
之后运行命令时，也在 `uv run` 后加上 `--python <Python路径>`。

## 常见问题

**插入后还能在 Word 中继续编辑吗？**

可以。引用由已安装的 EndNote 插件生成，之后仍可用插件添加文献和更新格式。

**可以使用自己的期刊格式吗？**

可以。提供已安装的 `.ens` 文件即可，例如自己调整过的 Nature 格式。

**支持书籍、会议论文等文献吗？**

随附工具的自动整理和匹配目前用于期刊论文。书籍和会议论文请先在 EndNote 中
整理，再用 Word 插件插入。

**插件没有连接，或者引用没有更新怎么办？**

先运行 `probe` 检查连接，再查看[问题处理](references/troubleshooting.md)。

## 许可

本项目采用 [MIT 许可](LICENSE)。Word、EndNote 和期刊格式文件需自行准备，
不包含在本仓库中。
