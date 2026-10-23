# 排版、导出与质量检查

## 环境与依赖

Python 3.10+。优先使用执行环境提供的文档运行时。在 Codex 桌面可通过 `load_workspace_dependencies` 获取 Python 和包路径；不要把本机的绝对路径写进技能默认配置。

普通 Python 环境缺少依赖时，在用户认可的环境中运行：

```bash
python -m pip install -r scripts/requirements.txt
```

不要自动更改系统 Python 或安装整个办公套件。生成器的依赖是 `python-docx`、`reportlab`、`pypdf`；`PyYAML` 仅在验证技能元数据时需要。

## 默认排版

- 默认 `reference-blue` 采用 [参考格式规范](reference-style.md)：A4 单栏、左右 12 mm、上下 10 mm；姓名 22 pt、栏目标题 14 pt、正文 10.5 pt、1.35 倍行距；深蓝 `#244A68` 名称与栏目、细蓝分隔线、页首尾装饰线、正文深灰。`ats` 保留旧的黑白简洁版式，仅在用户要求时选择。
- 中文默认微软雅黑，西文默认 Arial；PDF 必须使用可嵌入且覆盖正文字符的 TrueType 字体。也支持用户指定字体。
- 用真实段落、标题样式和列表编号组织 DOCX；正文按自然阅读顺序排列，不使用图片文字、多栏布局、正文表格或文本框。
- URL 在 DOCX、PDF 中保持完整可见并可点击。不替用户访问作品集，也不创建未确认的外部链接。除非用户提供，不添加照片和二维码。
- 栏目标题与下一段保持同页，经历标题与第一条正文尽量相连；长经历允许在条目间跨页，避免整体强制锁页留下大块空白。
- 日期、项目／公司名、角色／地点对齐在一行；长名称自动改为自然换行。经历正文支持短段落和贡献列表，关键词用真实字体局部加粗。
- 保持编辑性。Word 因本机字体替换可能重排，所以“PDF 检查合格”不能代替 DOCX 预览。

脚本 `--body-size` 允许 10–12 pt；默认页数只告警，不自动删内容或无限缩小字号。用户模板和明确要求优先于默认外观。

## 常用命令

以下命令假设当前目录是技能目录；实际使用可改为脚本绝对路径。

```bash
# 先查看内容深度缺口，无文件输出
python scripts/build_resume.py --input profile.json --output-dir output --review-only

# 默认参考蓝色版；有可用 LibreOffice 时从 DOCX 导出，否则独立排版 PDF 并告警
python scripts/build_resume.py --input profile.json --output-dir output --basename 姓名_岗位_简历 --layout reference-blue

# 用户明确选择事实受限稿时才绕过过短经历检查
python scripts/build_resume.py --input profile.json --output-dir output --basename 简历_事实受限稿 --allow-sparse

# 用户要求纯文本友好的黑白版时
python scripts/build_resume.py --input profile.json --output-dir output --basename resume --layout ats

# 可用转换器；在提供捆绑运行时的环境里选择其转换器
python scripts/build_resume.py --input profile.json --output-dir output --basename resume --pdf-engine libreoffice --soffice /absolute/path/to/soffice

# 无办公转换器时直接生成 PDF；指定覆盖正文的字体文件
python scripts/build_resume.py --input profile.json --output-dir output --basename resume --pdf-engine reportlab --pdf-font /absolute/path/to/font.ttf --font-name 用户指定的Word字体

# 两页目标，只改变告警标准；不会自动截断内容
python scripts/build_resume.py --input profile.json --output-dir output --basename resume --max-pages 2

# 修改后覆盖自己已生成的文件，需要明确加该参数
python scripts/build_resume.py --input profile.json --output-dir output --basename resume --overwrite

# 文本一致性检查
python scripts/verify_resume.py --docx output/resume.docx --pdf output/resume.pdf
```

`--pdf-font` 支持 TrueType `.ttf` 和 TrueType `.ttc`；集合默认使用第 0 个子字体，可通过 `--font-index` 指定。基于 CFF 的 OpenType 文件不受此备用 PDF 后端支持。中文字体自动搜索 Windows、macOS 和常见 Linux 字体位置；搜索失败或字符覆盖不足时明确报错，不输出方框字。

生成器先检查核心经历深度，再在输出目录的临时文件夹生成两份文件和校验报告，最后移动到目标路径；转换失败不会留下声称完成的最终 PDF。不覆盖同名文件，除非传入 `--overwrite`。报告为 `<basename>.qa.json`，记录版式、内容深度信号、PDF 后端、PDF 页数、DOCX/PDF 文本检查和自动警告；报告不含简历全文，不默认交付。内容深度信号以长度和条目数量排查明显缺口，不能检验个人贡献、真实性或空泛措辞，仍需人工审阅。

`auto` 中存在转换器但执行失败时直接报错，保留诊断信息，不悄悄改成独立排版。可根据实际原因修复，或告知用户后明确选择备用后端。需要 GUI 应用或写入受限目录时遵循当前环境权限，不绕过审批。

## Windows 文件可访问性

临时文件夹可能带有仅生成账号可访问的保护 ACL；同卷 `rename`／`replace` 会保留这个 ACL。沙箱能解析文档而用户的 WPS／Word 无法打开时，检查最终文件与输出目录的访问权限，不能据此直接判断文件损坏。

生成器先把已验证文件的字节复制到输出目录中新建的同级暂存文件，再发布，确保继承输出目录权限；不使用移动原临时文件或复制安全描述符的方式。修复旧成品时优先恢复其对既有输出目录的正常权限继承，保留内容和哈希；只修复受影响文件，不扩大整个目录的权限，不给 Everyone 追加完全控制。

## 视觉校验与备用路径

有文档技能或 `render_docx.py` 时按其要求渲染 DOCX。没有 DOCX 渲染器时，尝试执行现有渲染器获取诊断，保留结构检查结果，并说明“DOCX 已生成，未验证 Word 视觉分页”，不把独立生成 PDF 的图像当作 Word 预览。

用现有 PDF 渲染工具生成每页 PNG 并逐页查看，例如：

```bash
pdftoppm -r 150 -png output/resume.pdf output/preview
```

确认无乱码、缺字、截断、重叠、异常空白、标题孤页；页数符合用户要求。通过复制 PDF 正文或 `verify_resume.py` 确认可提取文本。链接需要同时查看 PDF 链接注释和 DOCX 超链接关系；不得把文本可读性说成可保证 ATS 兼容。

备用 PDF 与 DOCX 使用同一 JSON 和阅读顺序，但排版引擎不同。两份实际内容以文本一致性报告检查；视觉分页分别检查。若用户要求两份完全同版且缺少转换工具，这是工具缺口，应交付可完成的部分并说明，不能宣称完全一致。
