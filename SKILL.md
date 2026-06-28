---
name: ppt-analysis
description: PDF/PPT课件自动分析工具——将PDF按页切割为高清图片，AI逐张分析PPT内容并生成详细中文讲解，提取英文专业名词解释，最终输出格式化的Word文档（微软雅黑字体）。支持PDF和PPTX文件。
---

# ppt-analysis

当用户提供PDF或PPT文件并要求分析、解释课件内容时，使用此技能。

## 触发条件

用户消息中包含以下任一意图：
- 分析/解释PDF课件、PPT内容
- 将PDF/PPT转为Word讲解文档
- 提取PDF/PPT中的英文名词解释
- 例如："帮我分析这个PDF" / "解释这张PPT" / "生成Word讲解文档"
- 支持的文件格式：`.pdf`、`.pptx`

## AI 角色定义

**本技能要求AI以以下身份进行工作：**

### 讲解角色：专业教师

AI必须以**专业教师**的身份，对每一页PPT进行详细讲解，要求：

1. **教学口吻**：以老师上课的口吻讲解内容，语言自然流畅，不要机械化的模板套用
2. **逻辑清晰**：讲解要有层次和逻辑，从背景引入→核心内容→深入分析→总结归纳，根据每页内容灵活组织
3. **深度讲解**：不是简单翻译PPT文字，而是对内容进行深入解读，解释原理、分析机制、阐明意义
4. **灵活处理**：根据每页PPT的实际内容类型（纯文字页、机制图页、数据页、图表页等）灵活调整讲解方式和侧重点
5. **重点突出**：对关键概念、重要机制、易错点进行重点强调和提醒

### 术语筛选角色：领域专家

在进行英文名词解释汇总时，AI必须以**该PPT涉及的领域专家**的身份，对术语进行专业筛选和分级标注：

1. **专业筛选**：不是所有出现的英文词都需要解释，专家应判断哪些是真正的专业术语
2. **分级标注**：
   - **★★★ 必掌握**：课程核心术语、考试高频考点、理解本章内容不可或缺的概念
   - **★★ 重要**：帮助深入理解课程内容的重要术语
   - **★ 了解**：出现频率较低、属于拓展性知识的术语
3. **专业解释**：每个术语的解释要准确、专业，包含定义及其在细胞生物学中的具体含义
4. **关联说明**：对核心术语（★★★）补充说明与其他概念的关联或易混淆概念的区分

## 完整工作流程

### Step 1: 安装依赖

在执行任何操作前，确保依赖已安装：

```bash
python -c "import fitz; import docx" 2>nul && echo "core_ok" || pip install PyMuPDF python-docx
python -c "import pptx" 2>nul && echo "pptx_ok" || pip install python-pptx
```

> 核心依赖：`PyMuPDF`、`python-docx`（必需）；`python-pptx`（PPT文本提取fallback用，建议安装）。

### Step 1.5: 文件类型检测与预处理

根据用户提供的文件路径，判断文件类型并进行预处理：

```python
import os
import subprocess

def detect_and_convert(file_path):
    """
    检测文件类型并预处理。
    返回 (处理后的文件路径, 模式)
    - 模式 'pdf': 返回PDF路径，后续走标准切图流程
    - 模式 'text': 返回提取的文本数据，跳过切图，直接用文本生成Word
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.pdf':
        return file_path, 'pdf'
    
    if ext == '.pptx':
        # 方法A: 尝试通过 PowerPoint COM 自动化转PDF（仅Windows+已安装PPT）
        pdf_path = try_convert_via_com(file_path)
        if pdf_path:
            return pdf_path, 'pdf'
        
        # 方法B: 尝试通过 LibreOffice 转PDF（跨平台）
        pdf_path = try_convert_via_libreoffice(file_path)
        if pdf_path:
            return pdf_path, 'pdf'
        
        # 方法C: fallback — 用 python-pptx 提取文本内容
        slides_data = extract_pptx_text(file_path)
        return slides_data, 'text'
    
    raise ValueError(f"不支持的文件格式: {ext}，请提供 .pdf 或 .pptx 文件")


def try_convert_via_com(pptx_path):
    """方法A: Windows PowerPoint COM 自动化转PDF"""
    try:
        import win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True
        abs_path = os.path.abspath(pptx_path)
        presentation = powerpoint.Presentations.Open(abs_path)
        pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
        presentation.SaveAs(pdf_path, FileFormat=32)  # 32 = ppPDF
        presentation.Close()
        powerpoint.Quit()
        print(f"[方法A-成功] 通过PowerPoint COM转换为PDF: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"[方法A-失败] PowerPoint COM不可用: {e}")
        return None


def try_convert_via_libreoffice(pptx_path):
    """方法B: LibreOffice 转PDF"""
    try:
        output_dir = os.path.dirname(os.path.abspath(pptx_path))
        # 尝试多个可能的命令名
        for cmd_name in ["soffice", "soffice.exe", r"C:\Program Files\LibreOffice\program\soffice.exe"]:
            try:
                result = subprocess.run(
                    [cmd_name, "--headless", "--convert-to", "pdf", "--outdir", output_dir, 
                     os.path.abspath(pptx_path)],
                    capture_output=True, text=True, timeout=60
                )
                if result.returncode == 0:
                    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
                    pdf_path = os.path.join(output_dir, pdf_name)
                    if os.path.exists(pdf_path):
                        print(f"[方法B-成功] 通过LibreOffice转换为PDF: {pdf_path}")
                        return pdf_path
            except FileNotFoundError:
                continue
        print("[方法B-失败] LibreOffice不可用")
        return None
    except Exception as e:
        print(f"[方法B-失败] LibreOffice转换出错: {e}")
        return None


def extract_pptx_text(pptx_path):
    """方法C: 用 python-pptx 提取所有幻灯片文本内容（fallback）"""
    from pptx import Presentation
    prs = Presentation(pptx_path)
    slides_data = []
    for i, slide in enumerate(prs.slides):
        slide_text = []
        # 提取所有文本框内容
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text.strip())
            # 提取表格内容
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells]
                    slide_text.append(" | ".join(row_text))
        # 提取备注
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text
            if notes_text.strip():
                slide_text.append(f"[备注] {notes_text.strip()}")
        
        slides_data.append({
            "slide_number": i + 1,
            "text": "\n".join(slide_text) if slide_text else "(空白页)"
        })
    
    print(f"[方法C-成功] 通过python-pptx提取了 {len(slides_data)} 页幻灯片文本")
    return slides_data
```

**执行逻辑：**
- 如果是 `.pdf` → 直接进入 Step 2 切图流程
- 如果是 `.pptx` → 按优先级尝试 A → B → C
  - 方法A/B 转换成功 → 得到 PDF，进入 Step 2 切图流程
  - 方法C（fallback）→ 提取文本，跳过 Step 2 和 Step 3，直接进入增强版 Step 4

### Step 2: PDF切图

使用 PyMuPDF 将PDF的每一页渲染为高清PNG图片（2倍缩放以保证清晰度），输出到 `pdf_pages/` 文件夹。

```python
import fitz  # PyMuPDF
import os

def extract_pdf_to_images(pdf_path, output_dir="pdf_pages"):
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    image_paths = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = fitz.Matrix(2, 2)  # 2x缩放
        pix = page.get_pixmap(matrix=mat)
        output_path = os.path.join(output_dir, f"page_{page_num + 1:02d}.png")
        pix.save(output_path)
        image_paths.append(output_path)
    doc.close()
    return image_paths
```

### Step 3: AI逐张分析图片（教师角色讲解）

对每一张切割出的PNG图片，AI以**细胞生物学专业教师**的身份进行分析和讲解：

1. 使用 `read_file` 工具打开并查看图片
2. 利用AI视觉能力识别PPT中的所有内容（文字、图表、示意图、机制图等）
3. 根据页面内容类型，以教师角色生成以下内容：

#### 每页输出结构

**a) 页面标题**
- 根据PPT内容拟定简洁准确的中文标题

**b) PPT原图**
- 将该页的PNG图片嵌入到文档中（图片紧跟标题之后）

**c) 教师详细讲解**
- 以专业教师的口吻对本页内容进行详细讲解
- 讲解要有逻辑、有层次，深入解读PPT内容
- **根据页面内容灵活组织讲解结构**，可包含但不限于：
  - 本页核心内容概述
  - 对文字内容的深入解读和分析
  - 对图表/示意图/机制图的详细解读（遇到机制图、流程图、示意图时应重点围绕图片进行讲解，详细说明图中各部分的关系、流程步骤、作用机制等）
  - 重点概念的强调
  - 易错点或易混淆概念的提醒（如有）
  - 知识拓展或前后知识关联（如有）
- 讲解文字中自然地引用图片内容，根据实际情况灵活处理，不强制使用固定句式

**d) 本页英文专业名词**
- 从图片中识别出现的英文专业术语
- AI直接根据生物学知识给出每个术语的中文解释
- 每个术语格式: `英文术语 —— 中文解释`

#### Step 3 与 Step 4 的分支

- **模式 `pdf`**（PDF文件或PPT成功转为PDF后）：执行完整的 Step 2 → Step 3 → Step 4 流程
- **模式 `text`**（PPT文本提取fallback）：跳过 Step 2 和 Step 3，直接用提取的文本数据进入 Step 4

### Step 4: 生成Word文档

使用 `write_to_file` 生成完整的Python脚本，用 `python-docx` 创建Word文档。

#### 模式 `pdf` 的文档结构（标准流程）：

1. **封面页**
   - 标题（如：文件名 + "内容详解"）
   - 总页数信息

2. **逐页详解**（每页一个章节）
   - 页面标题（Heading 1）
   - PPT原图嵌入（宽度5.5英寸）——**必须包含图片**
   - 教师详细讲解段落
   - 本页英文术语表（项目符号列表）

3. **附录：英文名词解释汇总**
   - 按字母排序的表格（序号 | 优先级 | 英文术语 | 中文解释 | 专家补充说明）
   - 表格前显示统计信息："本课件共收录X个专业术语，其中★★★必掌握X个，★★重要X个，★了解X个"

#### 模式 `text` 的文档结构（PPT文本fallback）：

1. **封面页**
   - 标题（如：文件名 + "内容详解"）
   - 总页数信息
   - 说明："本文件由PPT文本内容提取生成"

2. **逐页详解**（每页一个章节）
   - 页面标题（Heading 1）
   - 原始文本内容（引用格式，灰色背景）
   - AI基于文本生成的详细教师讲解
   - 本页英文术语表

3. **附录：英文名词解释汇总**
   - 按字母排序的表格（序号 | 优先级 | 英文术语 | 中文解释 | 专家补充说明）
   - 统计信息

#### 模式 `text` 的分析说明

当使用文本提取模式时，AI分析基于提取的文字内容（而非图片），需要：
- 根据文本内容推测可能的图表/示意图并描述
- 更加详细地展开文字内容的解释
- 在文档中标注"本页内容基于文本提取，无原始图片"

#### 英文名词解释汇总的生成规则

生成汇总表时，AI必须以**该PPT涉及领域的专家**身份进行工作：

1. **术语筛选**：遍历所有页面提取的术语，去除重复项，由专家判断每个术语的重要程度
2. **分级标注**：
   - **★★★ 必掌握**：课程核心术语、考试高频考点、理解本章不可或缺的概念
   - **★★ 重要**：帮助深入理解的重要术语
   - **★ 了解**：出现频率低、属于拓展性的术语
3. **专业解释**：每个术语给出准确专业的中文解释
4. **专家补充**：对★★★术语补充说明易混淆概念区分、与其他术语的关联等

### Step 5: 字体与格式控制（严格要求）

**这是最关键的步骤——必须杜绝任何意外加粗。**

#### 5.1 覆盖所有内置样式

```python
def setup_styles(doc):
    # Normal样式
    style = doc.styles['Normal']
    style.font.name = 'Microsoft YaHei'
    style.font.size = Pt(11)
    style.font.bold = False
    # 必须设置East Asian字体
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

    # 覆盖所有Heading样式
    for i in range(0, 10):
        try:
            hs = doc.styles[f'Heading {i}']
            hs.font.name = 'Microsoft YaHei'
            hs.font.bold = False  # 关键：关闭默认加粗
            hs.font.size = Pt({0: 22, 1: 18, 2: 15, 3: 13}.get(i, 13))
            rPr_h = hs.element.get_or_add_rPr()
            rFonts_h = rPr_h.find(qn('w:rFonts'))
            if rFonts_h is None:
                rFonts_h = OxmlElement('w:rFonts')
                rPr_h.insert(0, rFonts_h)
            rFonts_h.set(qn('w:eastAsia'), 'Microsoft YaHei')
        except KeyError:
            pass
```

#### 5.2 显式设置每个Run的字体

**绝不使用 `p.add_run()` 裸调用后设置格式。必须通过 `set_run_font()` 统一处理。**

```python
def set_run_font(run, font_name='Microsoft YaHei', size=None, bold=False, italic=False, color=None):
    run.bold = bold
    run.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    if size:
        run.font.size = size
    if color:
        run.font.color.rgb = color

def make_run(paragraph, text, bold=False, italic=False, color=None, size=None):
    run = paragraph.add_run(text)
    set_run_font(run, bold=bold, italic=italic, color=color, size=size)
    return run
```

#### 5.3 标题创建后强制去除加粗

```python
def add_heading(doc, text, level=1, color=None):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        set_run_font(run, bold=False, color=color)
        run.font.size = Pt({0: 22, 1: 18, 2: 15, 3: 13}.get(level, 13))
    return heading
```

#### 5.4 字体加粗的使用原则

| 场景 | bold值 |
|------|--------|
| 标题（Heading） | False |
| 正文段落 | False |
| "核心论点："、"关键意义："、"图片说明："等标签 | True |
| 术语英文名 | True |
| 术语中文解释 | False |
| 表格表头 | False |
| 表格内容 | False |

### Step 6: 执行脚本并输出

1. 将生成的Python脚本通过 `write_to_file` 保存为 `.py` 文件
2. 通过 `execute_command` 运行脚本
3. 确认Word文件生成成功
4. 告知用户输出文件路径

### Step 7: 清理临时文件

Word文档生成成功后，**必须**删除 `pdf_pages/` 目录下的所有临时PNG图片，保持工作目录整洁。原始PDF/PPT文件不受影响。

在生成的Python脚本中，确认Word文件已成功保存后，调用以下清理代码：

```python
import shutil
import os

def cleanup_images(output_dir="pdf_pages"):
    """删除切分出的临时图片，保留原始PDF"""
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
        print(f"已清理临时图片目录: {output_dir}")
```

在脚本末尾、确认Word生成成功后调用 `cleanup_images()`。


## 输出规范

- 文件格式：`.docx`
- 文件命名：`<原文件名>_详解.docx`
- 字体：全文微软雅黑（Microsoft YaHei）
- 不使用任何未明确指定的格式

## 注意事项

1. 每次只处理一个文件（支持 `.pdf` 和 `.pptx`）
2. PDF模式：图片使用2倍缩放渲染，保证PPT中的文字清晰可读
3. PPT模式：自动按优先级尝试三种转换方式（PowerPoint COM → LibreOffice → python-pptx文本提取）
4. 如果PDF页数很多（>30页），建议分批处理
5. 所有中文内容使用微软雅黑，英文内容同样使用微软雅黑
6. 任务完成后必须清理 `pdf_pages/` 目录下的临时图片，只保留原始文件
7. 绝不允许出现"字体莫名加粗"的问题——每个Run必须显式设置 `bold` 属性
8. 新用户首次使用时，只需安装 `pip install PyMuPDF python-docx python-pptx`，无需安装 LibreOffice
9. 如果系统安装了 PowerPoint（Windows）或 LibreOffice，PPT转PDF效果最佳；否则自动fallback到文本提取模式
10. 每页的讲解必须包含PPT原图（图片+文字配合输出），text模式除外
11. 讲解以教师角色进行，内容详细有逻辑，语言自然灵活不模板化
12. 机制图、流程图、示意图等视觉内容要重点围绕图片进行详细解读
13. 英文名词汇总必须由专家角色进行分级标注（★★★/★★/★）