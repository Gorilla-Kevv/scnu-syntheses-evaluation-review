# python-docx 批注与单元格锚定 实现要点

> 用于在登记表副本上生成**原生 Word 批注**（显示在页边讨论区，可在 Word/WPS 中逐条查看）。
> 环境要求：**python-docx ≥ 1.2.0**（1.1.x 及更早**不支持** `add_comment`）。

## 一、核心 API

```python
from docx import Document

doc = Document(path)
# 只此一种写法；Paragraph.add_comment / Run.add_comment 均不存在
doc.add_comment(runs, text="批注内容", author="综测审核组", initials="SJ")
doc.save(dst)
```

- `runs` 可以是**单个 `Run`**，也可以是 **`Run` 序列**（传 `paragraph.runs` 则批注覆盖整段文本）。
- 写入结果：新增 `word/comments.xml`，并在 `document.xml` 中插入 `w:commentRangeStart` / `w:commentRangeEnd` 锚点。

**校验批注是否真的写入**：

```python
import zipfile
z = zipfile.ZipFile(dst)
assert "word/comments.xml" in z.namelist()
docxml = z.read("word/document.xml").decode("utf-8")
print("commentRangeStart:", docxml.count("commentRangeStart"))  # 应等于批注条数
```

## 二、合并单元格去重（关键坑）

python-docx 的 `row.cells` 对横向合并的单元格会**重复返回同一个 `tc`**。
直接按索引取 label/value 会错位。必须先按 `id(cell._tc)` 去重：

```python
def row_cells(row):
    seen, out = set(), []
    for c in row.cells:
        k = id(c._tc)
        if k in seen:
            continue
        seen.add(k)
        out.append(c)
    return out
```

去重后，登记表典型行的结构是 `[模块名, 标签1, 值1, 标签2, 值2, ...]`，即可按
"找标签 → 取下一个"的方式定位值单元格。

## 三、三类锚点定位器

登记表结构不统一，需要支持多种锚点。推荐如下签名：`anchor = (kind, ...)`

| kind | 形式 | 用途 |
|---|---|---|
| `hval` | `("hval", "年级")` | 定位**表头字段的值单元格**（第 0 行 label→value） |
| `hlbl` | `("hlbl", "姓名")` | 定位表头字段的**标签**单元格（值单元格为空时用） |
| `val` | `("val", 14, "劳动观念加分")` | 定位第 N 行的某个字段的**值单元格** |
| `lbl` | `("lbl", 14, "劳动观念加分")` | 定位第 N 行的某个字段的**标签**单元格 |
| `kw` | `("kw", "关键词")` | 遍历所有段落，定位**含关键词的段落**（表结构差异大时用） |

```python
def find_label(doc, row_idx, label):
    cells = row_cells(doc.tables[0].rows[row_idx])
    lab = label.replace(" ", "")
    for i, c in enumerate(cells):
        if c.text.strip().replace(" ", "") == lab:
            return c, (cells[i + 1] if i + 1 < len(cells) else None)
    return None, None
```

**遍历范围必须包含表格内段落**（否则定位不到单元格文字）：

```python
def all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    yield p
```

**关键词匹配要做"忽略空白"回退** —— 部分登记表文字被逐字加了空格（如 "本 模 块 总 得 分"）：

```python
def find_kw_para(doc, kw):
    for p in all_paragraphs(doc):
        if kw in p.text:
            return p
    k = kw.replace(" ", "")
    for p in all_paragraphs(doc):
        if k in p.text.replace(" ", ""):
            return p
    return None
```

**取 run 列表**（值单元格可能为空段落、无 run，需兜底建空 run 以承载锚点）：

```python
def run_of(cell):
    for p in cell.paragraphs:
        if p.runs:
            return p.runs
    return [cell.paragraphs[0].add_run("")]

def run_of_para(p):
    if p.runs:
        return p.runs
    return [p.add_run("")]
```

## 四、定位策略选择（经验）

1. **优先用 `('val', 行号, 标签)`** —— 最稳、语义最清晰。
   - 先用 `scripts/dump_docx.py` dump 出行号与字段名，再写锚点。
   - 值单元格**为空**时定位器自动回退到标签单元格（保证锚点可见）。
2. **`('kw', 关键词)` 仅在下列情况使用**：
   - 表结构与学生登记表差异较大（如"入团申请自荐表"，每格是独立整格文本）；
   - 需锚定明细行中的某一句话（如"作为队员参加寒假社会实践专项调研比赛"）。
3. **`('kw', ...)` 的失败率高** —— 关键词可能出现在整段合并的大单元格里而不独立成段。
   写之前**务必先 dump 验证关键词确实存在于某个段落**。
   本次实践中，`('kw','加分类')` 与 `('kw','加分：无')` 均需调整；前者改为 `('val', 8, '体质测试加分')` 才成功。
4. **总述批注**统一锚定 `("hval", "姓名")`，失败则回退 `("lbl", 0, "姓名")`。

## 五、总述批注里必须写"不可读材料清单"

批注总述除列问题总数外，还须列出**本文件夹内无文字层的扫描件/图片 + 原因**，
以满足"不可读文档也要说明原因及位置"的要求。检测函数：

```python
def unreadable_list(folder):
    res = []
    for dp, dn, fn in os.walk(folder):
        for f in sorted(fn):
            p = os.path.join(dp, f)
            if "审核批注版" in f or "审核修订版" in f:
                continue                     # 跳过自己的产出
            low = f.lower()
            if low.endswith((".jpg", ".jpeg", ".png", ".bmp", ".gif")):
                res.append((f, "图片文件（无文字层），无法直接提取文字核对"))
            elif low.endswith(".pdf"):
                import pymupdf
                d = pymupdf.open(p)
                t = "".join(pg.get_text() for pg in d)
                if len(t.strip()) < 20:      # 阈值：低于 20 字符视为扫描件
                    res.append((f, "扫描版PDF（无文字层），无法直接提取文字核对"))
    return res
```

## 六、安全约束

```python
import shutil, os
dst = os.path.join(os.path.dirname(src), f"{name}（审核批注版）.docx")
shutil.copy2(src, dst)      # 只操作副本；原件不写、不移动
doc = Document(dst)
...
doc.save(dst)
```

- **永远 `copy2` 出副本再改**，绝不 `Document(src).save(src)`。
- 副本落在**原文件同级目录**（用户明确要求，勿集中到新文件夹）。
- 完成后用 `ls -la` 比对原件 `mtime`，作为"原件未改动"的证据。

## 七、其他格式的处理

- **PDF 有文字层**：`pymupdf` 提取全文后直接检索姓名。
- **PDF 无文字层 / 图片**：`page.get_pixmap(dpi=150).save(png)` 逐页渲染，再视觉识别。
  几十页的大名单必须逐页看，本人可能在很靠后的页。
- **xls**：`xlrd`（老格式）/ **xlsx**：`openpyxl`。
- 若系统未装 OCR（tesseract 等），**逐页渲染 + 视觉识别**是唯一可靠路径，不要因"无 OCR"而跳过。
