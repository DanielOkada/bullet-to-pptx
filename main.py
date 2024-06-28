from pptx import Presentation
from pptx.enum.lang import MSO_LANGUAGE_ID


def parse_hierarchical_text(text):
    lines = text.strip().split('\n')
    result = []
    stack = [(-1, result)]

    for line in lines:
        indent = len(line) - len(line.lstrip())
        content = line.strip('- ').strip()

        while stack and indent <= stack[-1][0]:
            stack.pop()

        if not stack:
            stack = [(-1, result)]

        parent = stack[-1][1]
        current = []
        parent.append((content, current))
        stack.append((indent, current))

    return result


with open("input.txt", encoding="utf-8") as f:
    text = f.read()

# パース実行
hierarchy = parse_hierarchical_text(text)

prs = Presentation("template.pptx")


def add_body(shapes, items, level=0):
    for item, sub_items in items:
        # 段落を追加し、インデントを設定
        p = shapes.placeholders[1].text_frame.paragraphs[0]
        if p.text != "":
            p = shapes.placeholders[1].text_frame.add_paragraph()

        p.text = item
        p.level = level

        add_body(shapes, sub_items, level + 1)


for item, sub_items in hierarchy:
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    slide.shapes.title.text = item

    add_body(slide.shapes, sub_items)

    # 言語を日本語に設定
    for paragraph in slide.shapes.placeholders[1].text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.language_id = MSO_LANGUAGE_ID.JAPANESE


file_name = "out/example.pptx"
prs.save(file_name)
print(file_name + "を作成")
print("スライド番号を忘れずに")