from pathlib import Path
from docx import Document    # 문서를 불러와 열거나 생성
from docx.enum.text import WD_LINE_SPACING    # 텍스트의 줄간격 조정
from docx.oxml.ns import qn    # 한글 글꼴 설정
from docx.shared import Mm, Pt, RGBColor    # RGBColor: 색상을 RGB로 설정
# Mm: 단위를 mm로 설정, Pt: 단위를 point로 설정
from docx.styles.style import ParagraphStyle    # 들여쓰기, 줄간격 등 문단 스타일
from docx.text.run import Run    # 문서에 텍스트 추가

from .api_data import OUT_DIR

def apply_font(arg: Run | ParagraphStyle,
               face: str="Malgun Gothic", size_pt: int=None,
               is_bold: bool=None, rgb: str=None):
    if face is not None:
        arg.font.name = face    #  폰트
        for prop in ["asciiTheme", "cstheme", "eastAsia", "eastAsiaTheme", "hAnsiTheme"]:
            arg.element.rPr.rFonts.set(qn(f"w:{prop}"), face)    # 한글

    
    if size_pt is not None:
        arg.font.size = Pt(size_pt)    # 폰트 크기
    
    if is_bold is not None:
        arg.font.bold = is_bold    # 폰트 굵기

    if rgb is not None:
        arg.font.color.rgb = RGBColor.from_string(rgb.upper())    # 폰트 색상


OUT3 = OUT_DIR / f"{Path(__file__).stem}.docx" 
def init_docx(OUT_DIR):
    OUT3 = OUT_DIR / f"{Path(__file__).stem}.docx"

    doc = Document()    # 객체 생성
    section = doc.sections[0]    # 첫번째 섹션
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = Mm(20)
    section.left_margin = section.right_margin = Mm(12.7)

    style = doc.styles["Normal"]     # 표준 단락 서식
    p_format = style.paragraph_format    # 단락 객체
    p_format.space_before = p_format.space_after = 0    # 단락 간격
    p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE    # 줄 간격

    apply_font(style, size_pt=10)     # 폰트 설정
    doc.save(OUT3)

if __name__ == "__main__":
    from api_data import OUT_DIR
    init_docx(OUT_DIR)
