import tempfile
from pathlib import Path

import docx
import fitz  # PyMuPDF
from docx.enum.section import WD_SECTION_START
from docx.shared import Inches, Pt
from win32com.client import DispatchEx

SUPPORTED_INPUT_EXTENSIONS = {
    ".doc",
    ".docx",
    ".docm",
    ".docb",
    ".dot",
    ".dotx",
    ".dotm",
    ".rtf",
    ".odt",
    ".txt",
    ".htm",
    ".html",
    ".mht",
    ".mhtml",
    ".xml",
}


def ensure_supported_extension(input_path: str) -> None:
    ext = Path(input_path).suffix.lower()
    if ext not in SUPPORTED_INPUT_EXTENSIONS:
        raise ValueError(
            f"Неподдерживаемое расширение: {ext}\n"
            f"Поддерживаются: {', '.join(sorted(SUPPORTED_INPUT_EXTENSIONS))}"
        )


def _default_margins_inches() -> dict:
    """Fallback if PageSetup cannot be read (matches typical Word 'Normal')."""
    return {
        "left": 1.0,
        "right": 1.0,
        "top": 1.0,
        "bottom": 1.0,
        "header": 0.5,
        "footer": 0.5,
    }


def export_to_pdf_with_word(input_path: str, output_pdf_path: str) -> dict:
    """Export via Word; returns page margins from the source document (inches)."""
    wdFormatPDF = 17

    word = None
    doc = None
    margins = _default_margins_inches()
    try:
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(Path(input_path).resolve()))
        ps = doc.PageSetup
        margins = {
            "left": float(ps.LeftMargin) / 72.0,
            "right": float(ps.RightMargin) / 72.0,
            "top": float(ps.TopMargin) / 72.0,
            "bottom": float(ps.BottomMargin) / 72.0,
            "header": float(ps.HeaderDistance) / 72.0,
            "footer": float(ps.FooterDistance) / 72.0,
        }
        doc.SaveAs(str(Path(output_pdf_path).resolve()), FileFormat=wdFormatPDF)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
    return margins


def configure_section(section, margins: dict) -> None:
    section.left_margin = Inches(margins["left"])
    section.right_margin = Inches(margins["right"])
    section.top_margin = Inches(margins["top"])
    section.bottom_margin = Inches(margins["bottom"])
    section.header_distance = Inches(margins["header"])
    section.footer_distance = Inches(margins["footer"])


def pdf_to_docx_images(pdf_path: str, output_docx_path: str, dpi: int, on_progress, margins: dict) -> None:
    pdf_doc = fitz.open(pdf_path)
    out_doc = docx.Document()
    configure_section(out_doc.sections[0], margins)

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            image_infos = []
            zoom = dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            total_pages = len(pdf_doc)

            for page_index in range(total_pages):
                page = pdf_doc[page_index]
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                img_path = Path(temp_dir) / f"page_{page_index + 1}.png"
                pix.save(str(img_path))
                image_infos.append((img_path, pix.width, pix.height))
                if on_progress:
                    on_progress("render", page_index + 1, total_pages)

            for i, (img_path, img_w_px, img_h_px) in enumerate(image_infos):
                if i > 0:
                    section = out_doc.add_section(WD_SECTION_START.NEW_PAGE)
                    configure_section(section, margins)
                else:
                    section = out_doc.sections[0]

                max_width_inches = section.page_width.inches - section.left_margin.inches - section.right_margin.inches
                max_height_inches = section.page_height.inches - section.top_margin.inches - section.bottom_margin.inches
                max_height_inches = max(0.5, max_height_inches - 0.02)

                img_aspect = img_w_px / img_h_px
                box_aspect = max_width_inches / max_height_inches

                paragraph = out_doc.add_paragraph()
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                run = paragraph.add_run()

                if img_aspect >= box_aspect:
                    run.add_picture(str(img_path), width=Inches(max_width_inches))
                else:
                    run.add_picture(str(img_path), height=Inches(max_height_inches))

                if on_progress:
                    on_progress("insert", i + 1, len(image_infos))

            out_doc.save(output_docx_path)
    finally:
        pdf_doc.close()


def next_output_path(input_path: str, output_dir: Path | None = None) -> Path:
    input_file = Path(input_path)
    directory = output_dir if output_dir is not None else input_file.parent
    candidate = directory / f"Converted_{input_file.stem}.docx"
    if not candidate.exists():
        return candidate
    return directory / f"Converted_{input_file.stem}1.docx"


def convert_document(input_path: str, on_progress=None, output_dir: Path | None = None) -> Path:
    ensure_supported_extension(input_path)
    output_path = next_output_path(input_path, output_dir=output_dir)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pdf = str(Path(temp_dir) / "intermediate.pdf")
        if on_progress:
            on_progress("status", "Экспорт документа в PDF через Microsoft Word...")
        margins = export_to_pdf_with_word(input_path, temp_pdf)
        if on_progress:
            on_progress("status", "Рендер страниц и сборка DOCX...")
        pdf_to_docx_images(temp_pdf, str(output_path), dpi=220, on_progress=on_progress, margins=margins)
    return output_path
