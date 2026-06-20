import io
from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# Agence VU brand colors
COLOR_DARK_BLUE = RGBColor(0x1A, 0x2E, 0x4A)   # Dark navy — headers
COLOR_ACCENT = RGBColor(0x00, 0x8B, 0xD2)       # VU blue — accents
COLOR_LIGHT_GREY = RGBColor(0xF4, 0xF6, 0xF8)   # Background panels
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXT = RGBColor(0x2C, 0x3E, 0x50)         # Body text
COLOR_GOOD = RGBColor(0x27, 0xAE, 0x60)         # Green — bon
COLOR_MEDIUM = RGBColor(0xF3, 0x9C, 0x12)       # Orange — moyen
COLOR_BAD = RGBColor(0xE7, 0x4C, 0x3C)          # Red — faible

# Slide dimensions — 16:9 widescreen
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)


class PPTXAssembler:
    """Assembles PowerPoint presentations from generated slide content."""

    def __init__(self, template_path: Optional[str] = None):
        """
        Initialize assembler.

        Args:
            template_path: Optional path to a .pptx template file.
                           If None, creates a blank presentation.
        """
        self.template_path = template_path

    def _create_presentation(self) -> Presentation:
        """Create a new Presentation, from template if provided."""
        if self.template_path:
            try:
                prs = Presentation(self.template_path)
                return prs
            except Exception:
                pass
        prs = Presentation()
        prs.slide_width = SLIDE_WIDTH
        prs.slide_height = SLIDE_HEIGHT
        return prs

    def _add_cover_slide(self, prs: Presentation, pharmacy_name: str) -> None:
        """Add a cover/title slide for the Performance Globale section."""
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)

        # Background rectangle
        bg = slide.shapes.add_shape(
            1,  # MSO_SHAPE_TYPE.RECTANGLE
            Inches(0), Inches(0),
            SLIDE_WIDTH, SLIDE_HEIGHT,
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_DARK_BLUE
        bg.line.fill.background()

        # Accent bar
        bar = slide.shapes.add_shape(
            1,
            Inches(0), Inches(5.8),
            Inches(4), Inches(0.15),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = COLOR_ACCENT
        bar.line.fill.background()

        # Section label
        section_box = slide.shapes.add_textbox(
            Inches(1), Inches(2.5), Inches(11), Inches(0.8)
        )
        tf = section_box.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = "PARTIE 2 — PERFORMANCE GLOBALE"
        run.font.size = Pt(16)
        run.font.color.rgb = COLOR_ACCENT
        run.font.bold = True
        run.font.name = "Calibri"

        # Main title
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(3.1), Inches(11), Inches(1.4)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"Audit 360° — {pharmacy_name}" if pharmacy_name else "Audit 360°"
        run.font.size = Pt(36)
        run.font.color.rgb = COLOR_WHITE
        run.font.bold = True
        run.font.name = "Calibri"

        # Subtitle
        sub_box = slide.shapes.add_textbox(
            Inches(1), Inches(4.6), Inches(11), Inches(0.6)
        )
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "Analyse des indicateurs de performance — Agence VU"
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0xB0, 0xC4, 0xDE)
        run.font.name = "Calibri"

    def _add_content_slide(
        self,
        prs: Presentation,
        slide_number: int,
        total_slides: int,
        titre: str,
        contenu: str,
        chiffres_cites: list,
        sources: list,
        slide_id: str = "",
    ) -> None:
        """Add a content slide with title, body text, and metadata."""
        slide_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(slide_layout)

        # Header band
        header = slide.shapes.add_shape(
            1, Inches(0), Inches(0), SLIDE_WIDTH, Inches(1.3)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = COLOR_DARK_BLUE
        header.line.fill.background()

        # Accent line under header
        accent_line = slide.shapes.add_shape(
            1, Inches(0), Inches(1.3), SLIDE_WIDTH, Inches(0.05)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = COLOR_ACCENT
        accent_line.line.fill.background()

        # Slide title in header
        title_box = slide.shapes.add_textbox(
            Inches(0.4), Inches(0.2), Inches(11.5), Inches(0.9)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = titre
        run.font.size = Pt(22)
        run.font.color.rgb = COLOR_WHITE
        run.font.bold = True
        run.font.name = "Calibri"

        # Slide number in header (top right)
        num_box = slide.shapes.add_textbox(
            Inches(12.0), Inches(0.3), Inches(1.1), Inches(0.6)
        )
        tf = num_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = f"{slide_number}/{total_slides}"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0xB0, 0xC4, 0xDE)
        run.font.name = "Calibri"

        # Body content area
        body_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.55), Inches(8.5), Inches(5.0)
        )
        tf = body_box.text_frame
        tf.word_wrap = True

        # Split content into paragraphs
        paragraphs = [p.strip() for p in contenu.split("\n") if p.strip()]
        for i, para_text in enumerate(paragraphs):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.space_before = Pt(8)
            run = p.add_run()
            run.text = para_text
            run.font.size = Pt(14)
            run.font.color.rgb = COLOR_TEXT
            run.font.name = "Calibri"

        # Right panel: key figures
        if chiffres_cites:
            panel = slide.shapes.add_shape(
                1, Inches(9.3), Inches(1.55), Inches(3.7), Inches(5.0)
            )
            panel.fill.solid()
            panel.fill.fore_color.rgb = COLOR_LIGHT_GREY
            panel.line.color.rgb = RGBColor(0xD0, 0xD8, 0xE0)

            panel_title = slide.shapes.add_textbox(
                Inches(9.5), Inches(1.75), Inches(3.3), Inches(0.5)
            )
            tf = panel_title.text_frame
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = "Chiffres clés"
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = COLOR_DARK_BLUE
            run.font.name = "Calibri"

            # List key figures
            figures_box = slide.shapes.add_textbox(
                Inches(9.5), Inches(2.25), Inches(3.3), Inches(4.0)
            )
            tf = figures_box.text_frame
            tf.word_wrap = True

            for i, fig in enumerate(chiffres_cites[:8]):  # max 8 figures
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.space_before = Pt(6)
                run = p.add_run()
                # Format number nicely
                try:
                    num = float(fig)
                    if num >= 1_000_000:
                        display = f"{num/1_000_000:.2f} M"
                    elif num >= 1_000:
                        display = f"{num:,.0f}".replace(",", " ")
                    elif num == int(num):
                        display = str(int(num))
                    else:
                        display = f"{num:.2f}"
                    run.text = f"• {display}"
                except (TypeError, ValueError):
                    run.text = f"• {fig}"
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = COLOR_ACCENT
                run.font.name = "Calibri"

        # Sources footer
        if sources:
            src_text = "Sources: " + " | ".join(str(s) for s in sources[:4])
            src_box = slide.shapes.add_textbox(
                Inches(0.4), Inches(6.9), Inches(12.5), Inches(0.4)
            )
            tf = src_box.text_frame
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = src_text
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0x95, 0xA5, 0xA6)
            run.font.name = "Calibri"
            run.font.italic = True

        # Bottom accent line
        bottom_line = slide.shapes.add_shape(
            1, Inches(0), Inches(7.35), SLIDE_WIDTH, Inches(0.05)
        )
        bottom_line.fill.solid()
        bottom_line.fill.fore_color.rgb = COLOR_ACCENT
        bottom_line.line.fill.background()

    def _add_error_slide(
        self, prs: Presentation, slide_id: str, error_message: str
    ) -> None:
        """Add an error placeholder slide when generation failed."""
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)

        header = slide.shapes.add_shape(
            1, Inches(0), Inches(0), SLIDE_WIDTH, Inches(1.3)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = COLOR_BAD
        header.line.fill.background()

        title_box = slide.shapes.add_textbox(
            Inches(0.4), Inches(0.2), Inches(12.5), Inches(0.9)
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"[ERREUR] {slide_id}"
        run.font.size = Pt(20)
        run.font.color.rgb = COLOR_WHITE
        run.font.bold = True
        run.font.name = "Calibri"

        body_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0)
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"Erreur lors de la génération de cette diapositive:\n{error_message}"
        run.font.size = Pt(14)
        run.font.color.rgb = COLOR_TEXT
        run.font.name = "Calibri"

    def build(
        self,
        slides_content: list,
        pharmacy_name: str = "",
    ) -> bytes:
        """
        Build a complete PPTX from generated slide content.

        Args:
            slides_content: List of slide dicts from LLMGenerator.generate_performance_globale().
                            Each dict: {slide_id, titre, contenu, chiffres_cites, sources, erreur}
            pharmacy_name: Name of the pharmacy for the cover slide.

        Returns:
            Raw bytes of the .pptx file, ready for Streamlit download_button.
        """
        prs = self._create_presentation()

        # Add cover slide
        self._add_cover_slide(prs, pharmacy_name)

        # Count valid content slides
        total_content = len(slides_content)

        # Add content slides
        for idx, slide_data in enumerate(slides_content, start=1):
            erreur = slide_data.get("erreur")
            if erreur:
                self._add_error_slide(
                    prs,
                    slide_id=slide_data.get("slide_id", f"slide_{idx}"),
                    error_message=erreur,
                )
            else:
                self._add_content_slide(
                    prs=prs,
                    slide_number=idx,
                    total_slides=total_content,
                    titre=slide_data.get("titre", f"Slide {idx}"),
                    contenu=slide_data.get("contenu", ""),
                    chiffres_cites=slide_data.get("chiffres_cites", []),
                    sources=slide_data.get("sources", []),
                    slide_id=slide_data.get("slide_id", ""),
                )

        # Serialize to bytes
        buffer = io.BytesIO()
        prs.save(buffer)
        buffer.seek(0)
        return buffer.read()
