"""Installation sheet dialog for filling and exporting installation data."""

import os
import sys
import shutil
import tempfile
from datetime import date
from typing import Dict, List, Optional

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QScrollArea,
    QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog,
    QMessageBox, QGroupBox, QGridLayout, QSizePolicy, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap


# ---------------------------------------------------------------------------
# Field definitions: (label, excel_cell, multiline)
# ---------------------------------------------------------------------------
INSTALLATION_FIELDS: List[tuple] = [
    ("Client / Company Name",   "B21", False),
    ("Site / Project Name",     "B22", False),
    ("Installation Date",       "B23", False),
    ("Technician Name",         "B24", False),
    ("Vehicle Type / Model",    "B25", False),
    ("Vehicle Serial Number",   "B26", False),
    ("Engine Type",             "B27", False),
    ("CAN Channel",             "B28", False),
    ("Baud Rate (bps)",         "B29", False),
    ("Software Version",        "B30", False),
    ("Configuration Profile",   "B31", False),
    ("Notes / Comments",        "B32", True),
]

# Path to the bundled Excel template (relative to project root)
_TEMPLATE_FILENAME = "installation_sheet_template.xlsx"


def _template_path() -> str:
    """Return the absolute path to the Excel template, searching sensible locations."""
    candidates = [
        os.path.join(os.path.dirname(__file__), "..", "assets", _TEMPLATE_FILENAME),
        os.path.join(os.path.dirname(__file__), "..", _TEMPLATE_FILENAME),
        os.path.join(getattr(sys, "_MEIPASS", ""), "assets", _TEMPLATE_FILENAME),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return os.path.abspath(p)
    # If not found, return the expected location (it will be created on first use)
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "assets", _TEMPLATE_FILENAME)
    )


def _create_template(path: str) -> None:
    """Create the Excel template at *path* if it does not exist yet."""
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return  # silently skip — will surface properly in _fill_excel

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Installation Sheet"

    # ----- Column widths -----
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 45

    # ----- Title rows (1-18 reserved for header / logo) -----
    ws.merge_cells("A1:B1")
    title_cell = ws["A1"]
    title_cell.value = "SARTEL – Installation Sheet"
    title_cell.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="1F497D")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:B2")
    sub_cell = ws["A2"]
    sub_cell.value = "CAN Bus Monitoring – Field Installation Record"
    sub_cell.font = Font(name="Calibri", italic=True, size=11, color="1F497D")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    sub_cell.fill = PatternFill("solid", fgColor="DCE6F1")
    ws.row_dimensions[2].height = 20

    # Spacer rows 3-19
    for r in range(3, 20):
        ws.row_dimensions[r].height = 5

    # ----- Column headers (row 20) -----
    thin = Side(style="thin", color="4472C4")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_letter, text in (("A", "Field"), ("B", "Value")):
        cell = ws[f"{col_letter}20"]
        cell.value = text
        cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border
    ws.row_dimensions[20].height = 20

    # ----- Data rows (21-32) -----
    data_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    even_fill = PatternFill("solid", fgColor="DCE6F1")
    odd_fill = PatternFill("solid", fgColor="FFFFFF")

    for idx, (label, cell_ref, multiline) in enumerate(INSTALLATION_FIELDS):
        row = 21 + idx
        # Label column
        lbl_cell = ws[f"A{row}"]
        lbl_cell.value = label
        lbl_cell.font = Font(name="Calibri", bold=True, size=10)
        lbl_cell.alignment = Alignment(horizontal="left", vertical="center")
        lbl_cell.fill = even_fill if idx % 2 == 0 else odd_fill
        lbl_cell.border = data_border

        # Value column (pre-fill empty)
        val_cell = ws[f"B{row}"]
        val_cell.value = ""
        val_cell.font = Font(name="Calibri", size=10)
        val_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        val_cell.fill = even_fill if idx % 2 == 0 else odd_fill
        val_cell.border = data_border
        if multiline:
            ws.row_dimensions[row].height = 60
        else:
            ws.row_dimensions[row].height = 18

    # Row 33 — pictures label
    ws.merge_cells("A33:B33")
    pic_cell = ws["A33"]
    pic_cell.value = "Attached Pictures"
    pic_cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    pic_cell.alignment = Alignment(horizontal="center", vertical="center")
    pic_cell.fill = PatternFill("solid", fgColor="4472C4")
    ws.row_dimensions[33].height = 20

    wb.save(path)


class InstallationSheetDialog(QDialog):
    """Pop-up dialog for filling and exporting the installation sheet."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fill up Installation Sheet")
        self.setMinimumSize(640, 700)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        self._uploaded_pictures: List[str] = []
        self._field_widgets: Dict[str, QWidget] = {}  # cell_ref -> widget

        # Ensure template exists
        tpl = _template_path()
        if not os.path.isfile(tpl):
            _create_template(tpl)

        self._init_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _init_ui(self):
        outer = QVBoxLayout()
        outer.setSpacing(8)
        outer.setContentsMargins(12, 12, 12, 12)

        # Title label
        title_lbl = QLabel("Installation Sheet")
        title_lbl.setStyleSheet(
            "font-size: 15pt; font-weight: bold; color: #1F497D; padding-bottom: 4px;"
        )
        outer.addWidget(title_lbl)

        # ---- Scrollable form area ----
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        form_container = QWidget()
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        form_layout.setSpacing(8)
        form_layout.setContentsMargins(4, 4, 4, 4)

        for label_text, cell_ref, multiline in INSTALLATION_FIELDS:
            lbl = QLabel(f"{label_text}:")
            lbl.setStyleSheet("font-weight: bold; font-size: 9pt;")
            lbl.setToolTip(f"Excel cell: {cell_ref}")

            if multiline:
                widget = QTextEdit()
                widget.setPlaceholderText(f"Enter {label_text.lower()} here…")
                widget.setFixedHeight(80)
                widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            else:
                widget = QLineEdit()
                widget.setPlaceholderText(f"Enter {label_text.lower()} here…")

            # Pre-populate date field
            if "date" in label_text.lower():
                if isinstance(widget, QLineEdit):
                    widget.setText(date.today().strftime("%d/%m/%Y"))

            self._field_widgets[cell_ref] = widget
            form_layout.addRow(lbl, widget)

        form_container.setLayout(form_layout)
        scroll.setWidget(form_container)
        outer.addWidget(scroll, stretch=3)

        # ---- Pictures section ----
        pics_group = QGroupBox("Attached Pictures")
        pics_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        pics_layout = QVBoxLayout()
        pics_layout.setSpacing(6)

        self._pics_grid = QGridLayout()
        self._pics_grid.setSpacing(6)

        pics_scroll = QScrollArea()
        pics_scroll.setWidgetResizable(True)
        pics_scroll.setMinimumHeight(110)
        pics_scroll.setMaximumHeight(160)
        pics_scroll.setFrameShape(QFrame.StyledPanel)

        pics_inner = QWidget()
        pics_inner.setLayout(self._pics_grid)
        pics_scroll.setWidget(pics_inner)
        pics_layout.addWidget(pics_scroll)

        self._no_pics_label = QLabel("No pictures added yet.")
        self._no_pics_label.setAlignment(Qt.AlignCenter)
        self._no_pics_label.setStyleSheet("color: #6c757d; font-size: 9pt;")
        self._pics_grid.addWidget(self._no_pics_label, 0, 0)

        pics_group.setLayout(pics_layout)
        outer.addWidget(pics_group, stretch=1)

        # ---- Action buttons ----
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self._add_pics_btn = QPushButton("📷  Add Pictures")
        self._add_pics_btn.setToolTip("Upload one or more pictures to include in the PDF")
        self._add_pics_btn.clicked.connect(self._add_pictures)
        self._add_pics_btn.setMinimumHeight(36)

        self._create_pdf_btn = QPushButton("📄  Create PDF")
        self._create_pdf_btn.setToolTip("Fill the Excel template and export it to PDF")
        self._create_pdf_btn.clicked.connect(self._create_pdf)
        self._create_pdf_btn.setMinimumHeight(36)
        self._create_pdf_btn.setStyleSheet(
            "QPushButton { background-color: #1F497D; color: white; font-weight: bold; border-radius: 4px; }"
            "QPushButton:hover { background-color: #2E6DA4; }"
            "QPushButton:pressed { background-color: #163a62; }"
        )

        self._clear_btn = QPushButton("🗑  Clear")
        self._clear_btn.setToolTip("Clear all text boxes and remove pictures")
        self._clear_btn.clicked.connect(self._clear_all)
        self._clear_btn.setMinimumHeight(36)

        btn_layout.addWidget(self._add_pics_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self._clear_btn)
        btn_layout.addWidget(self._create_pdf_btn)

        outer.addLayout(btn_layout)
        self.setLayout(outer)

    # ------------------------------------------------------------------
    # Slot: Add Pictures
    # ------------------------------------------------------------------

    def _add_pictures(self):
        """Open a file dialog to select pictures and add thumbnails."""
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Pictures",
            "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.webp);;All Files (*)",
        )
        if not paths:
            return

        for p in paths:
            if p not in self._uploaded_pictures:
                self._uploaded_pictures.append(p)

        self._refresh_picture_grid()

    def _refresh_picture_grid(self):
        """Rebuild the thumbnail grid from self._uploaded_pictures."""
        # Clear existing widgets
        while self._pics_grid.count():
            item = self._pics_grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        if not self._uploaded_pictures:
            self._no_pics_label = QLabel("No pictures added yet.")
            self._no_pics_label.setAlignment(Qt.AlignCenter)
            self._no_pics_label.setStyleSheet("color: #6c757d; font-size: 9pt;")
            self._pics_grid.addWidget(self._no_pics_label, 0, 0)
            return

        cols = 5
        for idx, path in enumerate(self._uploaded_pictures):
            row, col = divmod(idx, cols)
            thumb_frame = QFrame()
            thumb_frame.setFrameShape(QFrame.Box)
            thumb_frame.setStyleSheet("border: 1px solid #ced4da; border-radius: 4px;")
            thumb_layout = QVBoxLayout()
            thumb_layout.setContentsMargins(2, 2, 2, 2)
            thumb_layout.setSpacing(2)

            # Thumbnail image
            pix = QPixmap(path)
            if not pix.isNull():
                thumb_lbl = QLabel()
                thumb_lbl.setPixmap(pix.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                thumb_lbl.setAlignment(Qt.AlignCenter)
                thumb_layout.addWidget(thumb_lbl)
            else:
                thumb_layout.addWidget(QLabel("(no preview)"))

            # Filename (truncated)
            fname = os.path.basename(path)
            if len(fname) > 12:
                fname = fname[:9] + "…"
            name_lbl = QLabel(fname)
            name_lbl.setAlignment(Qt.AlignCenter)
            name_lbl.setStyleSheet("font-size: 7pt; color: #495057; border: none;")
            thumb_layout.addWidget(name_lbl)

            # Remove button
            remove_btn = QPushButton("✕")
            remove_btn.setFixedSize(20, 20)
            remove_btn.setStyleSheet(
                "QPushButton { background: #dc3545; color: white; border-radius: 2px; "
                "font-size: 8pt; border: none; }"
                "QPushButton:hover { background: #c82333; }"
            )
            _path = path  # capture for lambda
            remove_btn.clicked.connect(lambda _, p=_path: self._remove_picture(p))
            thumb_layout.addWidget(remove_btn, alignment=Qt.AlignCenter)

            thumb_frame.setLayout(thumb_layout)
            self._pics_grid.addWidget(thumb_frame, row, col)

    def _remove_picture(self, path: str):
        """Remove a picture from the list and refresh the grid."""
        if path in self._uploaded_pictures:
            self._uploaded_pictures.remove(path)
        self._refresh_picture_grid()

    # ------------------------------------------------------------------
    # Slot: Clear
    # ------------------------------------------------------------------

    def _clear_all(self):
        """Clear all text fields and uploaded pictures."""
        for widget in self._field_widgets.values():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QTextEdit):
                widget.clear()
        self._uploaded_pictures.clear()
        self._refresh_picture_grid()

    # ------------------------------------------------------------------
    # Slot: Create PDF
    # ------------------------------------------------------------------

    def _create_pdf(self):
        """
        Fill the Excel template with form data, embed pictures, then export to PDF.

        Strategy:
          1. Copy the template to a temporary file.
          2. Write each field value to its mapped cell using openpyxl.
          3. Insert uploaded pictures below the data rows.
          4. Try to export via win32com (Excel on Windows).
          5. Fall back to reportlab if win32com is unavailable.
        """
        # --- Ask where to save the PDF ---
        default_name = "installation_sheet.pdf"
        pdf_path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", default_name, "PDF Files (*.pdf)"
        )
        if not pdf_path:
            return

        try:
            self._do_create_pdf(pdf_path)
            QMessageBox.information(
                self,
                "PDF Created",
                f"Installation sheet saved successfully:\n{pdf_path}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Error Creating PDF",
                f"Could not create PDF:\n{exc}",
            )

    def _read_field_values(self) -> Dict[str, str]:
        """Return a mapping of excel_cell → current text value."""
        values: Dict[str, str] = {}
        for cell_ref, widget in self._field_widgets.items():
            if isinstance(widget, QLineEdit):
                values[cell_ref] = widget.text().strip()
            elif isinstance(widget, QTextEdit):
                values[cell_ref] = widget.toPlainText().strip()
        return values

    def _do_create_pdf(self, pdf_path: str):
        """Internal: fill Excel, embed pictures, export to PDF."""
        try:
            import openpyxl
        except ImportError as exc:
            raise RuntimeError(
                "The 'openpyxl' package is required. "
                "Install it with:  pip install openpyxl"
            ) from exc

        # --- Ensure template exists ---
        tpl = _template_path()
        if not os.path.isfile(tpl):
            _create_template(tpl)
        if not os.path.isfile(tpl):
            raise RuntimeError(f"Excel template not found: {tpl}")

        # --- Fill template in a temp file ---
        tmp_dir = tempfile.mkdtemp(prefix="sartel_install_")
        tmp_xlsx = os.path.join(tmp_dir, "installation_sheet_filled.xlsx")
        shutil.copy2(tpl, tmp_xlsx)

        wb = openpyxl.load_workbook(tmp_xlsx)
        ws = wb.active

        field_values = self._read_field_values()
        for cell_ref, value in field_values.items():
            ws[cell_ref] = value

        # --- Embed pictures (if any) ---
        if self._uploaded_pictures:
            try:
                from openpyxl.drawing.image import Image as XLImage
                from openpyxl.utils import get_column_letter

                pic_start_row = 34  # row below the "Attached Pictures" header
                col = 1             # column A

                for pic_path in self._uploaded_pictures:
                    if not os.path.isfile(pic_path):
                        continue
                    try:
                        xl_img = XLImage(pic_path)
                        xl_img.width = 300
                        xl_img.height = 200
                        cell_addr = f"{get_column_letter(col)}{pic_start_row}"
                        ws.add_image(xl_img, cell_addr)
                        pic_start_row += 16   # ~200px / 12.75 px per row ≈ 16 rows
                    except Exception:
                        pass  # skip unreadable images
            except ImportError:
                pass  # openpyxl drawing support unavailable

        wb.save(tmp_xlsx)

        # --- Export to PDF ---
        pdf_created = False

        # Attempt 1: win32com (Excel on Windows)
        if not pdf_created:
            try:
                import win32com.client
                import pythoncom

                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb_com = excel.Workbooks.Open(os.path.abspath(tmp_xlsx))
                    wb_com.ExportAsFixedFormat(
                        0,  # xlTypePDF
                        os.path.abspath(pdf_path),
                        1,  # xlQualityStandard
                        True,
                        False,
                    )
                    wb_com.Close(False)
                    pdf_created = True
                finally:
                    excel.Quit()
                    pythoncom.CoUninitialize()
            except Exception:
                pass  # fall through to reportlab

        # Attempt 2: reportlab fallback
        if not pdf_created:
            self._create_pdf_reportlab(pdf_path, field_values)
            pdf_created = True

        # Clean up temp dir
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass

    def _create_pdf_reportlab(self, pdf_path: str, field_values: Dict[str, str]):
        """Generate a formatted PDF using reportlab when Excel COM is not available."""
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.lib import colors
            from reportlab.platypus import (
                SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                Image as RLImage, HRFlowable,
            )
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
        except ImportError as exc:
            raise RuntimeError(
                "Neither Microsoft Excel (win32com) nor the 'reportlab' package is available. "
                "Install reportlab with:  pip install reportlab"
            ) from exc

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "SheetTitle",
            parent=styles["Title"],
            fontSize=16,
            textColor=colors.HexColor("#1F497D"),
            spaceAfter=6,
            alignment=TA_CENTER,
        )
        sub_style = ParagraphStyle(
            "SheetSub",
            parent=styles["Normal"],
            fontSize=10,
            textColor=colors.HexColor("#4472C4"),
            spaceAfter=12,
            alignment=TA_CENTER,
            italic=True,
        )
        field_label_style = ParagraphStyle(
            "FieldLabel",
            parent=styles["Normal"],
            fontSize=9,
            fontName="Helvetica-Bold",
        )
        field_value_style = ParagraphStyle(
            "FieldValue",
            parent=styles["Normal"],
            fontSize=9,
        )

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=A4,
            leftMargin=20 * mm,
            rightMargin=20 * mm,
            topMargin=20 * mm,
            bottomMargin=20 * mm,
        )
        story = []

        story.append(Paragraph("SARTEL – Installation Sheet", title_style))
        story.append(Paragraph("CAN Bus Monitoring – Field Installation Record", sub_style))
        story.append(HRFlowable(width="100%", color=colors.HexColor("#4472C4"), thickness=1.5))
        story.append(Spacer(1, 8 * mm))

        # Build table rows
        table_data = [
            [
                Paragraph("<b>Field</b>", field_label_style),
                Paragraph("<b>Value</b>", field_label_style),
            ]
        ]
        for label_text, cell_ref, _ in INSTALLATION_FIELDS:
            val = field_values.get(cell_ref, "")
            table_data.append([
                Paragraph(label_text, field_label_style),
                Paragraph(val if val else "—", field_value_style),
            ])

        page_w = A4[0] - 40 * mm
        col_widths = [page_w * 0.38, page_w * 0.62]

        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
        header_bg = colors.HexColor("#4472C4")
        even_bg = colors.HexColor("#DCE6F1")
        white = colors.white

        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("TEXTCOLOR",  (0, 0), (-1, 0), white),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, even_bg]),
            ("GRID",  (0, 0), (-1, -1), 0.5, colors.HexColor("#ADB5BD")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ])
        tbl.setStyle(tbl_style)
        story.append(tbl)

        # Attached pictures
        if self._uploaded_pictures:
            story.append(Spacer(1, 8 * mm))
            story.append(HRFlowable(width="100%", color=colors.HexColor("#4472C4"), thickness=1))
            story.append(Spacer(1, 4 * mm))
            story.append(Paragraph("Attached Pictures", title_style))
            story.append(Spacer(1, 4 * mm))

            max_w = page_w * 0.75
            max_h = 120 * mm

            for pic_path in self._uploaded_pictures:
                if not os.path.isfile(pic_path):
                    continue
                try:
                    img = RLImage(pic_path)
                    # Scale proportionally to fit
                    scale = min(max_w / img.imageWidth, max_h / img.imageHeight, 1.0)
                    img.drawWidth  = img.imageWidth  * scale
                    img.drawHeight = img.imageHeight * scale
                    story.append(img)
                    story.append(Spacer(1, 4 * mm))
                    fname = os.path.basename(pic_path)
                    story.append(
                        Paragraph(
                            fname,
                            ParagraphStyle(
                                "PicCaption", parent=styles["Normal"],
                                fontSize=8, textColor=colors.grey, alignment=TA_CENTER,
                            ),
                        )
                    )
                    story.append(Spacer(1, 6 * mm))
                except Exception:
                    pass  # skip unreadable images

        doc.build(story)
