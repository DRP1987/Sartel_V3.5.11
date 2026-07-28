"""Installation sheet dialog for filling and exporting installation data."""

import json
import os
import sys
import shutil
import tempfile
from datetime import date
from typing import Any, Dict, List, Optional

from PyQt5.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QScrollArea, QWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
    QFileDialog, QMessageBox, QGroupBox, QGridLayout, QSizePolicy, QFrame,
    QDialogButtonBox,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap


# ---------------------------------------------------------------------------
# Helpers: locate the config folder
# ---------------------------------------------------------------------------

def _config_dir() -> str:
    """Return the absolute path to the config directory."""
    candidates = [
        os.path.join(os.path.dirname(__file__), "..", "config"),
        os.path.join(getattr(sys, "_MEIPASS", ""), "config"),
    ]
    for p in candidates:
        if os.path.isdir(p):
            return os.path.abspath(p)
    # fallback – sibling of gui package
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "config"))


def _sheets_config_path() -> str:
    """Return the path to installation_sheets.json inside config/."""
    return os.path.join(_config_dir(), "installation_sheets.json")


def _load_sheets_config() -> List[Dict[str, Any]]:
    """Load and return the list of sheet definitions from installation_sheets.json."""
    path = _sheets_config_path()
    if not os.path.isfile(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data.get("sheets", [])
    except Exception:
        return []


def _excel_template_path(filename: str) -> str:
    """Return the absolute path for a template Excel file stored in config/."""
    return os.path.join(_config_dir(), filename)


# ---------------------------------------------------------------------------
# Sheet-selection dialog (shown before the form)
# ---------------------------------------------------------------------------

class _SheetSelectionDialog(QDialog):
    """Small dialog that lets the user pick which installation sheet to fill."""

    def __init__(self, sheet_names: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Installation Sheet")
        self.setMinimumWidth(380)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        lbl = QLabel("Please select the installation sheet type:")
        lbl.setStyleSheet("font-size: 10pt; font-weight: bold;")
        layout.addWidget(lbl)

        self.combo = QComboBox()
        self.combo.addItems(sheet_names)
        self.combo.setMinimumHeight(32)
        layout.addWidget(self.combo)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    @property
    def selected_index(self) -> int:
        return self.combo.currentIndex()


# ---------------------------------------------------------------------------
# Main dialog
# ---------------------------------------------------------------------------

class InstallationSheetDialog(QDialog):
    """Pop-up dialog for filling and exporting the installation sheet.

    Opens a sheet-selection dropdown first, then renders a dynamic form whose
    fields (text, multiline, or checkbox) and Excel cell mappings are driven by
    ``config/installation_sheets.json``.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fill up Installation Sheet")
        self.setMinimumSize(640, 700)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        self._uploaded_pictures: List[str] = []
        self._field_widgets: Dict[str, QWidget] = {}   # cell_ref -> widget
        self._current_sheet_def: Optional[Dict[str, Any]] = None

        # Load available sheet definitions
        self._sheets = _load_sheets_config()

        # Ask the user which sheet to use *before* building the form
        if not self._select_sheet():
            # User cancelled selection – close dialog immediately
            self._cancelled = True
            return
        self._cancelled = False
        self._init_ui()

    # ------------------------------------------------------------------
    # Sheet selection
    # ------------------------------------------------------------------

    def _select_sheet(self) -> bool:
        """Show the sheet-selection dialog.  Returns True if accepted."""
        if not self._sheets:
            QMessageBox.warning(
                self,
                "No Sheet Definitions",
                "No installation sheet definitions were found.\n"
                f"Please check: {_sheets_config_path()}",
            )
            return False

        names = [s.get("name", s.get("id", f"Sheet {i}")) for i, s in enumerate(self._sheets)]
        dlg = _SheetSelectionDialog(names, parent=self)
        if dlg.exec_() != QDialog.Accepted:
            return False

        self._current_sheet_def = self._sheets[dlg.selected_index]
        return True

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _init_ui(self):
        sheet_def = self._current_sheet_def
        sheet_name = sheet_def.get("name", "Installation Sheet")
        fields: List[Dict[str, Any]] = sheet_def.get("fields", [])

        outer = QVBoxLayout()
        outer.setSpacing(8)
        outer.setContentsMargins(12, 12, 12, 12)

        # Title label (shows selected sheet name)
        title_lbl = QLabel(sheet_name)
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

        for field in fields:
            label_text = field.get("label", "")
            cell_ref = field.get("cell", "")
            field_type = field.get("type", "text").lower()

            lbl = QLabel(f"{label_text}:")
            lbl.setStyleSheet("font-weight: bold; font-size: 9pt;")
            lbl.setToolTip(f"Excel cell: {cell_ref}")

            if field_type == "checkbox":
                widget = QCheckBox()
                widget.setToolTip(f"Tick to mark '{label_text}' in Excel cell {cell_ref}")
            elif field_type == "multiline":
                widget = QTextEdit()
                widget.setPlaceholderText(f"Enter {label_text.lower()} here…")
                widget.setFixedHeight(80)
                widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            else:  # "text"
                widget = QLineEdit()
                widget.setPlaceholderText(f"Enter {label_text.lower()} here…")
                # Pre-populate date fields
                if "date" in label_text.lower():
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

    # Override exec_ so a cancelled selection closes without showing the window
    def exec_(self):
        if getattr(self, "_cancelled", False):
            return QDialog.Rejected
        return super().exec_()

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
        """Clear all text fields, uncheck checkboxes, and remove pictures."""
        for widget in self._field_widgets.values():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QTextEdit):
                widget.clear()
            elif isinstance(widget, QCheckBox):
                widget.setChecked(False)
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

    def _read_field_values(self) -> Dict[str, Any]:
        """Return a mapping of excel_cell → current value (str or bool)."""
        values: Dict[str, Any] = {}
        for cell_ref, widget in self._field_widgets.items():
            if isinstance(widget, QLineEdit):
                values[cell_ref] = widget.text().strip()
            elif isinstance(widget, QTextEdit):
                values[cell_ref] = widget.toPlainText().strip()
            elif isinstance(widget, QCheckBox):
                values[cell_ref] = widget.isChecked()
        return values

    def _do_create_pdf(self, pdf_path: str):
        """Internal: fill Excel template from config/, embed pictures, export to PDF."""
        try:
            import openpyxl
        except ImportError as exc:
            raise RuntimeError(
                "The 'openpyxl' package is required. "
                "Install it with:  pip install openpyxl"
            ) from exc

        sheet_def = self._current_sheet_def or {}
        excel_filename = sheet_def.get("excel_file", "")
        sheet_index = sheet_def.get("sheet_index", 0)

        # --- Locate the Excel template in config/ ---
        if excel_filename:
            tpl = _excel_template_path(excel_filename)
        else:
            tpl = ""

        if not tpl or not os.path.isfile(tpl):
            raise RuntimeError(
                f"Excel template not found in config folder: {excel_filename or '(none specified)'}\n"
                f"Please place the file in: {_config_dir()}"
            )

        # --- Fill template in a temp file ---
        tmp_dir = tempfile.mkdtemp(prefix="sartel_install_")
        tmp_xlsx = os.path.join(tmp_dir, "installation_sheet_filled.xlsx")
        shutil.copy2(tpl, tmp_xlsx)

        wb = openpyxl.load_workbook(tmp_xlsx)

        # Select the correct worksheet by index
        if sheet_index < len(wb.sheetnames):
            ws = wb.worksheets[sheet_index]
        else:
            ws = wb.active

        field_values = self._read_field_values()
        for cell_ref, value in field_values.items():
            if isinstance(value, bool):
                ws[cell_ref] = value  # write as boolean (TRUE/FALSE in Excel)
            else:
                ws[cell_ref] = value

        # --- Embed pictures (if any) ---
        if self._uploaded_pictures:
            try:
                from openpyxl.drawing.image import Image as XLImage
                from openpyxl.utils import get_column_letter

                pic_start_row = ws.max_row + 2
                col = 1  # column A

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

    def _create_pdf_reportlab(self, pdf_path: str, field_values: Dict[str, Any]):
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
        sheet_def = self._current_sheet_def or {}
        sheet_name = sheet_def.get("name", "Installation Sheet")
        fields: List[Dict[str, Any]] = sheet_def.get("fields", [])

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

        story.append(Paragraph(f"SARTEL – {sheet_name}", title_style))
        story.append(Paragraph("CAN Bus Monitoring – Field Installation Record", sub_style))
        story.append(HRFlowable(width="100%", color=colors.HexColor("#4472C4"), thickness=1.5))
        story.append(Spacer(1, 8 * mm))

        # Build table rows from the dynamic field list
        table_data = [
            [
                Paragraph("<b>Field</b>", field_label_style),
                Paragraph("<b>Value</b>", field_label_style),
            ]
        ]
        for field in fields:
            label_text = field.get("label", "")
            cell_ref = field.get("cell", "")
            raw_val = field_values.get(cell_ref, "")
            # Format booleans as ✔ / ✘
            if isinstance(raw_val, bool):
                display_val = "✔ Yes" if raw_val else "✘ No"
            else:
                display_val = str(raw_val) if raw_val else "—"
            table_data.append([
                Paragraph(label_text, field_label_style),
                Paragraph(display_val, field_value_style),
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
