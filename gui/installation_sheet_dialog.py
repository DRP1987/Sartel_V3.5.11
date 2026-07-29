"""Installation sheet dialog for filling and exporting installation data."""

import json
import os
import re
import sys
import shutil
import tempfile
from datetime import date
from typing import Any, Dict, List, Optional

from PyQt5.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QScrollArea, QWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
    QFileDialog, QMessageBox, QGroupBox, QGridLayout, QSizePolicy, QFrame,
    QDialogButtonBox, QRadioButton,
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


def _resolve_cell_ref(ws, cell_ref: str) -> str:
    """Resolve a cell reference to a writable top-left address.

    Handles two cases:
    - Range notation: ``"C10:E12"`` → returns ``"C10"`` (top-left of the range).
    - Single cell that belongs to a merged region in *ws*: returns the top-left
      cell of that merged region so openpyxl writes the value correctly.
    - Single cell outside any merge: returns *cell_ref* unchanged.
    """
    from openpyxl.utils import get_column_letter, column_index_from_string

    # Handle explicit range notation e.g. "C10:E12" or "c10:e12"
    cell_ref = cell_ref.strip()
    if ":" in cell_ref:
        top_left = cell_ref.split(":")[0].strip().upper()
        return top_left

    # Parse the column/row indices of the single cell
    m = re.match(r"([A-Za-z]+)(\d+)$", cell_ref)
    if not m:
        return cell_ref  # unrecognised format – pass through unchanged

    col_letter, row_str = m.groups()
    col_idx = column_index_from_string(col_letter.upper())
    row_idx = int(row_str)

    # Check every merged region in the worksheet
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= row_idx <= merged_range.max_row
            and merged_range.min_col <= col_idx <= merged_range.max_col
        ):
            return f"{get_column_letter(merged_range.min_col)}{merged_range.min_row}"

    return cell_ref


# ---------------------------------------------------------------------------
# Conditional (Yes/No) widget
# ---------------------------------------------------------------------------

class ConditionalWidget(QWidget):
    """A compound widget for yes/no conditional questions.

    Renders two radio buttons ("Yes" / "No").  When the user selects "Yes" a
    configurable text box appears; when "No" is selected a different text box
    appears.  The boolean answer is written to *checkbox_cell* in the Excel
    template; the visible text content is written to *yes_text_cell* or
    *no_text_cell* depending on the selection.

    JSON field definition example::

        {
          "label": "Is the crane control CanBus connected for LMB and Crane operation?",
          "type": "conditional",
          "checkbox_cell": "C34",
          "yes_text_label": "Connection Details",
          "yes_text_cell": "D34",
          "no_text_label": "Reason for Non-Connection",
          "no_text_cell": "E34"
        }
    """

    def __init__(self, field_def: Dict[str, Any], parent=None):
        super().__init__(parent)
        self._checkbox_cell: str = field_def.get("checkbox_cell", "")
        self._yes_text_cell: str = field_def.get("yes_text_cell", "")
        self._no_text_cell: str = field_def.get("no_text_cell", "")
        yes_label = field_def.get("yes_text_label", "Details (Yes)")
        no_label = field_def.get("no_text_label", "Details (No)")

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # --- Yes / No radio buttons ---
        radio_row = QHBoxLayout()
        radio_row.setSpacing(16)
        self._yes_radio = QRadioButton("Yes")
        self._no_radio = QRadioButton("No")
        self._yes_radio.setStyleSheet("font-size: 9pt;")
        self._no_radio.setStyleSheet("font-size: 9pt;")
        radio_row.addWidget(self._yes_radio)
        radio_row.addWidget(self._no_radio)
        radio_row.addStretch()
        layout.addLayout(radio_row)

        # --- Text box shown when "Yes" is selected ---
        self._yes_container: Optional[QWidget] = None
        self._yes_text: Optional[QTextEdit] = None
        if self._yes_text_cell:
            self._yes_container = QWidget()
            yes_inner = QVBoxLayout()
            yes_inner.setContentsMargins(0, 2, 0, 0)
            yes_inner.setSpacing(2)
            yes_lbl = QLabel(f"{yes_label}:")
            yes_lbl.setStyleSheet("font-size: 8pt; color: #1a6a2e; font-weight: bold;")
            self._yes_text = QTextEdit()
            self._yes_text.setPlaceholderText(yes_label)
            self._yes_text.setFixedHeight(60)
            self._yes_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            yes_inner.addWidget(yes_lbl)
            yes_inner.addWidget(self._yes_text)
            self._yes_container.setLayout(yes_inner)
            self._yes_container.setVisible(False)
            layout.addWidget(self._yes_container)

        # --- Text box shown when "No" is selected ---
        self._no_container: Optional[QWidget] = None
        self._no_text: Optional[QTextEdit] = None
        if self._no_text_cell:
            self._no_container = QWidget()
            no_inner = QVBoxLayout()
            no_inner.setContentsMargins(0, 2, 0, 0)
            no_inner.setSpacing(2)
            no_lbl = QLabel(f"{no_label}:")
            no_lbl.setStyleSheet("font-size: 8pt; color: #a93226; font-weight: bold;")
            self._no_text = QTextEdit()
            self._no_text.setPlaceholderText(no_label)
            self._no_text.setFixedHeight(60)
            self._no_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            no_inner.addWidget(no_lbl)
            no_inner.addWidget(self._no_text)
            self._no_container.setLayout(no_inner)
            self._no_container.setVisible(False)
            layout.addWidget(self._no_container)

        self.setLayout(layout)

        # Connect radio button signals
        self._yes_radio.toggled.connect(self._on_selection_changed)
        self._no_radio.toggled.connect(self._on_selection_changed)

    # ------------------------------------------------------------------
    # Internal slot
    # ------------------------------------------------------------------

    def _on_selection_changed(self):
        """Show/hide the appropriate text box when the selection changes."""
        yes_checked = self._yes_radio.isChecked()
        no_checked = self._no_radio.isChecked()
        if self._yes_container is not None:
            self._yes_container.setVisible(yes_checked)
        if self._no_container is not None:
            self._no_container.setVisible(no_checked)

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------

    def get_values(self) -> Dict[str, Any]:
        """Return a ``{cell_ref: value}`` mapping for all configured cells.

        * *checkbox_cell* → ``True`` (Yes), ``False`` (No), or omitted if
          neither radio button has been selected.
        * *yes_text_cell* → text content when "Yes" is selected.
        * *no_text_cell*  → text content when "No" is selected.
        """
        result: Dict[str, Any] = {}

        if self._yes_radio.isChecked():
            is_yes: Optional[bool] = True
        elif self._no_radio.isChecked():
            is_yes = False
        else:
            is_yes = None

        if self._checkbox_cell and is_yes is not None:
            result[self._checkbox_cell] = is_yes

        if is_yes is True and self._yes_text_cell and self._yes_text is not None:
            result[self._yes_text_cell] = self._yes_text.toPlainText().strip()
        elif is_yes is False and self._no_text_cell and self._no_text is not None:
            result[self._no_text_cell] = self._no_text.toPlainText().strip()

        return result

    def clear(self):
        """Reset both radio buttons and clear all text boxes."""
        # Temporarily disable mutual exclusion to allow unchecking both
        self._yes_radio.setAutoExclusive(False)
        self._no_radio.setAutoExclusive(False)
        self._yes_radio.setChecked(False)
        self._no_radio.setChecked(False)
        self._yes_radio.setAutoExclusive(True)
        self._no_radio.setAutoExclusive(True)
        if self._yes_text is not None:
            self._yes_text.clear()
        if self._no_text is not None:
            self._no_text.clear()
        # Ensure both containers are hidden
        if self._yes_container is not None:
            self._yes_container.setVisible(False)
        if self._no_container is not None:
            self._no_container.setVisible(False)


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

            if field_type == "conditional":
                # Conditional question: Yes/No radio buttons + conditional text boxes.
                # The question may be long, so allow the label to word-wrap and
                # align it to the top of the row.
                lbl.setWordWrap(True)
                lbl.setAlignment(Qt.AlignRight | Qt.AlignTop)
                checkbox_cell = field.get("checkbox_cell", "")
                cells_hint = f"Checkbox cell: {checkbox_cell}"
                if field.get("yes_text_cell"):
                    cells_hint += f"  |  Yes text cell: {field['yes_text_cell']}"
                if field.get("no_text_cell"):
                    cells_hint += f"  |  No text cell: {field['no_text_cell']}"
                lbl.setToolTip(cells_hint)
                widget = ConditionalWidget(field)
                # Store using checkbox_cell as key so _read_field_values can
                # detect and call widget.get_values()
                self._field_widgets[checkbox_cell or cell_ref] = widget
                form_layout.addRow(lbl, widget)
                continue

            # Build a human-readable cell hint for tooltips (range or single cell)
            cell_hint = f"Excel cell{'s' if ':' in cell_ref else ''}: {cell_ref}"
            lbl.setToolTip(cell_hint)

            if field_type == "checkbox":
                widget = QCheckBox()
                widget.setToolTip(f"Tick to mark '{label_text}' in {cell_hint}")
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
            elif isinstance(widget, ConditionalWidget):
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

    def _read_field_values(self) -> Dict[str, Any]:
        """Return a mapping of excel_cell → current value (str or bool).

        For :class:`ConditionalWidget` entries the widget's own
        :meth:`~ConditionalWidget.get_values` method is called, which may
        contribute up to three cell mappings (checkbox_cell, yes_text_cell,
        no_text_cell).
        """
        values: Dict[str, Any] = {}
        for cell_ref, widget in self._field_widgets.items():
            if isinstance(widget, ConditionalWidget):
                values.update(widget.get_values())
            elif isinstance(widget, QLineEdit):
                values[cell_ref] = widget.text().strip()
            elif isinstance(widget, QTextEdit):
                values[cell_ref] = widget.toPlainText().strip()
            elif isinstance(widget, QCheckBox):
                values[cell_ref] = widget.isChecked()
        return values

    def _embed_pictures_in_excel(self, ws, anchor_cell: str = None):
        """Embed uploaded pictures directly into the Excel worksheet.

        Inserts each picture as an over-cell image (Excel's "Insert Picture
        Over Cells") starting at *anchor_cell* (e.g. ``"A35"``).  Every image
        is unconditionally resized to exactly 25 cm wide × 45 cm tall so the
        output is consistent regardless of the original picture dimensions.
        Note: images that do not share the 25:45 aspect ratio will appear
        stretched or compressed; this is intentional to ensure a uniform size
        in the exported PDF.  When no anchor is given the pictures are placed
        three rows below the last populated row.  Multiple pictures are stacked
        vertically, each starting directly below the previous one.
        """
        if not self._uploaded_pictures:
            return

        try:
            from openpyxl.drawing.image import Image as OXLImage
        except ImportError:
            return  # openpyxl image support unavailable – skip

        # Determine the starting row / column for picture insertion
        if anchor_cell:
            m = re.match(r"([A-Za-z]+)(\d+)$", anchor_cell.strip())
            if m:
                start_col_letter = m.group(1).upper()
                current_row = int(m.group(2))
            else:
                start_col_letter = "A"
                current_row = (ws.max_row or 0) + 3
        else:
            start_col_letter = "A"
            current_row = (ws.max_row or 0) + 3

        # Fixed target size: 25 cm wide × 45 cm tall at 96 DPI (1 cm ≈ 37.795 px)
        CM_TO_PX = 96 / 2.54
        target_w_px = round(25 * CM_TO_PX)   # ≈ 945 px
        target_h_px = round(45 * CM_TO_PX)   # ≈ 1701 px
        # Approximate Excel row height in pixels (default row ≈ 15 pt ≈ 20 px at 96 DPI)
        px_per_row = 20

        for pic_path in self._uploaded_pictures:
            if not os.path.isfile(pic_path):
                continue
            try:
                img = OXLImage(pic_path)
                # Always force the picture to exactly 25 cm × 45 cm
                img.width = target_w_px
                img.height = target_h_px

                img.anchor = f"{start_col_letter}{current_row}"
                ws.add_image(img)

                # Advance the row pointer so the next picture starts below this one
                rows_used = max(1, target_h_px // px_per_row) + 2
                current_row += rows_used
            except (IOError, OSError, ValueError):
                continue  # skip unreadable or unsupported images

    def _do_create_pdf(self, pdf_path: str):
        """Internal: fill Excel template from config/, embed pictures, export to PDF.

        Strategy
        --------
        1. Copy the template to a temporary file.
        2. Write each field value into its mapped cell using openpyxl.
        3. Embed uploaded pictures directly into the worksheet using openpyxl
           (simulating Excel's "Insert Picture Over Cells"), so the single
           Excel → PDF export step captures both the data and the pictures.
        4. Export to PDF via win32com / Excel (Windows) if available.
        5. Fall back to reportlab when win32com is unavailable.
        """
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
        picture_anchor = sheet_def.get("picture_anchor_cell", None)

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
            # Resolve the writable cell (handles merged cells and range notation)
            writable_ref = _resolve_cell_ref(ws, cell_ref)
            ws[writable_ref] = value  # bool writes as TRUE/FALSE; str writes as text

        # Embed pictures directly into the worksheet so they are captured
        # when Excel exports the sheet to PDF.
        self._embed_pictures_in_excel(ws, anchor_cell=picture_anchor)

        wb.save(tmp_xlsx)

        # --- Export to PDF ---
        pdf_created = False

        # Attempt 1: win32com (Excel on Windows)
        # Pictures are already embedded in the xlsx, so a single export produces
        # the final PDF with both the sheet data and the pictures.
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
                finally:
                    excel.Quit()
                    pythoncom.CoUninitialize()

                pdf_created = True
            except Exception:
                # win32com raises pywintypes.com_error (a dynamic type only importable
                # when pywin32 is installed) in addition to ImportError when the package
                # is absent.  A broad catch is intentional here so that any Excel/COM
                # failure falls through to the reportlab fallback.
                pass

        # Attempt 2: reportlab fallback (when Excel / win32com is not available)
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
            field_type = field.get("type", "text").lower()

            if field_type == "conditional":
                # Show the boolean answer + the associated text note
                checkbox_cell = field.get("checkbox_cell", "")
                yes_text_cell = field.get("yes_text_cell", "")
                no_text_cell = field.get("no_text_cell", "")
                raw_bool = field_values.get(checkbox_cell)
                if raw_bool is True:
                    yes_text = field_values.get(yes_text_cell, "")
                    display_val = f"✔ Yes" + (f" – {yes_text}" if yes_text else "")
                elif raw_bool is False:
                    no_text = field_values.get(no_text_cell, "")
                    display_val = f"✘ No" + (f" – {no_text}" if no_text else "")
                else:
                    display_val = "—"
            else:
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
