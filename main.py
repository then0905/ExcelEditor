#!/usr/bin/env python3
"""JsonEditor Pro — PySide6 · App-quality dark UI (Spec v2)"""

import sys, os, json, re
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QTabWidget,
    QListWidget, QListWidgetItem, QLineEdit, QPushButton, QLabel,
    QVBoxLayout, QHBoxLayout, QScrollArea, QCheckBox, QComboBox,
    QTextEdit, QTableView, QHeaderView, QAbstractItemView,
    QStackedWidget, QFileDialog, QMenu, QSizePolicy, QFrame,
    QInputDialog, QMessageBox, QStyledItemDelegate, QStyle,
    QDialog, QDialogButtonBox, QFormLayout, QLayout, QCompleter,
    QTableWidget, QTableWidgetItem,
)
from PySide6.QtCore import (
    Qt, Signal, QAbstractTableModel, QModelIndex,
    QTimer, QSize, QThread, QRect, QRectF, QPoint, QStringListModel,
)
from PySide6.QtGui import (
    QAction, QColor, QKeySequence, QFont, QBrush,
    QPainter, QPen, QLinearGradient, QPainterPath,
    QIcon, QPixmap, QFontDatabase,
)

from json_data_manager import JsonDataManager


# ═══════════════════════════════════════════════════════════════
#  DESIGN SYSTEM
# ═══════════════════════════════════════════════════════════════

_C = {
    "bg":      "#0F0F1A",
    "sidebar": "#14141F",
    "panel":   "#17172A",
    "card":    "#1E1E32",
    "cardH":   "#252540",
    "input":   "#1C1C30",
    "code":    "#12121F",
    "border":  "#2A2A40",
    "borderH": "#3A3A54",
    "accent":  "#6366F1",
    "txt":     "#F0F0F5",
    "txt2":    "#8B8BA3",
    "txt3":    "#5A5A72",
    "txtAcc":  "#A5B4FC",
    "green":   "#10B981",
    "yellow":  "#EAB308",
    "red":     "#EF4444",
    "cyan":    "#0891B2",
}

_CAT = [
    {"color": "#F59E0B", "r": 245, "g": 158, "b": 11,  "text": "#FDE68A"},
    {"color": "#8B5CF6", "r": 139, "g": 92,  "b": 246, "text": "#C4B5FD"},
    {"color": "#10B981", "r": 16,  "g": 185, "b": 129, "text": "#6EE7B7"},
    {"color": "#EC4899", "r": 236, "g": 72,  "b": 153, "text": "#F9A8D4"},
    {"color": "#6366F1", "r": 99,  "g": 102, "b": 241, "text": "#A5B4FC"},
    {"color": "#EF4444", "r": 239, "g": 68,  "b": 68,  "text": "#FCA5A5"},
    {"color": "#0891B2", "r": 8,   "g": 145, "b": 178, "text": "#67E8F9"},
    {"color": "#EAB308", "r": 234, "g": 179, "b": 8,   "text": "#FDE68A"},
]

_cat_assign: dict[str, int] = {}


def _cat_for(val: str) -> dict:
    s = str(val)
    if s not in _cat_assign:
        _cat_assign[s] = len(_cat_assign) % len(_CAT)
    return _CAT[_cat_assign[s]]


def _cat_qcolor(val: str, alpha: int = 255) -> QColor:
    c = _cat_for(val)
    return QColor(c["r"], c["g"], c["b"], alpha)


# ── Global QSS ────────────────────────────────────────────────────────────────

APP_QSS = f"""
* {{
    font-family: "Segoe UI", "Noto Sans TC", sans-serif;
    font-size: 12px;
    color: {_C['txt']};
}}
QMainWindow, QDialog {{ background: {_C['bg']}; }}
QWidget {{ background: transparent; }}

QSplitter::handle:horizontal {{ background: {_C['border']}; width: 1px; }}
QSplitter::handle:vertical   {{ background: {_C['border']}; height: 1px; }}

/* ── Outer tab bar (table switcher) ── */
QTabWidget#main-tabs::pane {{
    border: none;
    background: {_C['bg']};
}}
QTabWidget#main-tabs > QTabBar {{
    background: {_C['sidebar']};
    border-bottom: 1px solid {_C['border']};
}}
QTabWidget#main-tabs > QTabBar::tab {{
    background: transparent;
    color: {_C['txt2']};
    padding: 9px 20px;
    border: none;
    border-bottom: 2px solid transparent;
    font-size: 12px;
    font-weight: 500;
    min-width: 60px;
}}
QTabWidget#main-tabs > QTabBar::tab:selected {{
    color: {_C['txt']};
    border-bottom: 2px solid {_C['accent']};
    font-weight: 600;
}}
QTabWidget#main-tabs > QTabBar::tab:hover:!selected {{
    color: {_C['txt']};
    background: rgba(255,255,255,0.03);
}}

/* ── Sub-table tab bar ── */
QTabWidget#sub-tabs::pane {{
    border: none;
    background: {_C['panel']};
}}
QTabWidget#sub-tabs > QTabBar {{
    background: {_C['panel']};
    border-bottom: 1px solid {_C['border']};
}}
QTabWidget#sub-tabs > QTabBar::tab {{
    background: transparent;
    color: {_C['txt2']};
    padding: 6px 14px;
    border: none;
    border-bottom: 2px solid transparent;
    font-size: 11px;
}}
QTabWidget#sub-tabs > QTabBar::tab:selected {{
    color: {_C['accent']};
    border-bottom: 2px solid {_C['accent']};
}}
QTabWidget#sub-tabs > QTabBar::tab:hover:!selected {{
    color: {_C['txt']};
}}

/* ── Classification list ── */
QListWidget#cls-list {{
    background: transparent;
    border: none;
    outline: none;
    padding: 4px;
}}
QListWidget#cls-list::item {{
    padding: 8px 10px;
    border-radius: 7px;
    margin: 1px 0;
}}
QListWidget#cls-list::item:selected {{
    background: rgba(99,102,241,0.18);
}}
QListWidget#cls-list::item:hover:!selected {{
    background: rgba(255,255,255,0.04);
}}

/* ── Card list ── */
QListWidget#card-list {{
    background: transparent;
    border: none;
    outline: none;
}}
QListWidget#card-list::item {{
    background: transparent;
    border: none;
    padding: 0;
    margin: 0;
}}

/* ── Inputs ── */
QLineEdit {{
    background: {_C['input']};
    border: 1px solid {_C['border']};
    border-radius: 7px;
    padding: 7px 10px;
    color: {_C['txt']};
    selection-background-color: {_C['accent']};
}}
QLineEdit:focus {{ border-color: {_C['accent']}; }}
QLineEdit[invalid="true"] {{
    border-color: {_C['red']};
    background: rgba(239,68,68,0.08);
}}
QTextEdit {{
    background: {_C['input']};
    border: 1px solid {_C['border']};
    border-radius: 7px;
    padding: 6px 10px;
    color: {_C['txt']};
    selection-background-color: {_C['accent']};
}}
QTextEdit:focus {{ border-color: {_C['accent']}; }}
QTextEdit#code-view {{
    background: {_C['code']};
    border: 1px solid {_C['border']};
    font-family: "Consolas", monospace;
    font-size: 11px;
    color: {_C['txt2']};
    border-radius: 8px;
    padding: 10px 12px;
}}
QComboBox {{
    background: {_C['input']};
    border: 1px solid {_C['border']};
    border-radius: 7px;
    padding: 6px 10px;
    color: {_C['txt']};
}}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox::down-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid {_C['txt2']};
    margin-right: 6px;
}}
QComboBox:focus {{ border-color: {_C['accent']}; }}
QComboBox:disabled {{ color: {_C['txt3']}; }}
QComboBox QAbstractItemView {{
    background-color: {_C['card']};
    border: 1px solid {_C['borderH']};
    color: {_C['txt']};
    selection-background-color: {_C['accent']};
    selection-color: #FFFFFF;
    outline: none;
    padding: 2px;
}}
QComboBox QAbstractItemView::item {{
    color: {_C['txt']};
    background-color: transparent;
    padding: 5px 10px;
    min-height: 22px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {_C['cardH']};
    color: {_C['txt']};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {_C['accent']};
    color: #FFFFFF;
}}
QCheckBox::indicator {{
    width: 15px; height: 15px;
    border: 1px solid {_C['border']};
    background: {_C['input']};
    border-radius: 4px;
}}
QCheckBox::indicator:checked {{
    background: {_C['accent']};
    border-color: {_C['accent']};
}}

/* ── Buttons ── */
QPushButton {{
    background: {_C['card']};
    border: 1px solid {_C['border']};
    border-radius: 7px;
    padding: 6px 14px;
    color: {_C['txt2']};
    font-weight: 500;
}}
QPushButton:hover {{
    background: {_C['cardH']};
    border-color: {_C['borderH']};
    color: {_C['txt']};
}}
QPushButton:pressed {{ background: rgba(99,102,241,0.15); }}
QPushButton[role="primary"] {{
    background: {_C['accent']};
    border-color: {_C['accent']};
    color: white;
    font-weight: 600;
}}
QPushButton[role="primary"]:hover {{ background: #5558E0; border-color: #5558E0; }}
QPushButton[role="danger"] {{
    color: {_C['red']};
    border-color: rgba(239,68,68,0.35);
    background: transparent;
}}
QPushButton[role="danger"]:hover {{
    background: rgba(239,68,68,0.12);
    border-color: {_C['red']};
    color: {_C['red']};
}}
QPushButton[role="success"] {{
    background: {_C['green']};
    border-color: {_C['green']};
    color: white;
    font-weight: 600;
}}
QPushButton[role="success"]:hover {{ background: #0ea371; }}
QPushButton[role="ghost"] {{
    background: transparent;
    border: none;
    color: {_C['txt2']};
    padding: 4px 10px;
    border-radius: 6px;
}}
QPushButton[role="ghost"]:hover {{
    background: rgba(255,255,255,0.06);
    color: {_C['txt']};
}}

/* ── Table ── */
QTableView {{
    background: {_C['panel']};
    gridline-color: {_C['border']};
    border: none; outline: none;
    selection-background-color: rgba(99,102,241,0.2);
    color: {_C['txt']};
}}
QTableView::item:hover:!selected {{ background: rgba(255,255,255,0.04); }}
QHeaderView::section {{
    background: {_C['sidebar']};
    color: {_C['txt2']};
    padding: 5px 8px;
    border: none;
    border-right: 1px solid {_C['border']};
    border-bottom: 1px solid {_C['border']};
    font-size: 11px;
    font-weight: 600;
}}
QHeaderView::section:hover {{ color: {_C['txt']}; }}
QHeaderView::section:pressed {{ background: rgba(99,102,241,0.15); color: {_C['accent']}; }}

/* ── Scrollbars ── */
QScrollBar:vertical   {{ background: transparent; width: 8px; border: none; margin: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 8px; border: none; margin: 0; }}
QScrollBar::handle:vertical   {{ background: rgba(255,255,255,0.12); min-height: 24px; border-radius: 4px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: rgba(255,255,255,0.12); min-width: 24px;  border-radius: 4px; margin: 2px; }}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{ background: rgba(255,255,255,0.25); }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

/* ── Menus ── */
QMenu {{
    background: {_C['panel']};
    border: 1px solid {_C['border']};
    border-radius: 8px;
    padding: 4px;
}}
QMenu::item {{ padding: 6px 20px; border-radius: 5px; color: {_C['txt2']}; }}
QMenu::item:selected {{ background: rgba(99,102,241,0.15); color: {_C['txt']}; }}
QMenu::separator {{ height: 1px; background: {_C['border']}; margin: 3px 8px; }}

/* ── Dialogs ── */
QDialog {{ background: {_C['panel']}; }}
QScrollArea {{ border: none; background: transparent; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QPushButton::menu-indicator {{ image: none; width: 0; }}
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

# ── Tabler icon font (bundled) ──────────────────────────────────────────────
_TI_HEX = {
    "plus": "eb0b", "circle-plus": "ea69", "copy": "ea7a", "trash": "eb41",
    "table": "eba1", "columns": "eb83", "arrow-up": "ea25",
    "arrow-down": "ea16", "chevron-down": "ea5f", "folder-plus": "eaab",
    "pencil": "eb04", "note": "eb6d", "clipboard-plus": "efb2",
}
_TI_CODES = {k: chr(int(v, 16)) for k, v in _TI_HEX.items()}
_ti_family = None
_ti_cache = {}

def _resource_path(rel: str) -> str:
    bases = []
    try:
        bases.append(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        pass
    bases.append(os.path.dirname(os.path.abspath(sys.argv[0])))
    bases.append(os.getcwd())
    for b in bases:
        p = os.path.join(b, rel)
        if os.path.exists(p):
            return p
    return os.path.join(bases[0], rel)

def _ti_icon(name: str, color: str = None, px: int = 18) -> QIcon:
    """Render a Tabler glyph into a QIcon so the label keeps the normal font."""
    global _ti_family
    if _ti_family is None:
        fid = QFontDatabase.addApplicationFont(_resource_path("assets/tabler-icons.ttf"))
        fams = QFontDatabase.applicationFontFamilies(fid) if fid != -1 else []
        _ti_family = fams[0] if fams else ""
    color = color or _C["txt2"]
    ch = _TI_CODES.get(name)
    if not ch or not _ti_family:
        return QIcon()
    key = (name, color, px)
    if key in _ti_cache:
        return _ti_cache[key]
    pm = QPixmap(px, px)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    f = QFont(_ti_family)
    f.setPixelSize(px)
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(pm.rect(), Qt.AlignCenter, ch)
    p.end()
    ic = QIcon(pm)
    _ti_cache[key] = ic
    return ic

_ROLE_ICON_COLOR = {"success": "#10B981", "danger": "#EF4444", "primary": "white"}

def _mk_btn(text: str, role: str = "", icon: str = None,
            icon_color: str = None) -> QPushButton:
    b = QPushButton(text)
    if role:
        b.setProperty("role", role)
    if icon:
        b.setIcon(_ti_icon(icon, icon_color or _ROLE_ICON_COLOR.get(role)))
    return b


def _hsep() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setFixedHeight(1)
    f.setStyleSheet(f"background: {_C['border']}; border: none;")
    return f


def _vsep() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.VLine)
    f.setFixedWidth(1)
    f.setStyleSheet(f"background: {_C['border']}; border: none;")
    return f


def _sec_lbl(text: str) -> QLabel:
    lbl = QLabel(text.upper())
    lbl.setStyleSheet(
        f"color: {_C['txt3']}; font-size: 10px; font-weight: 600; "
        f"letter-spacing: 1px; padding: 10px 12px 5px; background: transparent;"
    )
    return lbl


def _json_highlight(raw: str) -> str:
    """JSON string → HTML with syntax highlighting."""
    import html as _html
    t = _html.escape(raw)
    t = re.sub(r'"([^"]+)":', f'<span style="color:#93C5FD">"\\1"</span>:', t)
    t = re.sub(r': "([^"]*)"', f': <span style="color:#86EFAC">"\\1"</span>', t)
    t = re.sub(r': (-?\d+\.?\d*)', f': <span style="color:#FBBF24">\\1</span>', t)
    t = re.sub(r'\b(true|false|null)\b', f'<span style="color:#F9A8D4">\\1</span>', t)
    return t


# ── Background workers ────────────────────────────────────────────────────────

class _LoadWorker(QThread):
    done  = Signal()
    error = Signal(str)

    def __init__(self, manager, path):
        super().__init__()
        self._manager = manager
        self._path    = path

    def run(self):
        try:
            self._manager.load_json(self._path)
            self.done.emit()
        except Exception as e:
            self.error.emit(str(e))


class _SaveWorker(QThread):
    done  = Signal()
    error = Signal(str)

    def __init__(self, manager):
        super().__init__()
        self._manager = manager

    def run(self):
        try:
            self._manager.save_json()
            self.done.emit()
        except Exception as e:
            self.error.emit(str(e))


# ── Image thumbnail helper ────────────────────────────────────────────────────

def _update_img_thumb(path_str: str, label: "QLabel", base_dir: "str | None") -> None:
    """Load an image into a QLabel thumbnail, resolving relative paths."""
    from PySide6.QtGui import QPixmap
    if not path_str:
        label.setPixmap(QPixmap())
        label.setText("No Image")
        return
    p = path_str if os.path.isabs(path_str) else (
        os.path.join(base_dir, path_str) if base_dir else path_str
    )
    p = os.path.normpath(p)
    px = QPixmap(p)
    if px.isNull():
        label.setPixmap(QPixmap())
        label.setText(f"找不到圖片:\n{p}")
        label.setWordWrap(True)
    else:
        label.setText("")
        w = label.width() or 220
        h = label.height() or 90
        label.setPixmap(px.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation))


# ── Non-scroll ComboBox (used in config dialog) ───────────────────────────────

class _NoscrollCombo(QComboBox):
    """ComboBox that ignores scroll wheel unless the user explicitly clicked into it."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setFocusPolicy(Qt.StrongFocus)   # only gain focus via click/tab, NOT wheel
        # Force dropdown popup colours directly on the internal QListView;
        # app-level QSS does not always cascade into the floating popup on Windows.
        self.view().setStyleSheet(
            f"QListView {{"
            f"  background-color: {_C['card']};"
            f"  color: {_C['txt']};"
            f"  border: 1px solid {_C['borderH']};"
            f"  outline: none;"
            f"}}"
            f"QListView::item {{"
            f"  color: {_C['txt']};"
            f"  background-color: transparent;"
            f"  padding: 4px 10px;"
            f"  min-height: 22px;"
            f"}}"
            f"QListView::item:hover {{"
            f"  background-color: {_C['cardH']};"
            f"  color: {_C['txt']};"
            f"}}"
            f"QListView::item:selected {{"
            f"  background-color: {_C['accent']};"
            f"  color: #FFFFFF;"
            f"}}"
        )

    def wheelEvent(self, e):
        if self.hasFocus():
            super().wheelEvent(e)
        else:
            e.ignore()


# ── ItemCardDelegate ──────────────────────────────────────────────────────────

class ItemCardDelegate(QStyledItemDelegate):
    """Draws each list item as a card with left category-color strip."""

    CARD_H  = 66   # card height
    PAD_V   = 4    # vertical outer padding (space between cards)
    PAD_H   = 10   # horizontal outer padding
    STRIP_W = 4    # left color strip width
    PAD_IN  = 12   # internal padding after strip

    R_PK  = Qt.UserRole + 1   # primary key string
    R_SUB = Qt.UserRole + 2   # subtitle string
    R_CAT = Qt.UserRole + 3   # category value (for color)

    def sizeHint(self, option, index):
        return QSize(option.rect.width(), self.CARD_H + self.PAD_V * 2)

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        pk_val   = index.data(self.R_PK)  or ""
        subtitle = index.data(self.R_SUB) or ""
        cat_val  = index.data(self.R_CAT) or ""
        cat      = _cat_for(cat_val) if cat_val else _CAT[0]
        cat_qc   = QColor(cat["color"])
        cat_bg   = QColor(cat["r"], cat["g"], cat["b"], 38)   # ~15% opacity

        selected = bool(option.state & QStyle.State_Selected)
        hovered  = bool(option.state & QStyle.State_MouseOver)

        # Card rect (with outer padding)
        full = option.rect
        card = QRectF(
            full.left()   + self.PAD_H,
            full.top()    + self.PAD_V,
            full.width()  - self.PAD_H * 2,
            full.height() - self.PAD_V * 2,
        )

        # ── Card background ──
        if selected:
            grad = QLinearGradient(card.left(), 0, card.right(), 0)
            grad.setColorAt(0.0, cat_bg)
            grad.setColorAt(0.4, QColor(_C["card"]))
            painter.setBrush(QBrush(grad))
            border_c = QColor(cat["r"], cat["g"], cat["b"], 90)
        elif hovered:
            painter.setBrush(QBrush(QColor(_C["cardH"])))
            border_c = QColor(_C["borderH"])
        else:
            painter.setBrush(QBrush(QColor(_C["card"])))
            border_c = QColor(_C["border"])

        painter.setPen(QPen(border_c, 1))
        painter.drawRoundedRect(card, 10, 10)

        # ── Left color strip (clipped to card shape) ──
        clip = QPainterPath()
        clip.addRoundedRect(card, 10, 10)
        painter.setClipPath(clip)
        painter.fillRect(
            QRectF(card.left(), card.top(), self.STRIP_W, card.height()),
            cat_qc
        )
        painter.setClipping(False)

        # ── Text ──
        tx = int(card.left()) + self.STRIP_W + self.PAD_IN
        tw = int(card.width()) - self.STRIP_W - self.PAD_IN * 2 - 22
        ct = int(card.top())
        ch = int(card.height())

        # ID — monospace, bold, category text color
        id_font = QFont("Consolas", 10, QFont.Bold)
        painter.setFont(id_font)
        painter.setPen(QColor(cat["text"]))
        painter.drawText(QRect(tx, ct + 9, tw, 22), Qt.AlignLeft | Qt.AlignVCenter, pk_val)

        # Subtitle — regular, muted
        sub_font = QFont("Segoe UI", 9)
        painter.setFont(sub_font)
        painter.setPen(QColor(_C["txt2"]))
        fm  = painter.fontMetrics()
        sub = fm.elidedText(str(subtitle), Qt.ElideRight, tw)
        painter.drawText(QRect(tx, ct + 33, tw, 20), Qt.AlignLeft | Qt.AlignVCenter, sub)

        # Right arrow
        arr_font = QFont("Segoe UI", 13)
        painter.setFont(arr_font)
        painter.setPen(QColor(cat["color"]) if selected else QColor(_C["txt3"]))
        painter.drawText(
            QRect(int(card.right()) - 22, ct, 20, ch),
            Qt.AlignCenter, "›"
        )

        painter.restore()


# ── SubTableModel ─────────────────────────────────────────────────────────────

_DIRTY_BG = QColor(234, 179, 8, 30)
_ROW_EVEN = QColor(0x17, 0x17, 0x2A)
_ROW_ODD  = QColor(0x1E, 0x1E, 0x32)


class SubTableModel(QAbstractTableModel):
    def __init__(self, df, cols_cfg, manager, sheet_full_name):
        super().__init__()
        self._df           = df if df is not None else pd.DataFrame()
        self._cols_cfg     = cols_cfg or {}
        self._manager      = manager
        self._sheet        = sheet_full_name

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c    = index.row(), index.column()
        col     = self._df.columns[c]
        col_type = self._cols_cfg.get(col, {}).get("type", "string")
        val     = self._df.iat[r, c]
        row_idx = self._df.index[r]

        if role in (Qt.DisplayRole, Qt.EditRole):
            return None if col_type == "bool" else (str(val) if val is not None else "")
        if role == Qt.CheckStateRole and col_type == "bool":
            v = val
            if isinstance(v, str):
                v = v.lower() in ("true", "1", "yes")
            return Qt.Checked if v else Qt.Unchecked
        if role == Qt.BackgroundRole:
            if (self._sheet, row_idx, col) in self._manager.dirty_cells:
                return QBrush(_DIRTY_BG)
            return QBrush(_ROW_EVEN if r % 2 == 0 else _ROW_ODD)
        if role == Qt.ForegroundRole:
            return QBrush(QColor(_C["txt"]))
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid():
            return False
        r, c    = index.row(), index.column()
        col     = self._df.columns[c]
        row_idx = self._df.index[r]
        if role == Qt.CheckStateRole:
            value = (value == Qt.Checked)
        self._manager.update_cell(self._sheet, row_idx, col, value)
        # Reflect the change in this filtered view copy at the SAME index label.
        # Do NOT swap in the full sub_table — that desyncs view rows from df rows
        # and routes later edits to whichever skill sits at low global indices.
        full = self._manager.sub_tables.get(self._sheet)
        if full is not None and row_idx in full.index and col in full.columns \
                and row_idx in self._df.index:
            self._df.at[row_idx, col] = full.at[row_idx, col]
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        col      = self._df.columns[index.column()]
        col_type = self._cols_cfg.get(col, {}).get("type", "string")
        base     = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        return base | (Qt.ItemIsUserCheckable if col_type == "bool" else Qt.ItemIsEditable)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal:
            if section < 0 or section >= len(self._df.columns):
                return None
            col = str(self._df.columns[section])
            if role == Qt.DisplayRole:
                return col
            if role == Qt.ToolTipRole:
                return self._cols_cfg.get(col, {}).get("note", "") or None
            return None
        if role == Qt.DisplayRole:
            return str(section + 1)
        return None

    def reload(self, df, cols_cfg=None):
        self.beginResetModel()
        self._df = df if df is not None else pd.DataFrame()
        if cols_cfg is not None:
            self._cols_cfg = cols_cfg
        self.endResetModel()

    def sort(self, column, order=Qt.AscendingOrder):
        if column < 0 or column >= len(self._df.columns):
            return
        col_name = self._df.columns[column]
        self.layoutAboutToBeChanged.emit()
        try:
            self._df = self._df.sort_values(
                col_name, ascending=(order == Qt.AscendingOrder), kind="mergesort"
            )
        except Exception:
            pass
        self.layoutChanged.emit()

    def df_index(self, view_row):
        if 0 <= view_row < len(self._df):
            return self._df.index[view_row]
        return None

    @property
    def df(self):
        return self._df


# ── EnumDelegate ──────────────────────────────────────────────────────────────

class EnumDelegate(QStyledItemDelegate):
    def __init__(self, options, parent=None):
        super().__init__(parent)
        self._options = [str(o) for o in options if str(o) != ""]

    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.addItem("")              # 允許清空 / 不選 enum 值
        cb.addItems(self._options)
        return cb

    def setEditorData(self, editor, index):
        val = index.data(Qt.DisplayRole) or ""
        idx = editor.findText(val)
        if idx >= 0:
            editor.setCurrentIndex(idx)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


# ── Suggest field: context-filtered autocomplete ──────────────────────────────

def _get_suggestions(df, this_col, context_col, context_value):
    """Sorted, deduped previously-used values of `this_col` from `df`,
       filtered by `df[context_col] == context_value` when context_col is set."""
    if df is None or this_col not in df.columns:
        return []
    if context_col and context_col in df.columns:
        mask = df[context_col].astype(str) == str(context_value)
        vals = df.loc[mask, this_col]
    else:
        vals = df[this_col]
    seen, out = set(), []
    for v in vals.astype(str):
        v = v.strip()
        if v and v not in seen:
            seen.add(v); out.append(v)
    return sorted(out)


class _SuggestLineEdit(QLineEdit):
    """QLineEdit that auto-pops its completer on focus-in."""
    def focusInEvent(self, e):
        super().focusInEvent(e)
        if self.completer():
            QTimer.singleShot(0, self.completer().complete)


class SuggestDelegate(QStyledItemDelegate):
    """Sub-table cell editor: line edit + popup of previously-used values,
       filtered by a sibling 'context' column on the same row."""
    def __init__(self, df_provider, this_col, context_col, parent=None):
        super().__init__(parent)
        self._df_provider  = df_provider
        self._this_col     = this_col
        self._context_col  = context_col

    def createEditor(self, parent, option, index):
        df = self._df_provider()
        ctx_val = ""
        try:
            ctx_val = str(index.model()._df.iloc[index.row()][self._context_col])
        except Exception:
            ctx_val = ""
        items = _get_suggestions(df, self._this_col, self._context_col, ctx_val)
        editor = _SuggestLineEdit(parent)
        m = QStringListModel(items, editor)
        compl = QCompleter(m, editor)
        compl.setCompletionMode(QCompleter.PopupCompletion)
        compl.setCaseSensitivity(Qt.CaseInsensitive)
        editor.setCompleter(compl)
        return editor

    def setEditorData(self, editor, index):
        editor.setText(index.data(Qt.DisplayRole) or "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text(), Qt.EditRole)


# ── Array editor: chips widget ────────────────────────────────────────────────

class FlowLayout(QLayout):
    """Left-to-right wrapping layout used to lay out array chips."""

    def __init__(self, parent=None, spacing=4):
        super().__init__(parent)
        self._items = []
        self.setSpacing(spacing)
        self.setContentsMargins(3, 3, 3, 3)

    def addItem(self, item):       self._items.append(item)
    def count(self):               return len(self._items)
    def itemAt(self, i):           return self._items[i] if 0 <= i < len(self._items) else None
    def takeAt(self, i):           return self._items.pop(i) if 0 <= i < len(self._items) else None
    def expandingDirections(self): return Qt.Orientation(0)
    def hasHeightForWidth(self):   return True
    def heightForWidth(self, w):   return self._arrange(QRect(0, 0, w, 0), apply=False)
    def sizeHint(self):            return self.minimumSize()

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._arrange(rect, apply=True)

    def minimumSize(self):
        s = QSize()
        for it in self._items:
            s = s.expandedTo(it.minimumSize())
        m = self.contentsMargins()
        return s + QSize(m.left() + m.right(), m.top() + m.bottom())

    def _arrange(self, rect, apply):
        m = self.contentsMargins()
        x0 = rect.x() + m.left()
        x, y = x0, rect.y() + m.top()
        right = rect.right() - m.right()
        line_h = 0
        sp = self.spacing()
        for it in self._items:
            sz = it.sizeHint()
            if x > x0 and x + sz.width() > right:
                x = x0
                y += line_h + sp
                line_h = 0
            if apply:
                it.setGeometry(QRect(QPoint(x, y), sz))
            x += sz.width() + sp
            line_h = max(line_h, sz.height())
        return y + line_h + m.bottom() - rect.y()


class _Chip(QFrame):
    """One array element — double-click to edit, ✕ to remove."""
    removed = Signal(object)
    edited  = Signal()

    def __init__(self, text, parent=None):
        super().__init__(parent)
        self._text = text
        self._editing = False
        self.setObjectName("arrayChip")
        self.setStyleSheet(
            f"QFrame#arrayChip {{ background:{_C['card']}; "
            f"border:1px solid {_C['border']}; border-radius:9px; }}"
        )
        self.setToolTip("雙擊可編輯")
        lo = QHBoxLayout(self)
        lo.setContentsMargins(8, 2, 4, 2)
        lo.setSpacing(3)

        self._lbl = QLabel(text if text != "" else "(空)")
        self._lbl.setStyleSheet(
            f"color:{_C['txt']}; background:transparent; border:none; font-size:12px;"
        )
        self._edit = QLineEdit(text)
        self._edit.setStyleSheet(
            f"color:{_C['txt']}; background:transparent; border:none; font-size:12px;"
        )
        self._edit.hide()
        self._edit.editingFinished.connect(self._commit_edit)
        self._edit.textChanged.connect(self._resize_edit)

        btn = QPushButton("✕")
        btn.setFixedSize(18, 18)
        btn.setAutoDefault(False); btn.setDefault(False)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(
            f"QPushButton {{ background:transparent; border:none; padding:0px; "
            f"margin:0px; color:{_C['txt3']}; font-size:13px; }}"
            f"QPushButton:hover {{ color:{_C['red']}; }}"
        )
        btn.clicked.connect(lambda: self.removed.emit(self))

        lo.addWidget(self._lbl)
        lo.addWidget(self._edit)
        lo.addWidget(btn)

    def mouseDoubleClickEvent(self, event):
        self._start_edit()

    def _start_edit(self):
        self._editing = True
        self._edit.setText(self._text)
        self._resize_edit()
        self._lbl.hide()
        self._edit.show()
        self._edit.setFocus()
        self._edit.selectAll()
        self.updateGeometry()

    def _resize_edit(self):
        fm = self._edit.fontMetrics()
        self._edit.setFixedWidth(max(fm.horizontalAdvance(self._edit.text()) + 8, 24))

    def _commit_edit(self):
        if not self._editing:
            return
        self._editing = False
        new = self._edit.text().strip()
        self._edit.hide()
        if new == "":
            self.removed.emit(self)
            return
        changed = (new != self._text)
        self._text = new
        self._lbl.setText(new)
        self._lbl.show()
        self.updateGeometry()
        if changed:
            self.edited.emit()

    def text(self):
        return self._text


class ChipsEdit(QWidget):
    """Array field editor: wrapping removable chips + add-input + copy button."""
    changed = Signal()

    def __init__(self, parent=None, chip_area_height=64):
        super().__init__(parent)
        self._chips = []
        lo = QVBoxLayout(self)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.setSpacing(4)

        self._chip_box = QWidget()
        sp = self._chip_box.sizePolicy()
        sp.setHeightForWidth(True)
        self._chip_box.setSizePolicy(sp)
        self._chip_box.setStyleSheet(
            f"background:{_C['code']}; border:1px solid {_C['border']}; border-radius:6px;"
        )
        self._flow = FlowLayout(self._chip_box, spacing=4)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(chip_area_height)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea{border:none; background:transparent;}")
        scroll.setWidget(self._chip_box)
        lo.addWidget(scroll)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)
        self._input = QLineEdit()
        self._input.setPlaceholderText("輸入後按 Enter 或逗號新增…")
        self._input.returnPressed.connect(self._commit_input)
        self._input.textChanged.connect(self._on_input_text)
        copy_btn = QPushButton("⧉ 複製")
        copy_btn.setAutoDefault(False); copy_btn.setDefault(False)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setToolTip("複製整個陣列內容（以逗號分隔）")
        copy_btn.setStyleSheet(
            f"background:{_C['card']}; border:1px solid {_C['border']}; "
            f"color:{_C['txt2']}; border-radius:5px; padding:3px 8px; font-size:11px;"
        )
        copy_btn.clicked.connect(self._copy_all)
        row.addWidget(self._input, 1)
        row.addWidget(copy_btn)
        lo.addLayout(row)

    # ── public API ──
    def set_value(self, comma_str):
        """Populate chips from a comma-separated string (does not emit `changed`)."""
        for ch in self._chips:
            self._flow.removeWidget(ch)
            ch.setParent(None)
            ch.deleteLater()
        self._chips = []
        for tok in str(comma_str or "").split(","):
            tok = tok.strip()
            if tok != "":
                self._add_chip(tok)
        self._chip_box.updateGeometry()

    def value(self):
        return ", ".join(c.text() for c in self._chips)

    # ── internal ──
    def _add_chip(self, text):
        chip = _Chip(text, self._chip_box)
        chip.removed.connect(self._remove_chip)
        chip.edited.connect(self.changed.emit)
        self._flow.addWidget(chip)
        self._chips.append(chip)
        chip.show()

    def _remove_chip(self, chip):
        if chip in self._chips:
            self._chips.remove(chip)
        self._flow.removeWidget(chip)
        chip.setParent(None)
        chip.deleteLater()
        self._chip_box.updateGeometry()
        self.changed.emit()

    def _commit_input(self):
        raw = self._input.text()
        self._input.blockSignals(True)
        self._input.clear()
        self._input.blockSignals(False)
        added = False
        for tok in raw.split(","):
            tok = tok.strip()
            if tok != "":
                self._add_chip(tok)
                added = True
        if added:
            self._chip_box.updateGeometry()
            self.changed.emit()

    def _on_input_text(self, text):
        if "," in text:
            self._commit_input()

    def _copy_all(self):
        QApplication.clipboard().setText(self.value())


class ArrayEditDialog(QDialog):
    """Popup chips editor for array cells in sub-tables."""

    def __init__(self, comma_str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("編輯陣列")
        self.setMinimumWidth(380)
        self.setStyleSheet(APP_QSS)
        lo = QVBoxLayout(self)
        self._chips = ChipsEdit(chip_area_height=150)
        self._chips.set_value(comma_str)
        lo.addWidget(self._chips)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        for _b in bb.buttons():
            _b.setAutoDefault(False); _b.setDefault(False)
        lo.addWidget(bb)

    def keyPressEvent(self, event):
        # Enter is for committing/adding an item (handled by the child QLineEdits'
        # returnPressed/editingFinished, which fire first). Swallow it here so it
        # doesn't trigger a default button (close dialog / delete chip). OK closes.
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            event.accept()
            return
        super().keyPressEvent(event)

    def value(self):
        return self._chips.value()


class ArrayDelegate(QStyledItemDelegate):
    """Sub-table array cell — double-click opens the chips editor dialog."""

    def createEditor(self, parent, option, index):
        cur = index.data(Qt.DisplayRole) or ""
        dlg = ArrayEditDialog(cur, parent)
        if dlg.exec() == QDialog.Accepted:
            index.model().setData(index, dlg.value(), Qt.EditRole)
        return None


# ── FieldEditorWidget ─────────────────────────────────────────────────────────

class FieldEditorWidget(QWidget):
    """Build-once field form with indigo ● dots and dirty-state highlighting."""
    field_changed = Signal(str, object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._widgets       = {}
        self._col_types     = {}
        self._lbl_widgets   = {}   # col → QLabel (the col name label)
        self._bool_updaters = {}   # col → update_style(checked: bool)
        self._img_preview_label:  "QLabel | None" = None  # table-level image preview
        self._img_path_segments:  list = []  # [{"type":"col","col":"X"} | {"type":"lit","value":"Y"}]
        self._img_base_folder:    str  = ""  # configured base folder for images
        self._img_ext:            str  = ""  # file extension appended to assembled path e.g. ".png"
        self._text_ref_json:    str = ""   # table-level external text-ref JSON path
        self._text_ref_key_col: str = "TextID"
        self._text_ref_val_col: str = "TextContent"
        self._ref_labels    = {}   # col → QLabel (resolved lookup text for "text_ref" type)
        self._row_idx     = None
        self._table_name  = None
        self._manager     = None
        self._built       = False

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background: transparent; border: none;")

        self._content = QWidget()
        self._content.setStyleSheet(f"background: {_C['panel']};")
        self._form_lo = QVBoxLayout(self._content)
        self._form_lo.setContentsMargins(0, 4, 0, 12)
        self._form_lo.setSpacing(0)
        scroll.setWidget(self._content)

        lo = QVBoxLayout(self)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.addWidget(scroll)

    def build_for(self, df, cfg, table_name, manager):
        # Clear
        while self._form_lo.count():
            item = self._form_lo.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._widgets.clear()
        self._col_types.clear()
        self._lbl_widgets.clear()
        self._bool_updaters.clear()
        self._img_preview_label  = None  # cleared by layout wipe above
        self._img_path_segments  = []
        self._img_base_folder    = ""
        self._img_ext            = ""
        self._text_ref_json     = ""
        self._text_ref_key_col  = "TextID"
        self._text_ref_val_col  = "TextContent"
        self._ref_labels.clear()
        self._table_name = table_name
        self._manager    = manager
        cols_cfg  = cfg.get("columns", {})
        _img_cfg  = cfg.get("image_preview", {})
        self._img_base_folder  = _img_cfg.get("base_folder", "")
        self._img_ext          = _img_cfg.get("ext", "")
        _segs = _img_cfg.get("path_segments", [])
        if not _segs:
            # backward compat: old single "col" key
            _old_col = _img_cfg.get("col", "")
            if _old_col and _old_col in df.columns:
                _segs = [{"type": "col", "col": _old_col}]
        self._img_path_segments = _segs
        _trs = cfg.get("text_ref_source", {})
        self._text_ref_json    = _trs.get("json_path", "")
        self._text_ref_key_col = _trs.get("key_col", "TextID")   or "TextID"
        self._text_ref_val_col = _trs.get("val_col", "TextContent") or "TextContent"

        # ── Table-level image preview (shown at top if configured) ─────────────
        if self._img_path_segments:
            _seg_desc = "/".join(
                s.get("col", "?") if s.get("type") == "col" else s.get("value", "?")
                for s in self._img_path_segments
            )
            prev_card = QWidget()
            prev_card.setStyleSheet(f"background:{_C['panel']};")
            pclo = QVBoxLayout(prev_card)
            pclo.setContentsMargins(14, 10, 14, 6); pclo.setSpacing(4)
            lbl_img = QLabel(f"● IMAGE  [{_seg_desc}]")
            lbl_img.setStyleSheet(
                f"color:{_C['txt3']}; font-size:10px; font-weight:600; "
                f"letter-spacing:1px; background:transparent;"
            )
            self._img_preview_label = QLabel("No Image")
            self._img_preview_label.setAlignment(Qt.AlignCenter)
            self._img_preview_label.setFixedHeight(160)
            self._img_preview_label.setStyleSheet(
                f"background:{_C['code']}; border:1px solid {_C['border']}; "
                f"border-radius:6px; color:{_C['txt3']}; font-size:12px;"
            )
            pclo.addWidget(lbl_img)
            pclo.addWidget(self._img_preview_label)
            sep_img = QFrame(); sep_img.setFixedHeight(1)
            sep_img.setStyleSheet(f"background:{_C['border']}; border:none;")
            self._form_lo.addWidget(prev_card)
            self._form_lo.addWidget(sep_img)

        for col in df.columns:
            col_conf = cols_cfg.get(col, {})
            col_type = col_conf.get("type", "string")
            self._col_types[col] = col_type

            # ── Field group ──
            grp = QWidget()
            grp.setStyleSheet(f"background: {_C['panel']};")
            glo = QVBoxLayout(grp)
            glo.setContentsMargins(14, 8, 14, 4)
            glo.setSpacing(4)

            # Label row: ● ColName
            lbl_row = QWidget()
            lbl_row.setStyleSheet("background: transparent;")
            lrlo = QHBoxLayout(lbl_row)
            lrlo.setContentsMargins(0, 0, 0, 0)
            lrlo.setSpacing(5)

            dot = QLabel("●")
            dot.setStyleSheet(
                f"color: {_C['accent']}; font-size: 11px; background: transparent;"
            )
            lbl = QLabel(col)
            lbl.setStyleSheet(
                f"color: {_C['txt2']}; font-size: 11px; font-weight: 500; background: transparent;"
            )
            self._lbl_widgets[col] = lbl
            _note = col_conf.get("note", "")
            if _note:
                lbl.setToolTip(_note)
                lbl_row.setToolTip(_note)
            lrlo.addWidget(dot)
            lrlo.addWidget(lbl)
            lrlo.addStretch()
            glo.addWidget(lbl_row)

            # ── Input widget ──
            if col_type == "bool":
                w = QPushButton()
                w.setCheckable(True)
                w.setFixedHeight(36)

                def _make_bool_style(btn):
                    def _upd(checked):
                        if checked:
                            btn.setText("  ✓   True")
                            btn.setStyleSheet(
                                f"background:rgba(16,185,129,0.18);"
                                f"border:1px solid rgba(16,185,129,0.65);"
                                f"color:{_C['green']}; border-radius:7px;"
                                f"font-size:13px; font-weight:600; text-align:left; padding:0 12px;"
                            )
                        else:
                            btn.setText("  ✗   False")
                            btn.setStyleSheet(
                                f"background:rgba(239,68,68,0.10);"
                                f"border:1px solid rgba(239,68,68,0.35);"
                                f"color:{_C['red']}; border-radius:7px;"
                                f"font-size:13px; font-weight:600; text-align:left; padding:0 12px;"
                            )
                    return _upd

                updater = _make_bool_style(w)
                self._bool_updaters[col] = updater
                updater(False)
                w.toggled.connect(updater)
                w.toggled.connect(lambda checked, c=col: self.field_changed.emit(c, checked))

            elif col_type == "enum":
                opts = [str(o) for o in (col_conf.get("options") or []) if str(o) != ""]
                w = _NoscrollCombo()
                w.addItem("")
                w.addItems(opts)
                w.currentTextChanged.connect(
                    lambda v, c=col: self.field_changed.emit(c, v)
                )

            elif col_type in ("int", "float"):
                w = QLineEdit()
                w.textChanged.connect(
                    lambda v, c=col, ct=col_type: self._on_numeric(c, v, ct)
                )

            elif col_type == "array":
                w = ChipsEdit()
                w.changed.connect(
                    lambda c=col, _w=w: self.field_changed.emit(c, _w.value())
                )

            elif col_type == "text_ref":
                # Editable string (same as string type) — value IS the lookup key
                w = QTextEdit()
                w.setMaximumHeight(76)
                w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

                # Read-only resolved text label (live-updating)
                ref_lbl = QLabel("—")
                ref_lbl.setWordWrap(True)
                ref_lbl.setStyleSheet(
                    f"color:#FFFFFF; font-size:13px; font-weight:700; font-style:normal; "
                    f"background:{_C['code']}; border-radius:4px; padding:4px 8px;"
                )
                self._ref_labels[col] = ref_lbl

                def _on_text_ref(c=col, widget=w, lbl=ref_lbl):
                    val = widget.toPlainText()
                    self.field_changed.emit(c, val)
                    self._update_ref_label(lbl, val)

                w.textChanged.connect(_on_text_ref)
                glo.addWidget(w)
                glo.addWidget(ref_lbl)
                self._widgets[col] = w
                sep = QFrame(); sep.setFixedHeight(1)
                sep.setStyleSheet(f"background: {_C['border']}; border: none;")
                self._form_lo.addWidget(grp)   # must parent grp or GC kills the child widgets
                self._form_lo.addWidget(sep)
                continue

            else:
                w = QTextEdit()
                w.setMaximumHeight(76)
                w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                w.textChanged.connect(
                    lambda c=col, widget=w: self.field_changed.emit(c, widget.toPlainText())
                )

            glo.addWidget(w)
            self._widgets[col] = w

            # Thin separator line below each field
            sep = QFrame()
            sep.setFixedHeight(1)
            sep.setStyleSheet(f"background: {_C['border']}; border: none;")
            self._form_lo.addWidget(grp)
            self._form_lo.addWidget(sep)

        self._form_lo.addStretch(1)
        self._built = True

    # ── Validation ────────────────────────────────────────────────────────────

    def _on_numeric(self, col, value, col_type):
        w = self._widgets.get(col)
        if w:
            valid = self._validate(value, col_type)
            w.setProperty("invalid", "true" if not valid else "false")
            w.style().unpolish(w)
            w.style().polish(w)
        self.field_changed.emit(col, value)

    @staticmethod
    def _validate(value, col_type):
        if col_type == "int":
            return value == "" or value.lstrip("-").isdigit()
        if col_type == "float":
            if value in ("", "-", "."): return True
            try: float(value); return True
            except ValueError: return False
        return True

    # ── Text-ref lookup ───────────────────────────────────────────────────────

    def _update_ref_label(self, lbl: "QLabel", key_val: str) -> None:
        if not self._text_ref_json:
            lbl.setText("（未設定外部文字表）")
            return
        if not self._manager:
            return
        json_dir = os.path.dirname(self._manager.json_path) if self._manager.json_path else ""
        abs_ref  = (os.path.join(json_dir, self._text_ref_json)
                    if json_dir and not os.path.isabs(self._text_ref_json)
                    else self._text_ref_json)
        resolved = self._manager.get_ref_text(
            abs_ref, self._text_ref_key_col, key_val, self._text_ref_val_col
        )
        lbl.setText(resolved if resolved else "（找不到對應文字）")

    # ── Load row ──────────────────────────────────────────────────────────────

    def load_row(self, row_data, row_idx):
        if not self._built:
            return
        self._row_idx = row_idx
        dirty = self._manager.dirty_cells if self._manager else set()

        for col, w in self._widgets.items():
            w.blockSignals(True)
            try:
                try:    val = row_data[col]
                except: val = ""

                col_type = self._col_types[col]
                is_dirty = (self._table_name, row_idx, col) in dirty

                # Dirty styling
                lbl = self._lbl_widgets.get(col)
                if is_dirty:
                    w.setStyleSheet(
                        "border-color: rgba(234,179,8,0.55); background: rgba(234,179,8,0.07);"
                    )
                    if lbl:
                        lbl.setStyleSheet(
                            f"color: {_C['yellow']}; font-size: 11px; font-weight: 500; background: transparent;"
                        )
                else:
                    w.setStyleSheet("")
                    if lbl:
                        lbl.setStyleSheet(
                            f"color: {_C['txt2']}; font-size: 11px; font-weight: 500; background: transparent;"
                        )

                # Value
                if col_type == "bool":
                    v = val
                    if isinstance(v, str):
                        v = v.lower() in ("true", "1", "yes")
                    checked = bool(v) if val != "" else False
                    w.setChecked(checked)
                    upd = self._bool_updaters.get(col)
                    if upd:
                        upd(checked)
                elif col_type == "enum":
                    w.setCurrentText(str(val) if val is not None else "")
                elif col_type in ("int", "float"):
                    w.setText(str(val) if val is not None else "")
                    w.setProperty("invalid", "false")
                    w.style().unpolish(w); w.style().polish(w)
                elif col_type == "array":
                    w.set_value(str(val) if val is not None else "")
                elif col_type == "text_ref":
                    val_str = str(val) if val is not None else ""
                    w.setPlainText(val_str)
                    ref_lbl = self._ref_labels.get(col)
                    if ref_lbl:
                        self._update_ref_label(ref_lbl, val_str)
                else:
                    w.setPlainText(str(val) if val is not None else "")

            finally:
                w.blockSignals(False)

        # Update table-level image preview
        if self._img_path_segments and self._img_preview_label:
            try:
                parts = []
                for _seg in self._img_path_segments:
                    if _seg.get("type") == "col":
                        _c = _seg.get("col", "")
                        parts.append(str(row_data[_c]) if _c and _c in row_data.index else "")
                    else:
                        parts.append(_seg.get("value", ""))
                img_val = "/".join(parts) + self._img_ext
            except Exception:
                img_val = ""
            # Resolve base: configured folder first, else JSON dir
            base = self._img_base_folder
            if base and not os.path.isabs(base) and self._manager and self._manager.json_path:
                base = os.path.join(os.path.dirname(self._manager.json_path), base)
            elif not base and self._manager and self._manager.json_path:
                base = os.path.dirname(self._manager.json_path)
            _update_img_thumb(img_val, self._img_preview_label, base)


# ── SubTablePanel ─────────────────────────────────────────────────────────────

class SubTablePanel(QWidget):
    row_deleted    = Signal(str, object)
    copy_requested = Signal()
    paste_requested = Signal()
    compare_requested = Signal(object)  # df_index of the right-clicked row

    def __init__(self, sheet_full_name, cols_cfg, manager, parent=None):
        super().__init__(parent)
        self._sheet   = sheet_full_name
        self._manager = manager

        _existing = manager.sub_tables.get(sheet_full_name)
        empty_df = pd.DataFrame(columns=list(
            _existing.columns if _existing is not None else []
        ))
        self._model = SubTableModel(empty_df, cols_cfg, manager, sheet_full_name)

        self._view = QTableView()
        self._view.setModel(self._model)
        self._view.setSortingEnabled(True)
        self._view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self._view.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self._view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._view.horizontalHeader().setStretchLastSection(True)
        self._view.verticalHeader().setDefaultSectionSize(28)
        self._view.verticalHeader().hide()
        self._view.setContextMenuPolicy(Qt.CustomContextMenu)
        self._view.customContextMenuRequested.connect(self._ctx_menu)
        self._view.keyPressEvent = self._key_press
        self._refresh_delegates(cols_cfg)

        lo = QVBoxLayout(self)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.addWidget(self._view)

    def _refresh_delegates(self, cols_cfg):
        for c, col in enumerate(self._model.df.columns):
            col_conf = (cols_cfg or {}).get(col, {})
            col_type = col_conf.get("type", "string")
            if col_type == "enum":
                opts = col_conf.get("options") or [""]
                self._view.setItemDelegateForColumn(c, EnumDelegate(opts, self._view))
            elif col_type == "array":
                self._view.setItemDelegateForColumn(c, ArrayDelegate(self._view))
            elif col_conf.get("suggest_from"):
                ctx = col_conf["suggest_from"]
                df_getter = (lambda sheet=self._sheet:
                             self._manager.sub_tables.get(sheet))
                self._view.setItemDelegateForColumn(c, SuggestDelegate(
                    df_getter, col, ctx, self._view
                ))

    def reload(self, df, cols_cfg=None):
        self._model.reload(df, cols_cfg)
        if cols_cfg:
            self._refresh_delegates(cols_cfg)
        if not df.empty:
            self._view.resizeColumnsToContents()

    def _ctx_menu(self, pos):
        menu = QMenu(self)
        menu.addAction("複製列", self.copy_requested.emit)
        menu.addAction("貼上列", self.paste_requested.emit)
        idx = self._view.indexAt(pos)
        if idx.isValid():
            di = self._model.df_index(idx.row())
            menu.addAction("對照相同效果…", lambda: self.compare_requested.emit(di))
            menu.addSeparator()
            menu.addAction("刪除此列", self._delete_selected)
        menu.exec(self._view.viewport().mapToGlobal(pos))

    def _key_press(self, event):
        if event.key() == Qt.Key_Delete:
            self._delete_selected()
        elif event.matches(QKeySequence.Copy):
            self.copy_requested.emit()
        elif event.matches(QKeySequence.Paste):
            self.paste_requested.emit()
        else:
            QTableView.keyPressEvent(self._view, event)

    def _delete_selected(self):
        sel = self._view.selectionModel().selectedRows()
        if not sel:
            return
        df_idx = self._model.df_index(sel[0].row())
        if df_idx is not None:
            self.row_deleted.emit(self._sheet, df_idx)

    def selected_df_index(self):
        sel = self._view.selectionModel().selectedRows()
        if not sel:
            return None
        return self._model.df_index(sel[0].row())

    def selected_df_indices(self):
        rows = sorted(i.row() for i in self._view.selectionModel().selectedRows())
        out = []
        for r in rows:
            di = self._model.df_index(r)
            if di is not None:
                out.append(di)
        return out


class EffectCompareDialog(QDialog):
    """Read-only cross-skill effect comparison: list every row in the sub-table
    that matches the seed row's key columns (InfluenceStatus / EffectReceive…)."""

    _FILTERS = [("InfluenceStatus", "影響屬性"), ("EffectReceive", "目標"),
                ("SkillComponentID", "組件")]
    _SHOW = ["SkillComponentID", "EffectValue", "InfluenceStatus",
             "EffectReceive", "AddType", "EffectDurationTime"]

    def __init__(self, sub_df, fk_key, seed, name_of, sub_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"對照相同效果 — {sub_name}")
        self.resize(700, 540)
        self.setStyleSheet(APP_QSS)
        self._df = sub_df
        self._fk = fk_key
        self._name_of = name_of
        self._cols = [c for c in self._SHOW if c in sub_df.columns]

        outer = QVBoxLayout(self)
        frow = QWidget(); frow.setStyleSheet("background:transparent;")
        fl = QHBoxLayout(frow); fl.setContentsMargins(10, 8, 10, 2); fl.setSpacing(12)
        fl.addWidget(QLabel("篩選："))
        self._filters = {}
        for col, label in self._FILTERS:
            if col not in sub_df.columns:
                continue
            val = str(seed.get(col, "")).strip()
            cb = QCheckBox(f"{label} =")
            cb.setChecked(col in ("InfluenceStatus", "EffectReceive") and val != "")
            cb.stateChanged.connect(self._rebuild)
            combo = _NoscrollCombo()
            combo.setMaximumWidth(150)
            vals = sorted({str(v).strip() for v in sub_df[col] if str(v).strip() != ""})
            combo.addItems(vals)
            if val in vals:
                combo.setCurrentText(val)
            combo.currentTextChanged.connect(self._rebuild)
            self._filters[col] = (cb, combo)
            fl.addWidget(cb); fl.addWidget(combo)
        fl.addStretch(1)
        outer.addWidget(frow)

        self._tbl = QTableWidget()
        self._tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self._tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self._tbl.setSortingEnabled(True)
        self._tbl.verticalHeader().setVisible(False)
        outer.addWidget(self._tbl, 1)

        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        outer.addWidget(bb)
        self._rebuild()

    def _rebuild(self):
        df = self._df
        mask = pd.Series([True] * len(df), index=df.index)
        for col, (cb, combo) in self._filters.items():
            if cb.isChecked():
                mask &= (df[col].astype(str).str.strip() == combo.currentText())
        sub = df[mask]
        headers = ["來源 ID", "名稱"] + self._cols
        self._tbl.setSortingEnabled(False)
        self._tbl.clear()
        self._tbl.setColumnCount(len(headers))
        self._tbl.setHorizontalHeaderLabels(headers)
        self._tbl.setRowCount(len(sub))
        for r, (_idx, row) in enumerate(sub.iterrows()):
            fkval = str(row[self._fk])
            cells = [fkval, self._name_of(fkval)] + [str(row[c]) for c in self._cols]
            for c, v in enumerate(cells):
                it = QTableWidgetItem(v)
                if headers[c] == "EffectValue":
                    try: it.setData(Qt.EditRole, float(v))
                    except (ValueError, TypeError): pass
                self._tbl.setItem(r, c, it)
        self._tbl.setSortingEnabled(True)
        self._tbl.resizeColumnsToContents()
        self.setWindowTitle(
            self.windowTitle().split("　(")[0] + f"　({len(sub)} 筆符合)")


# ── TableEditor ───────────────────────────────────────────────────────────────

class TableEditor(QWidget):
    status_message = Signal(str, str)
    _sub_clipboard = None  # {"sub": tab_name, "rows": [ {col: val, ...} ]} — shared across items/tables

    def __init__(self, table_name, manager, parent=None):
        super().__init__(parent)
        self.table_name = table_name
        self.manager    = manager
        self.df         = manager.tables[table_name]
        self.cfg        = manager.config.get(table_name, {})
        cols            = list(self.df.columns)
        self.cls_key    = self.cfg.get("classification_key", "") or (cols[0] if cols else "")
        self.pk_key     = self.cfg.get("primary_key",        "") or (cols[0] if cols else "")

        self.current_cls_val    = None
        self.current_master_idx = None
        self.current_master_pk  = None

        self._field_panel: FieldEditorWidget | None = None
        self._sub_panels:  dict[str, SubTablePanel] = {}
        self._sub_tab_order: list[str] = []

        self._setup_ui()
        self._load_cls_list()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ══ LEFT SIDEBAR — classification groups (200px) ══════════════════════
        left = QWidget()
        left.setFixedWidth(200)
        left.setStyleSheet(f"background: {_C['sidebar']};")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.setSpacing(0)

        lv.addWidget(_sec_lbl(f"Groups · {self.cls_key}"))

        self._cls_list = QListWidget()
        self._cls_list.setObjectName("cls-list")
        self._cls_list.setFocusPolicy(Qt.NoFocus)
        self._cls_list.currentItemChanged.connect(self._on_cls_changed)
        self._cls_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self._cls_list.customContextMenuRequested.connect(self._cls_ctx_menu)
        lv.addWidget(self._cls_list, 1)

        lv.addWidget(_hsep())

        cls_btns = QHBoxLayout()
        cls_btns.setContentsMargins(8, 6, 8, 6)
        cls_btns.setSpacing(4)
        b_add_cls = _mk_btn("分類", "ghost", icon="folder-plus", icon_color="#10B981"); b_add_cls.clicked.connect(self.add_classification)
        b_del_cls = _mk_btn("", "ghost", icon="trash", icon_color="#EF4444"); b_del_cls.setFixedWidth(30); b_del_cls.setToolTip("刪除分類"); b_del_cls.clicked.connect(self.delete_classification)
        b_up_cls  = _mk_btn("", "ghost", icon="arrow-up");   b_up_cls.setFixedWidth(30); b_up_cls.setToolTip("上移"); b_up_cls.clicked.connect(lambda: self.move_classification(-1))
        b_dn_cls  = _mk_btn("", "ghost", icon="arrow-down"); b_dn_cls.setFixedWidth(30); b_dn_cls.setToolTip("下移"); b_dn_cls.clicked.connect(lambda: self.move_classification(1))
        for b in [b_add_cls, b_del_cls, b_up_cls, b_dn_cls]:
            cls_btns.addWidget(b)
        lv.addLayout(cls_btns)

        root.addWidget(left)
        root.addWidget(_vsep())

        # ══ MIDDLE — search + card list (flex) ════════════════════════════════
        mid = QWidget()
        mid.setMinimumWidth(240)
        mid.setStyleSheet(f"background: {_C['bg']};")
        mv = QVBoxLayout(mid)
        mv.setContentsMargins(0, 0, 0, 0)
        mv.setSpacing(0)

        # Search + action bar
        action_bar = QWidget()
        action_bar.setStyleSheet(
            f"background: {_C['sidebar']}; border-bottom: 1px solid {_C['border']};"
        )
        action_bar.setFixedHeight(50)
        alo = QHBoxLayout(action_bar)
        alo.setContentsMargins(10, 0, 10, 0)
        alo.setSpacing(6)

        search_icon = QLabel("⌕")
        search_icon.setStyleSheet(f"color:{_C['txt2']}; background:transparent; font-size:16px;")
        self._filter_edit = QLineEdit()
        self._filter_edit.setPlaceholderText("搜尋項目…")
        self._filter_edit.setClearButtonEnabled(True)
        self._filter_edit.textChanged.connect(self._apply_filter)
        alo.addWidget(search_icon)
        alo.addWidget(self._filter_edit, 1)
        alo.addWidget(_vsep())

        for text, role, icon, tip, slot in [
            ("新增", "success", "circle-plus", "",   self.add_master_item),
            ("複製", "",        "copy",        "",   self.copy_master_item),
            ("",     "ghost",   "arrow-up",    "上移", lambda: self.move_master_item(-1)),
            ("",     "ghost",   "arrow-down",  "下移", lambda: self.move_master_item(1)),
            ("刪除", "danger",  "trash",       "",   self.delete_master_item),
        ]:
            b = _mk_btn(text, role, icon=icon)
            b.setFixedHeight(32)
            if not text:
                b.setFixedWidth(36); b.setToolTip(tip)
            b.clicked.connect(slot)
            alo.addWidget(b)

        mv.addWidget(action_bar)

        # Card list
        self._card_list = QListWidget()
        self._card_list.setObjectName("card-list")
        self._card_list.setItemDelegate(ItemCardDelegate(self._card_list))
        self._card_list.setMouseTracking(True)
        self._card_list.setFocusPolicy(Qt.NoFocus)
        self._card_list.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._card_list.currentItemChanged.connect(self._on_item_changed)
        self._card_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self._card_list.customContextMenuRequested.connect(self._item_ctx_menu)
        mv.addWidget(self._card_list, 1)

        # mid is placed into the draggable splitter below

        # ══ RIGHT — field editor + JSON preview + sub-tables ══════════════════
        right = QWidget()
        right.setMinimumWidth(320)
        right.setStyleSheet(f"background: {_C['panel']};")
        self._right_panel = right
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)

        # Panel header (item ID) + ⊞ 欄位 button
        _ph_widget = QWidget()
        _ph_widget.setFixedHeight(44)
        _ph_widget.setStyleSheet(
            f"background: {_C['panel']}; border-bottom: 1px solid {_C['border']};"
        )
        _ph_lo = QHBoxLayout(_ph_widget)
        _ph_lo.setContentsMargins(14, 0, 8, 0)
        _ph_lo.setSpacing(6)
        self._panel_hdr = QLabel("— 請選擇項目 —")
        self._panel_hdr.setStyleSheet(
            f"color: {_C['txt2']}; font-size: 13px; font-weight: 500; background: transparent;"
        )
        _col_btn = _mk_btn("欄位 ▾", "ghost", icon="columns")
        _col_btn.setFixedHeight(28)
        _col_menu = QMenu(_col_btn)
        _col_menu.addAction(_ti_icon("plus", "#10B981"), "新增欄位", self.add_master_column)
        _col_menu.addAction(_ti_icon("pencil", "#A5B4FC"), "重新命名", self.rename_master_column)
        _col_menu.addAction(_ti_icon("trash", "#EF4444"), "刪除欄位", self.delete_master_column)
        _col_btn.setMenu(_col_menu)
        _ph_lo.addWidget(self._panel_hdr, 1)
        _ph_lo.addWidget(_col_btn)
        rv.addWidget(_ph_widget)

        # Right splitter: field editor (top) / sub-tables (bottom)
        rsplit = QSplitter(Qt.Vertical)
        rsplit.setHandleWidth(2)
        rsplit.setStyleSheet(f"QSplitter::handle:vertical {{ background: {_C['border']}; }}")

        # ── Field + JSON area ──
        field_area = QWidget()
        field_area.setStyleSheet(f"background: {_C['panel']};")
        fv = QVBoxLayout(field_area)
        fv.setContentsMargins(0, 0, 0, 0)
        fv.setSpacing(0)

        self._field_container = QWidget()
        self._field_container.setStyleSheet(f"background: {_C['panel']};")
        fclo = QVBoxLayout(self._field_container)
        fclo.setContentsMargins(0, 0, 0, 0)
        fv.addWidget(self._field_container, 1)

        # JSON Preview (collapsible)
        self._json_toggle = QPushButton("▶  JSON PREVIEW")
        self._json_toggle.setStyleSheet(
            f"QPushButton {{ background:{_C['sidebar']}; border:none; "
            f"border-top:1px solid {_C['border']}; border-bottom:1px solid {_C['border']}; "
            f"color:{_C['txt3']}; font-size:10px; font-weight:600; letter-spacing:1px; "
            f"text-align:left; padding:7px 14px; }}"
            f"QPushButton:hover {{ background:{_C['cardH']}; color:{_C['txt2']}; }}"
        )
        self._json_toggle.clicked.connect(self._toggle_json)
        self._json_preview = QTextEdit()
        self._json_preview.setObjectName("code-view")
        self._json_preview.setReadOnly(True)
        self._json_preview.setMaximumHeight(180)
        self._json_preview.hide()
        fv.addWidget(self._json_toggle)
        fv.addWidget(self._json_preview)
        rsplit.addWidget(field_area)

        # ── Sub-tables area ──
        sub_area = QWidget()
        sub_area.setStyleSheet(f"background: {_C['panel']};")
        sv = QVBoxLayout(sub_area)
        sv.setContentsMargins(0, 0, 0, 0)
        sv.setSpacing(0)

        sub_hdr = QWidget()
        self._sub_hdr = sub_hdr
        sub_hdr.setFixedHeight(40)
        sub_hdr.setStyleSheet(
            f"background:{_C['sidebar']}; border-top:1px solid {_C['border']}; border-bottom:1px solid {_C['border']};"
        )
        sh = QHBoxLayout(sub_hdr)
        sh.setContentsMargins(12, 0, 8, 0)
        sh.setSpacing(4)
        lbl_sub = QLabel("SUB-TABLES")
        lbl_sub.setStyleSheet(
            f"color:{_C['txt3']}; font-size:10px; font-weight:600; letter-spacing:1px; background:transparent;"
        )
        sh.addWidget(lbl_sub)
        self._sub_picker = _NoscrollCombo()
        self._sub_picker.setFixedWidth(150)
        self._sub_picker.setToolTip("跳到子表")
        self._sub_picker.activated.connect(self._on_sub_picker)
        sh.addWidget(self._sub_picker)
        sh.addStretch(1)

        for text, role, icon, tip, slot in [
            ("新增列", "success", "circle-plus", "",   self.add_sub_item),
            ("複製",   "ghost",   "copy",        "原地複製一列", self.copy_sub_item),
            ("",       "ghost",   "arrow-up",    "上移", lambda: self.move_sub_item(-1)),
            ("",       "ghost",   "arrow-down",  "下移", lambda: self.move_sub_item(1)),
            ("刪除",   "danger",  "trash",       "",   self.delete_sub_item),
            ("複製列", "ghost",   "copy",        "複製選取列到剪貼簿（可跨筆/跨子表貼上）", self.copy_sub_rows),
            ("貼上列", "ghost",   "clipboard-plus", "把剪貼簿的列貼到目前子表", self.paste_sub_rows),
        ]:
            b = _mk_btn(text, role, icon=icon)
            b.setFixedHeight(28)
            if tip:
                b.setToolTip(tip)
            if not text:
                b.setFixedWidth(32)
            b.clicked.connect(slot)
            sh.addWidget(b)

        sh.addWidget(_vsep())

        _scol_btn = _mk_btn("欄位 ▾", "ghost", icon="columns"); _scol_btn.setFixedHeight(28)
        _scol_menu = QMenu(_scol_btn)
        _scol_menu.addAction(_ti_icon("plus", "#10B981"), "新增欄位", self.add_sub_column)
        _scol_menu.addAction(_ti_icon("pencil", "#A5B4FC"), "重新命名", self.rename_sub_column)
        _scol_menu.addAction(_ti_icon("trash", "#EF4444"), "刪除欄位", self.delete_sub_column)
        _scol_btn.setMenu(_scol_menu)
        sh.addWidget(_scol_btn)

        _stbl_btn = _mk_btn("子表 ▾", "ghost", icon="table"); _stbl_btn.setFixedHeight(28)
        _stbl_menu = QMenu(_stbl_btn)
        _stbl_menu.addAction(_ti_icon("plus", "#10B981"), "新增子表", self.add_sub_table)
        _stbl_menu.addAction(_ti_icon("trash", "#EF4444"), "刪除子表", self.delete_sub_table)
        _stbl_btn.setMenu(_stbl_menu)
        sh.addWidget(_stbl_btn)

        sv.addWidget(sub_hdr)
        self._sub_tabs = QTabWidget()
        self._sub_tabs.setObjectName("sub-tabs")
        self._sub_tabs.setDocumentMode(True)
        self._sub_tabs.currentChanged.connect(self._sync_sub_picker)
        sv.addWidget(self._sub_tabs, 1)
        sub_area.setMinimumHeight(110)     # ensure header + tab bar always visible
        rsplit.addWidget(sub_area)

        rsplit.setSizes([360, 200])
        rsplit.setCollapsible(0, False)
        rsplit.setCollapsible(1, False)
        rv.addWidget(rsplit, 1)

        # Draggable divider: middle list ↔ right panel (field editor + sub-tables)
        mid_right_split = QSplitter(Qt.Horizontal)
        mid_right_split.setObjectName("mid-right-split")
        mid_right_split.setStyleSheet(
            f"QSplitter#mid-right-split::handle:horizontal "
            f"{{ background: {_C['border']}; width: 4px; }}"
        )
        mid_right_split.setHandleWidth(4)
        mid_right_split.setChildrenCollapsible(False)
        mid_right_split.addWidget(mid)
        mid_right_split.addWidget(right)
        mid_right_split.setStretchFactor(0, 1)
        mid_right_split.setStretchFactor(1, 0)
        mid_right_split.setSizes([620, 440])
        root.addWidget(mid_right_split, 1)

        self._build_sub_tabs()

    # ── JSON preview ──────────────────────────────────────────────────────────

    def _toggle_json(self):
        visible = self._json_preview.isVisible()
        self._json_preview.setVisible(not visible)
        self._json_toggle.setText(
            ("▼" if not visible else "▶") + "  JSON PREVIEW"
        )

    def _update_json(self, row_data):
        data = {col: val for col, val in row_data.items()}
        raw  = json.dumps(data, ensure_ascii=False, indent=2, default=str)
        hl   = _json_highlight(raw)
        self._json_preview.setHtml(
            f'<div style="font-family:Consolas;font-size:11px;line-height:1.6;">'
            f'<pre style="margin:0;">{hl}</pre></div>'
        )

    # ── Classification ────────────────────────────────────────────────────────

    def _load_cls_list(self):
        self._cls_list.blockSignals(True)
        self._cls_list.clear()
        if self.cls_key not in self.df.columns:
            self._cls_list.blockSignals(False)
            return
        groups = self.df[self.cls_key].unique()
        for g in groups:
            cat   = _cat_for(str(g))
            count = int((self.df[self.cls_key] == g).sum())
            item  = QListWidgetItem(f"  {g}  ({count})")
            item.setData(Qt.UserRole, g)
            item.setForeground(QBrush(QColor(cat["text"])))
            self._cls_list.addItem(item)
            if str(g) == str(self.current_cls_val):
                self._cls_list.setCurrentItem(item)
        self._cls_list.blockSignals(False)

    def _on_cls_changed(self, cur, _prev):
        if cur is None:
            return
        self.current_cls_val = cur.data(Qt.UserRole)
        self._load_item_list()

    def add_classification(self):
        if self.cls_key not in self.df.columns:
            col, ok = QInputDialog.getText(self, "設定分類欄位", "分類欄位名稱:")
            if not ok or not col.strip():
                return
            col = col.strip()
            self.manager.add_column(self.table_name, col)
            self.df = self.manager.tables[self.table_name]
            self.cls_key = col
            self.cfg["classification_key"] = col
            if not self.pk_key or self.pk_key not in self.df.columns:
                self.pk_key = col
                self.cfg["primary_key"] = col
        name, ok = QInputDialog.getText(self, "新增分類", "分類名稱:")
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in self.df[self.cls_key].values:
            QMessageBox.warning(self, "錯誤", "此分類已存在"); return
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = name
        if self.pk_key in self.df.columns and self.pk_key != self.cls_key:
            new_id, ok2 = QInputDialog.getText(self, "新增分類", f"首筆資料的 {self.pk_key}:")
            if not ok2 or not new_id.strip(): return
            new_row[self.pk_key] = new_id.strip()
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self._reload_all(select_cls=name)

    def delete_classification(self):
        g = self.current_cls_val
        if g is None: return
        if QMessageBox.question(self, "確認刪除", f"刪除分類 [{g}] 及其所有資料？",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        self.df = self.df[self.df[self.cls_key] != g].reset_index(drop=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self.current_cls_val = None
        self.current_master_idx = None
        self._reload_all()

    def move_classification(self, delta):
        if self.current_cls_val is None: return
        groups = list(self.df[self.cls_key].unique())
        try:    pos = [str(g) for g in groups].index(str(self.current_cls_val))
        except: return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(groups): return
        groups[pos], groups[new_pos] = groups[new_pos], groups[pos]
        self.df = pd.concat(
            [self.df[self.df[self.cls_key] == g] for g in groups], ignore_index=True
        )
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self._reload_all(select_cls=self.current_cls_val)

    # ── Item list (card list) ─────────────────────────────────────────────────

    def _load_item_list(self):
        query    = self._filter_edit.text().strip().lower()
        self._card_list.blockSignals(True)
        self._card_list.clear()
        if self.current_cls_val is None or self.pk_key not in self.df.columns:
            self._card_list.blockSignals(False)
            return

        sub_df   = self.df[self.df[self.cls_key] == self.current_cls_val]
        sub_col  = next(
            (c for c in self.df.columns if c != self.pk_key and c != self.cls_key),
            None
        )

        # Build text-ref resolver for display
        cols_cfg = self.cfg.get("columns", {})
        _trs     = self.cfg.get("text_ref_source", {})
        _trs_json = _trs.get("json_path", "")
        _trs_key  = _trs.get("key_col", "TextID")
        _trs_val  = _trs.get("val_col", "TextContent")
        if _trs_json and self.manager.json_path:
            _json_dir = os.path.dirname(self.manager.json_path)
            _abs_ref  = os.path.join(_json_dir, _trs_json) if not os.path.isabs(_trs_json) else _trs_json
        else:
            _abs_ref  = ""

        def _disp(col, raw):
            if _abs_ref and cols_cfg.get(col, {}).get("type") == "text_ref":
                resolved = self.manager.get_ref_text(_abs_ref, _trs_key, str(raw), _trs_val)
                return resolved if resolved else str(raw)
            return str(raw)

        for df_idx, row in sub_df.iterrows():
            pk_raw   = str(row[self.pk_key])
            sub_raw  = str(row[sub_col]) if sub_col else ""
            pk_disp  = _disp(self.pk_key, pk_raw)
            sub_disp = _disp(sub_col, sub_raw) if sub_col else ""

            if query and query not in pk_raw.lower() and query not in pk_disp.lower() \
                     and query not in sub_raw.lower() and query not in sub_disp.lower():
                continue

            item = QListWidgetItem()
            item.setData(Qt.UserRole,            df_idx)
            item.setData(ItemCardDelegate.R_PK,  pk_disp)
            item.setData(ItemCardDelegate.R_SUB, sub_disp)
            item.setData(ItemCardDelegate.R_CAT, str(self.current_cls_val))
            self._card_list.addItem(item)
            if df_idx == self.current_master_idx:
                self._card_list.setCurrentItem(item)

        self._card_list.blockSignals(False)

    def _apply_filter(self, _text=""):
        self._load_item_list()

    def _on_item_changed(self, cur, _prev):
        if cur is None: return
        self._load_editor(cur.data(Qt.UserRole))

    def add_master_item(self):
        if self.current_cls_val is None:
            QMessageBox.warning(self, "提示", "請先選擇分類"); return
        new_id, ok = QInputDialog.getText(self, "新增項目", f"{self.pk_key}:")
        if not ok or not new_id.strip(): return
        new_id = new_id.strip()
        if new_id in self.df[self.pk_key].astype(str).values:
            QMessageBox.warning(self, "錯誤", "此 ID 已存在"); return
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = self.current_cls_val
        new_row[self.pk_key]  = new_id
        cls_rows     = self.df[self.df[self.cls_key] == self.current_cls_val]
        insert_after = cls_rows.index.max() + 1 if not cls_rows.empty else len(self.df)
        top, bot = self.df.iloc[:insert_after], self.df.iloc[insert_after:]
        self.df  = pd.concat([top, pd.DataFrame([new_row]), bot], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_df_idx = self.df[self.df[self.pk_key].astype(str) == str(new_id)].index[0]
        self._reload_all(select_cls=self.current_cls_val, select_idx=new_df_idx)

    def copy_master_item(self):
        if self.current_master_idx is None: return
        new_id, ok = QInputDialog.getText(self, "複製項目", f"新的 {self.pk_key}:")
        if not ok or not new_id.strip(): return
        new_id = new_id.strip()
        if new_id in self.df[self.pk_key].astype(str).values:
            QMessageBox.warning(self, "錯誤", "此 ID 已存在"); return
        new_row = self.df.loc[self.current_master_idx].copy()
        new_row[self.pk_key] = new_id
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_df_idx = self.df[self.df[self.pk_key].astype(str) == str(new_id)].index[0]
        self._reload_all(select_cls=self.current_cls_val, select_idx=new_df_idx)

    def delete_master_item(self):
        if self.current_master_idx is None: return
        pk = str(self.df.at[self.current_master_idx, self.pk_key])
        self.df.drop(self.current_master_idx, inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self.current_master_idx = None
        self._reload_all(select_cls=self.current_cls_val)
        self.status_message.emit(f"已刪除 {pk}", _C["yellow"])

    def move_master_item(self, delta):
        if self.current_master_idx is None or self.current_cls_val is None: return
        cls_idxs = list(self.df[self.df[self.cls_key] == self.current_cls_val].index)
        try:    pos = cls_idxs.index(self.current_master_idx)
        except: return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(cls_idxs): return
        cls_idxs[pos], cls_idxs[new_pos] = cls_idxs[new_pos], cls_idxs[pos]
        other_idxs = [i for i in self.df.index if i not in cls_idxs]
        all_cls    = list(self.df[self.df[self.cls_key] == self.current_cls_val].index)
        first_cls  = all_cls[0]
        before = [i for i in other_idxs if i < first_cls]
        after  = [i for i in other_idxs if i > max(all_cls)]
        self.df = self.df.loc[before + cls_idxs + after].reset_index(drop=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_idx = self.df[self.df[self.pk_key].astype(str) == str(self.current_master_pk)].index
        self._reload_all(select_cls=self.current_cls_val,
                         select_idx=new_idx[0] if not new_idx.empty else None)

    # ── Field editor ──────────────────────────────────────────────────────────

    def _load_editor(self, df_idx):
        self.current_master_idx = df_idx
        if df_idx not in self.df.index:
            return
        row_data = self.df.loc[df_idx]
        self.current_master_pk = row_data[self.pk_key]

        # Update panel header
        cat = _cat_for(str(self.current_cls_val))
        self._panel_hdr.setText(f"{self.current_master_pk}")
        self._panel_hdr.setStyleSheet(
            f"color:{cat['text']}; font-family:Consolas; font-size:13px; "
            f"font-weight:700; background:transparent;"
        )

        if self._field_panel is None:
            self._field_panel = FieldEditorWidget(self._field_container)
            self._field_panel.field_changed.connect(self._on_field_change)
            self._field_container.layout().addWidget(self._field_panel)
            self._field_panel.build_for(self.df, self.cfg, self.table_name, self.manager)

        self._field_panel.load_row(row_data, df_idx)
        self._update_json(row_data)
        self._refresh_sub_tables()

    def _on_field_change(self, col, value):
        if self.current_master_idx is None: return
        self.manager.update_cell(self.table_name, self.current_master_idx, col, value)

    # ── Sub-tables ────────────────────────────────────────────────────────────

    def _build_sub_tabs(self):
        self._sub_tabs.clear()
        self._sub_panels.clear()
        self._sub_tab_order = []
        prefix = self.table_name + "."
        for key in self.manager.sub_tables:
            if not key.startswith(prefix):
                continue
            tab_name = key[len(prefix):]
            self._sub_tab_order.append(tab_name)
            # Create the real panel immediately — no placeholder swap needed
            sub_cfg  = self.cfg.get("sub_tables", {}).get(tab_name, {})
            cols_cfg = sub_cfg.get("columns", {})
            panel = SubTablePanel(key, cols_cfg, self.manager)
            panel.row_deleted.connect(self._on_sub_delete)
            panel.copy_requested.connect(self.copy_sub_rows)
            panel.paste_requested.connect(self.paste_sub_rows)
            panel.compare_requested.connect(self.compare_sub_effect)
            self._sub_panels[tab_name] = panel
            idx = self._sub_tabs.addTab(panel, tab_name)
            note = sub_cfg.get("note", "")
            if note:
                self._sub_tabs.setTabToolTip(idx, note)

        if not self._sub_tab_order:
            no_sub = QLabel("此表格無巢狀子表")
            no_sub.setAlignment(Qt.AlignCenter)
            no_sub.setStyleSheet(
                f"color:{_C['txt3']}; font-size:12px; background:{_C['panel']};"
            )
            self._sub_tabs.addTab(no_sub, "—")

        # Sub-table quick-jump picker (handy when there are many sub-tables)
        if hasattr(self, "_sub_picker"):
            self._sub_picker.blockSignals(True)
            self._sub_picker.clear()
            self._sub_picker.addItems(self._sub_tab_order)
            self._sub_picker.blockSignals(False)
            self._sub_picker.setVisible(len(self._sub_tab_order) > 1)

        self._fit_right_panel()

        self.status_message.emit(
            f"從表: 偵測到 {len(self._sub_tab_order)} 個"
            + (f"  ({', '.join(self._sub_tab_order)})" if self._sub_tab_order else ""),
            _C["txt2"],
        )

    def _refresh_sub_tables(self):
        """Reload each sub-table panel with rows matching the currently selected master pk."""
        if self.current_master_pk is None:
            return
        prefix = self.table_name + "."
        for tab_name in self._sub_tab_order:
            panel = self._sub_panels.get(tab_name)
            if panel is None:
                continue
            full   = prefix + tab_name
            sub_df = self.manager.sub_tables.get(full)
            if sub_df is None:
                continue
            sub_cfg  = self.cfg.get("sub_tables", {}).get(tab_name, {})
            fk_key   = sub_cfg.get("foreign_key", self.pk_key)
            cols_cfg = sub_cfg.get("columns", {})
            try:
                filtered = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]
            except KeyError:
                # FK column not found — fall back to first column
                fk_key   = sub_df.columns[0] if len(sub_df.columns) > 0 else None
                if fk_key is None:
                    continue
                filtered = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]
            panel.reload(filtered, cols_cfg)

    def _on_sub_delete(self, sheet_full, df_idx):
        sub_df = self.manager.sub_tables.get(sheet_full)
        if sub_df is None or df_idx not in sub_df.index: return
        sub_df.drop(df_idx, inplace=True)
        sub_df.reset_index(drop=True, inplace=True)
        self.manager.sub_tables[sheet_full] = sub_df
        self.manager.dirty = True
        self._refresh_sub_tables()
        self.status_message.emit("已刪除子表列", _C["yellow"])

    def _current_sub_panel(self) -> SubTablePanel | None:
        idx = self._sub_tabs.currentIndex()
        if idx < 0: return None
        return self._sub_panels.get(self._sub_tabs.tabText(idx))

    def add_sub_item(self):
        if self.current_master_pk is None: return
        panel = self._current_sub_panel()
        if panel is None: return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full     = self.table_name + "." + tab_name
        sub_df   = self.manager.sub_tables.get(full)
        if sub_df is None: return
        sub_cfg  = self.cfg.get("sub_tables", {}).get(tab_name, {})
        fk_key   = sub_cfg.get("foreign_key", self.pk_key)
        new_row  = {col: "" for col in sub_df.columns}
        new_row[fk_key] = self.current_master_pk
        siblings = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]
        insert_at = siblings.index.max() + 1 if not siblings.empty else len(sub_df)
        top, bot = sub_df.iloc[:insert_at], sub_df.iloc[insert_at:]
        self.manager.sub_tables[full] = pd.concat([top, pd.DataFrame([new_row]), bot], ignore_index=True)
        self.manager.dirty = True
        self._refresh_sub_tables()

    def delete_sub_item(self):
        panel = self._current_sub_panel()
        if panel is None: return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full     = self.table_name + "." + tab_name
        df_idx   = panel.selected_df_index()
        if df_idx is None: return
        self._on_sub_delete(full, df_idx)

    def move_sub_item(self, delta):
        panel = self._current_sub_panel()
        if panel is None: return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full     = self.table_name + "." + tab_name
        sub_df   = self.manager.sub_tables.get(full)
        if sub_df is None: return
        df_idx   = panel.selected_df_index()
        if df_idx is None: return
        sub_cfg  = self.cfg.get("sub_tables", {}).get(tab_name, {})
        fk_key   = sub_cfg.get("foreign_key", self.pk_key)
        siblings = list(sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)].index)
        try:    pos = siblings.index(df_idx)
        except: return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(siblings): return
        siblings[pos], siblings[new_pos] = siblings[new_pos], siblings[pos]
        others = [i for i in sub_df.index if i not in siblings]
        first  = siblings[0] if siblings else 0
        before = [i for i in others if i < first]
        after  = [i for i in others if i > max(siblings)]
        self.manager.sub_tables[full] = sub_df.loc[before + siblings + after].reset_index(drop=True)
        self.manager.dirty = True
        self._refresh_sub_tables()

    def copy_sub_item(self):
        panel = self._current_sub_panel()
        if panel is None: return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full     = self.table_name + "." + tab_name
        sub_df   = self.manager.sub_tables.get(full)
        if sub_df is None: return
        df_idx = panel.selected_df_index()
        if df_idx is None: return
        new_row = sub_df.loc[df_idx].copy()
        self.manager.sub_tables[full] = pd.concat([sub_df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.dirty = True
        self._refresh_sub_tables()

    # ── Cross-item row copy / paste ─────────────────────────────────────────────
    def copy_sub_rows(self):
        """Copy the selected sub-table row(s) to a shared buffer (FK excluded),
        so they can be pasted into another master item's (or sheet's) sub-table."""
        panel = self._current_sub_panel()
        if panel is None: return
        tab_name, full = self._current_sub_full()
        if full is None: return
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None: return
        idxs = panel.selected_df_indices()
        if not idxs:
            self.status_message.emit("請先選取要複製的子表列", _C["yellow"]); return
        fk_key = self.cfg.get("sub_tables", {}).get(tab_name, {}).get("foreign_key", self.pk_key)
        rows = []
        for di in idxs:
            if di not in sub_df.index: continue
            row = {c: sub_df.at[di, c] for c in sub_df.columns if c != fk_key}
            rows.append(row)
        TableEditor._sub_clipboard = {"sub": tab_name, "rows": rows}
        self.status_message.emit(f"已複製 {len(rows)} 列（子表: {tab_name}）", _C["green"])

    def paste_sub_rows(self):
        """Paste buffered rows into the current sub-table under the current master
        item's FK. Only columns that exist in the target sub-table are filled."""
        if self.current_master_pk is None:
            self.status_message.emit("請先選擇一個母表項目", _C["yellow"]); return
        buf = TableEditor._sub_clipboard
        if not buf or not buf.get("rows"):
            self.status_message.emit("列剪貼簿是空的", _C["yellow"]); return
        tab_name, full = self._current_sub_full()
        if full is None: return
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None: return
        fk_key = self.cfg.get("sub_tables", {}).get(tab_name, {}).get("foreign_key", self.pk_key)
        cols = list(sub_df.columns)
        new_rows = []
        for src in buf["rows"]:
            nr = {c: "" for c in cols}
            nr[fk_key] = self.current_master_pk
            for c in cols:
                if c != fk_key and c in src:
                    nr[c] = src[c]
            new_rows.append(nr)
        if not new_rows: return
        siblings = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]
        insert_at = siblings.index.max() + 1 if not siblings.empty else len(sub_df)
        top, bot = sub_df.iloc[:insert_at], sub_df.iloc[insert_at:]
        self.manager.sub_tables[full] = pd.concat(
            [top, pd.DataFrame(new_rows), bot], ignore_index=True)
        self.manager.dirty = True
        self._refresh_sub_tables()
        note = "" if buf["sub"] == tab_name else f"（來源子表: {buf['sub']}，只貼同名欄位）"
        self.status_message.emit(f"已貼上 {len(new_rows)} 列{note}", _C["green"])

    def _make_name_resolver(self):
        """Return fk_value -> readable name (via master 'Name' column + text_ref)."""
        df = self.df; pk = self.pk_key
        namecol = "Name" if "Name" in df.columns else None
        raw_map = {}
        if namecol and pk in df.columns:
            for _, row in df.iterrows():
                raw_map[str(row[pk])] = str(row[namecol])
        trs = self.cfg.get("text_ref_source", {})
        ref = (trs.get("json_path") or "").strip()
        kc = trs.get("key_col", "TextID"); vc = trs.get("val_col", "TextContent")
        abs_ref = None
        if ref and self.manager.json_path:
            jd = os.path.dirname(self.manager.json_path)
            abs_ref = ref if os.path.isabs(ref) else os.path.join(jd, ref)
        def resolve(fk):
            raw = raw_map.get(str(fk), "")
            if abs_ref and raw:
                r = self.manager.get_ref_text(abs_ref, kc, raw, vc)
                if r: return r
            return raw or str(fk)
        return resolve

    def compare_sub_effect(self, df_idx):
        tab_name, full = self._current_sub_full()
        if full is None: return
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None or df_idx not in sub_df.index: return
        seed = {c: sub_df.at[df_idx, c] for c in sub_df.columns}
        fk_key = self.cfg.get("sub_tables", {}).get(tab_name, {}).get("foreign_key", self.pk_key)
        dlg = EffectCompareDialog(sub_df, fk_key, seed, self._make_name_resolver(), tab_name, self)
        dlg.exec()

    # ── Context menus ─────────────────────────────────────────────────────────

    def _cls_ctx_menu(self, pos):
        item = self._cls_list.itemAt(pos)
        if item is None: return
        g = item.data(Qt.UserRole)
        menu = QMenu(self)
        menu.addAction("重新命名", lambda: self._rename_cls(g))
        menu.addAction("刪除此分類", self.delete_classification)
        menu.exec(self._cls_list.mapToGlobal(pos))

    def _rename_cls(self, old_val):
        new_val, ok = QInputDialog.getText(self, "重新命名", "新名稱:", text=str(old_val))
        if not ok or not new_val.strip() or new_val.strip() == str(old_val): return
        self.df[self.cls_key] = self.df[self.cls_key].replace(old_val, new_val.strip())
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        if str(self.current_cls_val) == str(old_val):
            self.current_cls_val = new_val.strip()
        self._reload_all(select_cls=self.current_cls_val)

    def _item_ctx_menu(self, pos):
        item = self._card_list.itemAt(pos)
        if item is None: return
        menu = QMenu(self)
        menu.addAction("複製", self.copy_master_item)
        menu.addAction("刪除", self.delete_master_item)
        menu.exec(self._card_list.mapToGlobal(pos))

    # ── Refresh ───────────────────────────────────────────────────────────────

    def _reload_all(self, select_cls=None, select_idx=None):
        self.df = self.manager.tables[self.table_name]
        if select_cls is not None:  self.current_cls_val    = select_cls
        if select_idx is not None:  self.current_master_idx = select_idx
        self._load_cls_list()
        if self.current_cls_val is not None:
            self._load_item_list()
        if self.current_master_idx is not None and self.current_master_idx in self.df.index:
            self._load_editor(self.current_master_idx)

    def add_master_column(self):
        col_name, ok = QInputDialog.getText(self, "新增欄位", "欄位名稱:")
        if not ok or not col_name.strip(): return
        col_name = col_name.strip()
        if col_name in self.df.columns:
            QMessageBox.warning(self, "錯誤", f"欄位 [{col_name}] 已存在"); return
        self.df[col_name] = ""
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        if self._field_panel:
            self._field_panel.deleteLater()
            self._field_panel = None
        self._reload_all(select_cls=self.current_cls_val, select_idx=self.current_master_idx)
        self.status_message.emit(f"欄位 [{col_name}] 已新增", _C["green"])

    def add_sub_column(self):
        tab_idx = self._sub_tabs.currentIndex()
        if tab_idx < 0: return
        tab_name = self._sub_tabs.tabText(tab_idx)
        if not tab_name or tab_name == "—": return
        full   = self.table_name + "." + tab_name
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None: return
        col_name, ok = QInputDialog.getText(self, "新增欄位", f"欄位名稱（子表: {tab_name}）:")
        if not ok or not col_name.strip(): return
        col_name = col_name.strip()
        if col_name in sub_df.columns:
            QMessageBox.warning(self, "錯誤", f"欄位 [{col_name}] 已存在"); return
        sub_df[col_name] = ""
        self.manager.sub_tables[full] = sub_df
        self.manager.dirty = True
        self._build_sub_tabs()
        self._select_sub_tab(tab_name)
        self._refresh_sub_tables()
        self.status_message.emit(f"從表欄位 [{col_name}] 已新增", _C["green"])

    def add_sub_table(self):
        if self.table_name not in self.manager.tables:
            return
        if not self.pk_key:
            QMessageBox.warning(self, "提示", "母表沒有主鍵，無法建立子表"); return
        name, ok = QInputDialog.getText(self, "新增子表", "子表名稱（巢狀陣列欄位名）:")
        if not ok or not name.strip(): return
        name = name.strip()
        full = self.table_name + "." + name
        if full in self.manager.sub_tables:
            QMessageBox.warning(self, "錯誤", f"子表 [{name}] 已存在"); return
        if name in self.df.columns:
            QMessageBox.warning(self, "錯誤", f"名稱 [{name}] 與母表欄位重複"); return
        if not self.manager.add_sub_table(self.table_name, name):
            QMessageBox.warning(self, "錯誤", "建立子表失敗"); return
        self._build_sub_tabs()
        self._select_sub_tab(name)
        self._refresh_sub_tables()
        self.status_message.emit(f"子表 [{name}]（FK={self.pk_key}）已新增", _C["green"])

    def delete_sub_table(self):
        tab_name, full = self._current_sub_full()
        if full is None:
            QMessageBox.information(self, "提示", "目前沒有可刪除的子表"); return
        if QMessageBox.question(
                self, "確認刪除",
                f"確定刪除整張子表 [{tab_name}]？\n會移除其所有資料列與欄位定義。") != QMessageBox.Yes:
            return
        self.manager.delete_sub_table(self.table_name, tab_name)
        self._build_sub_tabs()
        self._refresh_sub_tables()
        self.status_message.emit(f"子表 [{tab_name}] 已刪除", _C["yellow"])

    # ── Column rename / delete (master + sub) ───────────────────────────────────

    def _select_sub_tab(self, name):
        for i in range(self._sub_tabs.count()):
            if self._sub_tabs.tabText(i) == name:
                self._sub_tabs.setCurrentIndex(i); return

    def _fit_right_panel(self):
        """Make the right panel's min width track what the sub-table toolbar
        actually needs, so its buttons never get squeezed (hardcoded → adaptive)."""
        if not hasattr(self, "_sub_hdr") or not hasattr(self, "_right_panel"):
            return
        self._sub_hdr.layout().activate()
        need = self._sub_hdr.sizeHint().width() + 24
        self._right_panel.setMinimumWidth(max(320, need))

    def _on_sub_picker(self, idx):
        name = self._sub_picker.itemText(idx)
        if name:
            self._select_sub_tab(name)

    def _sync_sub_picker(self, idx):
        if not hasattr(self, "_sub_picker"):
            return
        name = self._sub_tabs.tabText(idx) if idx >= 0 else ""
        pi = self._sub_picker.findText(name)
        if pi >= 0 and pi != self._sub_picker.currentIndex():
            self._sub_picker.blockSignals(True)
            self._sub_picker.setCurrentIndex(pi)
            self._sub_picker.blockSignals(False)

    def _current_sub_full(self):
        idx = self._sub_tabs.currentIndex()
        if idx < 0:
            return None, None
        tab_name = self._sub_tabs.tabText(idx)
        if not tab_name or tab_name == "—":
            return None, None
        return tab_name, self.table_name + "." + tab_name

    def _pick_column(self, columns, exclude, title):
        choices = [c for c in columns if c not in exclude]
        if not choices:
            QMessageBox.information(self, "提示", "沒有可操作的欄位"); return None
        col, ok = QInputDialog.getItem(self, title, "選擇欄位:", choices, 0, False)
        if not ok or not col:
            return None
        return col

    def rename_master_column(self):
        col = self._pick_column(list(self.df.columns), {self.pk_key, self.cls_key}, "重新命名欄位")
        if col is None: return
        new, ok = QInputDialog.getText(self, "重新命名", f"新欄位名稱（{col} →）:")
        if not ok or not new.strip(): return
        new = new.strip()
        if new in self.df.columns:
            QMessageBox.warning(self, "錯誤", f"欄位 [{new}] 已存在"); return
        self.manager.rename_column(self.table_name, col, new)
        self.df = self.manager.tables[self.table_name]
        if self._field_panel:
            self._field_panel.deleteLater(); self._field_panel = None
        self._reload_all(select_cls=self.current_cls_val, select_idx=self.current_master_idx)
        self.status_message.emit(f"欄位 [{col}] → [{new}]", _C["green"])

    def delete_master_column(self):
        col = self._pick_column(list(self.df.columns), {self.pk_key, self.cls_key}, "刪除欄位")
        if col is None: return
        if QMessageBox.question(self, "確認刪除", f"確定刪除欄位 [{col}]？") != QMessageBox.Yes:
            return
        self.manager.delete_column(self.table_name, col)
        self.df = self.manager.tables[self.table_name]
        if self._field_panel:
            self._field_panel.deleteLater(); self._field_panel = None
        self._reload_all(select_cls=self.current_cls_val, select_idx=self.current_master_idx)
        self.status_message.emit(f"欄位 [{col}] 已刪除", _C["yellow"])

    def rename_sub_column(self):
        tab_name, full = self._current_sub_full()
        if full is None: return
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None: return
        fk = self.cfg.get("sub_tables", {}).get(tab_name, {}).get("foreign_key", self.pk_key)
        col = self._pick_column(list(sub_df.columns), {fk}, f"重新命名欄位（子表: {tab_name}）")
        if col is None: return
        new, ok = QInputDialog.getText(self, "重新命名", f"新欄位名稱（{col} →）:")
        if not ok or not new.strip(): return
        new = new.strip()
        if new in sub_df.columns:
            QMessageBox.warning(self, "錯誤", f"欄位 [{new}] 已存在"); return
        self.manager.rename_column(full, col, new)
        self._build_sub_tabs(); self._select_sub_tab(tab_name); self._refresh_sub_tables()
        self.status_message.emit(f"子表欄位 [{col}] → [{new}]", _C["green"])

    def delete_sub_column(self):
        tab_name, full = self._current_sub_full()
        if full is None: return
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None: return
        fk = self.cfg.get("sub_tables", {}).get(tab_name, {}).get("foreign_key", self.pk_key)
        col = self._pick_column(list(sub_df.columns), {fk}, f"刪除欄位（子表: {tab_name}）")
        if col is None: return
        if QMessageBox.question(self, "確認刪除", f"確定刪除子表欄位 [{col}]？") != QMessageBox.Yes:
            return
        self.manager.delete_column(full, col)
        self._build_sub_tabs(); self._select_sub_tab(tab_name); self._refresh_sub_tables()
        self.status_message.emit(f"子表欄位 [{col}] 已刪除", _C["yellow"])

    def reload_after_config(self):
        self.cfg     = self.manager.config.get(self.table_name, {})
        cols         = list(self.df.columns)
        self.cls_key = self.cfg.get("classification_key", "") or (cols[0] if cols else "")
        self.pk_key  = self.cfg.get("primary_key",        "") or (cols[0] if cols else "")
        if self._field_panel:
            self._field_panel.deleteLater()
            self._field_panel = None
        self._build_sub_tabs()
        self._reload_all()


# ── WelcomeWidget ─────────────────────────────────────────────────────────────

class WelcomeWidget(QWidget):
    open_file   = Signal()
    new_file    = Signal()
    open_recent = Signal(str)

    def __init__(self, manager: JsonDataManager, parent=None):
        super().__init__(parent)
        self._manager = manager
        self.setStyleSheet(f"background: {_C['bg']};")
        self._setup_ui()

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setAlignment(Qt.AlignCenter)

        card = QWidget()
        card.setFixedWidth(440)
        card.setStyleSheet(
            f"background:{_C['card']}; border-radius:12px; "
            f"border:1px solid {_C['border']};"
        )
        lo = QVBoxLayout(card)
        lo.setContentsMargins(40, 36, 40, 36)
        lo.setSpacing(0)

        # Logo
        logo = QLabel("{ }")
        logo.setAlignment(Qt.AlignCenter)
        logo.setFixedSize(60, 60)
        logo.setStyleSheet(
            f"font-family:Consolas; font-size:22px; font-weight:bold; color:white; "
            f"background:qlineargradient(x1:0,y1:0,x2:1,y2:1,"
            f"stop:0 {_C['accent']},stop:1 #8B5CF6); border-radius:12px; border:none;"
        )
        logo_wrap = QHBoxLayout()
        logo_wrap.addStretch()
        logo_wrap.addWidget(logo)
        logo_wrap.addStretch()
        lo.addLayout(logo_wrap)
        lo.addSpacing(16)

        title = QLabel("JsonEditor")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(
            f"font-size:20px; font-weight:700; color:{_C['txt']}; background:transparent; border:none;"
        )
        lo.addWidget(title)

        subtitle = QLabel("輕量 JSON 資料編輯器")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(
            f"color:{_C['txt2']}; font-size:12px; background:transparent; border:none;"
        )
        lo.addWidget(subtitle)
        lo.addSpacing(28)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        b_open = _mk_btn("📂  開啟 JSON", "primary")
        b_open.setFixedHeight(38)
        b_open.clicked.connect(self.open_file)
        b_new = _mk_btn("📄  新建 JSON")
        b_new.setFixedHeight(38)
        b_new.clicked.connect(self.new_file)
        btn_row.addWidget(b_open, 1)
        btn_row.addWidget(b_new, 1)
        lo.addLayout(btn_row)

        recent = self._manager._recent_files
        if recent:
            lo.addSpacing(20)
            sep = QFrame()
            sep.setFixedHeight(1)
            sep.setStyleSheet(f"background:{_C['border']}; border:none;")
            lo.addWidget(sep)
            lo.addSpacing(12)

            hdr = QLabel("最近開啟")
            hdr.setStyleSheet(
                f"color:{_C['txt3']}; font-size:10px; font-weight:600; "
                f"letter-spacing:1px; background:transparent; border:none;"
            )
            lo.addWidget(hdr)
            lo.addSpacing(6)

            for path in recent[:6]:
                fname = os.path.basename(path)
                dirn  = os.path.dirname(path)
                row   = QWidget()
                row.setStyleSheet(
                    f"QWidget {{ background:transparent; border-radius:6px; border:none; }}"
                    f"QWidget:hover {{ background:{_C['cardH']}; }}"
                )
                row.setCursor(Qt.PointingHandCursor)
                rlo = QHBoxLayout(row)
                rlo.setContentsMargins(8, 6, 8, 6)
                lbl_fn  = QLabel(fname)
                lbl_fn.setStyleSheet(
                    f"color:{_C['txtAcc']}; font-weight:600; background:transparent; border:none;"
                )
                lbl_dir = QLabel(dirn)
                lbl_dir.setStyleSheet(
                    f"color:{_C['txt3']}; font-size:10px; background:transparent; border:none;"
                )
                lbl_dir.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                rlo.addWidget(lbl_fn)
                rlo.addWidget(lbl_dir, 1)
                row.mousePressEvent = lambda e, p=path: self.open_recent.emit(p)
                lo.addWidget(row)

        outer.addWidget(card)

    def refresh(self, manager):
        self._manager = manager
        while self.layout().count():
            item = self.layout().takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._setup_ui()


# ── App ───────────────────────────────────────────────────────────────────────

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JsonEditor")
        self.resize(1280, 800)
        self.setMinimumSize(900, 600)
        self.manager         = JsonDataManager()
        self._editors:       dict[str, TableEditor | None] = {}
        self._active_worker  = None
        self._snackbar_timer = QTimer(self)
        self._snackbar_timer.setSingleShot(True)
        self._snackbar_timer.timeout.connect(self._clear_status)

        self._setup_content()

        # Keyboard shortcuts
        self.addAction(QAction(parent=self, shortcut=QKeySequence("Ctrl+S"), triggered=self.save_file))
        self.addAction(QAction(parent=self, shortcut=QKeySequence("Ctrl+O"), triggered=self.load_file))
        self.addAction(QAction(parent=self, shortcut=QKeySequence("Ctrl+F"), triggered=self._show_search))

        self._show_welcome()
        self._update_sync()

    # ── Content layout ────────────────────────────────────────────────────────

    def _setup_content(self):
        central = QWidget()
        central.setStyleSheet(f"background:{_C['bg']};")
        self.setCentralWidget(central)
        clo = QVBoxLayout(central)
        clo.setContentsMargins(0, 0, 0, 0)
        clo.setSpacing(0)

        # ── Top bar ──
        topbar = QWidget()
        topbar.setObjectName("topbar")
        topbar.setFixedHeight(54)
        topbar.setStyleSheet(
            f"QWidget#topbar {{ background:{_C['sidebar']}; border-bottom:1px solid {_C['border']}; }}"
        )
        tlo = QHBoxLayout(topbar)
        tlo.setContentsMargins(16, 0, 16, 0)
        tlo.setSpacing(10)

        # Logo
        logo_box = QLabel("{ }")
        logo_box.setFixedSize(34, 34)
        logo_box.setAlignment(Qt.AlignCenter)
        logo_box.setStyleSheet(
            f"font-family:Consolas; font-size:14px; font-weight:bold; color:white; "
            f"background:qlineargradient(x1:0,y1:0,x2:1,y2:1,"
            f"stop:0 {_C['accent']},stop:1 #8B5CF6); border-radius:9px;"
        )
        tlo.addWidget(logo_box)

        app_name = QLabel("JsonEditor")
        app_name.setStyleSheet(
            f"font-size:15px; font-weight:700; color:{_C['txt']}; background:transparent;"
        )
        tlo.addWidget(app_name)

        app_sub = QLabel("Pro")
        app_sub.setStyleSheet(
            f"font-size:11px; color:{_C['txt2']}; background:transparent;"
        )
        tlo.addWidget(app_sub)
        tlo.addStretch(1)

        # Action buttons
        for text, role, slot in [
            ("📂  開啟",   "",        self.load_file),
            ("+ 新建",    "",        self._new_file),
            ("💾  儲存",  "success", self.save_file),
            ("📤  匯出 Excel", "",  self._export_to_excel),
        ]:
            b = _mk_btn(text, role)
            b.setFixedHeight(34)
            b.clicked.connect(slot)
            tlo.addWidget(b)

        tlo.addSpacing(10)
        tlo.addWidget(_vsep())
        tlo.addSpacing(10)

        # Sync indicator
        self._sync_dot = QLabel("●")
        self._sync_dot.setFixedWidth(14)
        self._sync_dot.setStyleSheet(f"color:{_C['green']}; background:transparent; font-size:9px;")
        self._sync_lbl = QLabel("已儲存")
        self._sync_lbl.setStyleSheet(
            f"font-size:12px; font-weight:500; color:{_C['green']}; background:transparent;"
        )
        tlo.addWidget(self._sync_dot)
        tlo.addWidget(self._sync_lbl)

        tlo.addSpacing(14)

        # Util buttons
        for text, slot in [
            ("🔍", self._show_search),
            ("🩺", self.open_health_check),
            ("⚙",  self.open_config),
            ("🕓", self._show_recent_menu),
        ]:
            b = _mk_btn(text, "ghost")
            b.setFixedSize(36, 34)
            b.clicked.connect(slot)
            tlo.addWidget(b)

        clo.addWidget(topbar)

        # ── Stack ──
        self._stack = QStackedWidget()
        clo.addWidget(self._stack, 1)

        self._welcome = WelcomeWidget(self.manager)
        self._welcome.open_file.connect(self.load_file)
        self._welcome.new_file.connect(self._new_file)
        self._welcome.open_recent.connect(self._load_recent)
        self._stack.addWidget(self._welcome)

        self._tab_widget = QTabWidget()
        self._tab_widget.setObjectName("main-tabs")
        self._tab_widget.setDocumentMode(True)
        self._tab_widget.currentChanged.connect(self._on_tab_changed)
        self._stack.addWidget(self._tab_widget)

        # ── Status strip ──
        status_bar = QWidget()
        status_bar.setFixedHeight(26)
        status_bar.setStyleSheet(
            f"background:{_C['sidebar']}; border-top:1px solid {_C['border']};"
        )
        slo = QHBoxLayout(status_bar)
        slo.setContentsMargins(14, 0, 14, 0)
        slo.setSpacing(0)
        self._status_lbl = QLabel("就緒")
        self._status_lbl.setStyleSheet(
            f"font-size:11px; color:{_C['txt2']}; background:transparent;"
        )
        self._path_lbl = QLabel("")
        self._path_lbl.setStyleSheet(
            f"font-size:11px; color:{_C['txt3']}; background:transparent;"
        )
        self._path_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        slo.addWidget(self._status_lbl, 1)
        slo.addWidget(self._path_lbl)
        clo.addWidget(status_bar)

    def _show_welcome(self):
        self._stack.setCurrentWidget(self._welcome)

    def _show_editor(self):
        self._stack.setCurrentWidget(self._tab_widget)

    # ── Sync / status ─────────────────────────────────────────────────────────

    def _update_sync(self):
        dirty = self.manager.dirty
        if dirty:
            self._sync_dot.setStyleSheet(
                f"color:{_C['yellow']}; background:transparent; font-size:9px;"
            )
            self._sync_lbl.setText("未儲存")
            self._sync_lbl.setStyleSheet(
                f"font-size:12px; font-weight:500; color:{_C['yellow']}; background:transparent;"
            )
        else:
            self._sync_dot.setStyleSheet(
                f"color:{_C['green']}; background:transparent; font-size:9px;"
            )
            self._sync_lbl.setText("已儲存")
            self._sync_lbl.setStyleSheet(
                f"font-size:12px; font-weight:500; color:{_C['green']}; background:transparent;"
            )
        if self.manager.json_path:
            fname = os.path.basename(self.manager.json_path)
            self.setWindowTitle(f"JsonEditor — {fname}" + (" *" if dirty else ""))
            self._path_lbl.setText(self.manager.json_path)
        else:
            self.setWindowTitle("JsonEditor")
            self._path_lbl.setText("")

    def show_snackbar(self, text: str, duration_ms: int = 3000, color: str = ""):
        self._status_lbl.setText(text)
        self._status_lbl.setStyleSheet(
            f"font-size:11px; color:{color or _C['txtAcc']}; background:transparent;"
        )
        self._snackbar_timer.start(duration_ms)

    def _clear_status(self):
        self._status_lbl.setText("就緒")
        self._status_lbl.setStyleSheet(
            f"font-size:11px; color:{_C['txt2']}; background:transparent;"
        )

    # Alias for backwards-compat with TableEditor signal
    def _update_title(self):
        self._update_sync()

    # ── Loading ───────────────────────────────────────────────────────────────

    def _set_loading(self, loading: bool, msg: str = "就緒"):
        self.setEnabled(not loading)
        self._status_lbl.setText(msg)
        self._status_lbl.setStyleSheet(
            f"font-size:11px; color:{_C['cyan'] if loading else _C['txt2']}; background:transparent;"
        )
        QApplication.processEvents()

    # ── File I/O ──────────────────────────────────────────────────────────────

    def load_file(self):
        last_dir = os.path.dirname(self.manager._full_config.get("_last_file", "")) or ""
        path, _ = QFileDialog.getOpenFileName(
            self, "開啟 JSON", last_dir, "JSON 檔案 (*.json);;所有檔案 (*.*)"
        )
        if path:
            self._load_path(path)

    def _new_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "新建 JSON", "", "JSON 檔案 (*.json)")
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump([], f, indent=4, ensure_ascii=False)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e)); return
        self._load_path(path)

    def _load_recent(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, "錯誤", f"找不到檔案：\n{path}")
            self.manager._recent_files = [p for p in self.manager._recent_files if p != path]
            self.manager.save_config()
            return
        self._load_path(path)

    def _load_path(self, path):
        self._set_loading(True, f"載入 {os.path.basename(path)}…")
        orig = sys.getswitchinterval()
        sys.setswitchinterval(0.001)
        worker = _LoadWorker(self.manager, path)
        self._active_worker = worker

        def _done():
            sys.setswitchinterval(orig)
            self._active_worker = None
            self._set_loading(False)
            self._refresh_ui()
            self.manager._full_config["_last_file"] = path
            self.manager.save_config()
            self._check_config_paths()

        def _err(msg):
            sys.setswitchinterval(orig)
            self._active_worker = None
            self._set_loading(False)
            QMessageBox.critical(self, "載入失敗", msg)

        worker.done.connect(_done)
        worker.error.connect(_err)
        worker.start()

    # ── Config path validation (offer to re-pick stale external paths) ──────────
    def _store_path(self, p, base):
        """Store relative to the json dir when possible (mirrors config dialog)."""
        if not base:
            return p
        try:
            return os.path.relpath(p, base)
        except ValueError:
            return p

    def _check_config_paths(self):
        """After load: if config has external paths that no longer exist, offer
        to re-pick them (外部文字表 json_path / 圖片資料夾 base_folder)."""
        if not self.manager.json_path:
            return
        json_dir = os.path.dirname(self.manager.json_path)
        changed = False
        for tname, cfg in list(self.manager.config.items()):
            if not isinstance(cfg, dict):
                continue
            trs = cfg.get("text_ref_source", {})
            ref = (trs.get("json_path") or "").strip() if isinstance(trs, dict) else ""
            if ref:
                abs_ref = ref if os.path.isabs(ref) else os.path.join(json_dir, ref)
                if not os.path.isfile(abs_ref):
                    if self._prompt_repath(tname, "外部文字表", abs_ref):
                        new, _ = QFileDialog.getOpenFileName(
                            self, f"重新選擇外部文字表（{tname}）", json_dir,
                            "JSON 檔案 (*.json);;所有檔案 (*.*)")
                        if new:
                            trs["json_path"] = self._store_path(new, json_dir)
                            cfg["text_ref_source"] = trs
                            changed = True
            imgp = cfg.get("image_preview", {})
            base = (imgp.get("base_folder") or "").strip() if isinstance(imgp, dict) else ""
            if base:
                abs_base = base if os.path.isabs(base) else os.path.join(json_dir, base)
                if not os.path.isdir(abs_base):
                    if self._prompt_repath(tname, "圖片資料夾", abs_base):
                        new = QFileDialog.getExistingDirectory(
                            self, f"重新選擇圖片資料夾（{tname}）", json_dir)
                        if new:
                            imgp["base_folder"] = self._store_path(new, json_dir)
                            cfg["image_preview"] = imgp
                            changed = True
        if changed:
            self.manager.save_config()
            self.manager.invalidate_ref_cache()
            idx = self._tab_widget.currentIndex()
            if idx >= 0:
                ed = self._editors.get(self._tab_widget.tabText(idx))
                if ed:
                    ed.reload_after_config()
            self.show_snackbar("✓ 路徑已更新", color=_C["green"])

    def _prompt_repath(self, tname, label, badpath):
        return QMessageBox.question(
            self, "路徑不存在",
            f"[{tname}] 的{label}路徑找不到：\n{badpath}\n\n是否要重新設定？",
            QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes

    def save_file(self):
        if not self.manager.json_path:
            QMessageBox.warning(self, "提示", "尚未載入任何 JSON 檔案"); return
        self._set_loading(True, "儲存中…")
        orig = sys.getswitchinterval()
        sys.setswitchinterval(0.001)
        worker = _SaveWorker(self.manager)
        self._active_worker = worker

        def _done():
            sys.setswitchinterval(orig)
            self._active_worker = None
            self._set_loading(False)
            self.show_snackbar("✓ 已儲存", color=_C["green"])
            self._update_sync()

        def _err(msg):
            sys.setswitchinterval(orig)
            self._active_worker = None
            self._set_loading(False)
            QMessageBox.critical(self, "存檔失敗", msg)

        worker.done.connect(_done)
        worker.error.connect(_err)
        worker.start()

    # ── Export to Excel ───────────────────────────────────────────────────────

    def _export_to_excel(self):
        if not self.manager.json_path:
            QMessageBox.warning(self, "提示", "尚未載入任何 JSON 檔案"); return
        if not self.manager.tables:
            QMessageBox.warning(self, "提示", "沒有資料可以匯出"); return
        base = os.path.splitext(os.path.basename(self.manager.json_path))[0]
        default = os.path.join(os.path.dirname(self.manager.json_path), base + ".xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self, "匯出為 Excel（可選現有檔案直接覆蓋）",
            default, "Excel 檔案 (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self._write_xlsx(path)
        except Exception as e:
            QMessageBox.critical(self, "匯出失敗", str(e))
            return
        self.show_snackbar(f"✓ 已匯出 Excel: {os.path.basename(path)}", color=_C["green"])

    def _write_xlsx(self, path):
        import openpyxl
        from openpyxl.comments import Comment

        def _safe_sheet(name):
            bad = '\\/*?:[]'
            s = "".join("_" if c in bad else c for c in str(name))
            return s[:31] or "Sheet1"

        def _cell(val, col_type):
            s = "" if val is None else str(val).strip()
            if s == "":
                if col_type == "int":   return 0
                if col_type == "float": return 0.0
                if col_type == "bool":  return False
                return None
            if col_type == "int":
                try: return int(float(s))
                except (ValueError, TypeError): return s
            if col_type == "float":
                try: return float(s)
                except (ValueError, TypeError): return s
            if col_type == "bool":
                return s.lower() in ("true", "1", "yes")
            return s  # string / enum / array(comma-joined) / text_ref

        def _write_table(ws, df, cols_cfg):
            cols = list(df.columns)
            for ci, col in enumerate(cols, 1):
                cell = ws.cell(row=1, column=ci, value=str(col))
                note = (cols_cfg.get(col, {}) or {}).get("note", "")
                if note:
                    cell.comment = Comment(note, "JsonEditor")
            for ri, (_, row) in enumerate(df.iterrows(), 2):
                for ci, col in enumerate(cols, 1):
                    ct = (cols_cfg.get(col, {}) or {}).get("type", "string")
                    ws.cell(row=ri, column=ci, value=_cell(row[col], ct))

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default empty sheet

        for tname, tdf in self.manager.tables.items():
            cfg = self.manager.config.get(tname, {})
            ws = wb.create_sheet(title=_safe_sheet(tname))
            _write_table(ws, tdf, cfg.get("columns", {}))
            prefix = tname + "."
            for sub_full, sdf in self.manager.sub_tables.items():
                if not sub_full.startswith(prefix):
                    continue
                sub_name = sub_full[len(prefix):]
                ws2 = wb.create_sheet(title=_safe_sheet(sub_name))
                _write_table(
                    ws2, sdf,
                    cfg.get("sub_tables", {}).get(sub_name, {}).get("columns", {})
                )

        wb.save(path)

    # ── UI refresh ────────────────────────────────────────────────────────────

    def _refresh_ui(self):
        _cat_assign.clear()  # Reset category color assignments for new file
        self._tab_widget.blockSignals(True)
        self._tab_widget.clear()
        self._editors.clear()
        tables = list(self.manager.tables.keys())
        for tname in tables:
            self._editors[tname] = None
            self._tab_widget.addTab(QWidget(), tname)
        self._tab_widget.blockSignals(False)

        if tables:
            self._show_editor()
            QTimer.singleShot(0, lambda: self._ensure_editor(0))
        else:
            self._welcome.refresh(self.manager)
            self._show_welcome()

        self._update_sync()
        self.show_snackbar(f"已載入 {len(tables)} 個資料表")

    def _on_tab_changed(self, idx):
        self._ensure_editor(idx)

    def _ensure_editor(self, idx):
        if idx < 0 or idx >= self._tab_widget.count():
            return
        tname = self._tab_widget.tabText(idx)
        if self._editors.get(tname) is not None:
            return
        editor = TableEditor(tname, self.manager)
        editor.status_message.connect(
            lambda text, color: self.show_snackbar(text, color=color)
        )
        self._editors[tname] = editor          # guard before tab swap
        self._tab_widget.blockSignals(True)
        self._tab_widget.removeTab(idx)
        self._tab_widget.insertTab(idx, editor, tname)
        self._tab_widget.blockSignals(False)
        self._tab_widget.setCurrentIndex(idx)

    # ── Recent / search / config ──────────────────────────────────────────────

    def _show_recent_menu(self):
        recent = self.manager._recent_files
        if not recent:
            self.show_snackbar("沒有最近開啟的檔案"); return
        menu = QMenu(self)
        for path in recent:
            fname = os.path.basename(path)
            menu.addAction(f"{fname}  —  {os.path.dirname(path)}",
                           lambda p=path: self._load_recent(p))
        menu.addSeparator()
        menu.addAction("清除記錄", lambda: (
            setattr(self.manager, "_recent_files", []),
            self.manager.save_config(),
            self.show_snackbar("已清除最近記錄"),
        ))
        menu.exec(self.mapToGlobal(self.rect().topLeft()))

    def _show_search(self):
        query, ok = QInputDialog.getText(self, "搜尋", "搜尋所有資料表:")
        if not ok or not query.strip(): return
        results = self.manager.search_index(query.strip())
        if not results:
            self.show_snackbar("無結果"); return
        tname, is_sub, row_idx, _cols = results[0]
        if is_sub: return
        for i in range(self._tab_widget.count()):
            if self._tab_widget.tabText(i) == tname:
                self._tab_widget.setCurrentIndex(i)
                self._ensure_editor(i)
                editor = self._editors.get(tname)
                if editor:
                    df      = self.manager.tables[tname]
                    cls_val = df.at[row_idx, editor.cls_key]
                    editor.current_cls_val = cls_val
                    editor._load_cls_list()
                    editor._load_item_list()
                    editor._load_editor(row_idx)
                break
        self.show_snackbar(f"找到 {len(results)} 筆結果 (已跳至第一筆)")

    def open_config(self):
        idx = self._tab_widget.currentIndex()
        if idx < 0: return
        tname = self._tab_widget.tabText(idx)
        if tname not in self.manager.tables: return
        self._show_config_dialog(tname)

    def open_health_check(self):
        idx = self._tab_widget.currentIndex()
        if idx < 0:
            return
        tname = self._tab_widget.tabText(idx)
        if tname not in self.manager.tables:
            return
        issues = self.manager.health_check(tname)
        self._show_health_dialog(tname, issues)

    def _show_health_dialog(self, table_name, issues):
        dlg = QDialog(self)
        dlg.setWindowTitle(f"資料健檢 — {table_name}")
        dlg.setMinimumWidth(480)
        dlg.resize(560, 460)
        dlg.setStyleSheet(APP_QSS)
        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(0, 0, 0, 0); outer.setSpacing(0)

        errs = sum(1 for s, _ in issues if s == "error")
        warns = sum(1 for s, _ in issues if s == "warn")
        hdr = QLabel(
            "✓  未發現問題" if not issues
            else f"發現 {errs} 個錯誤、{warns} 個警告"
        )
        hdr.setStyleSheet(
            f"color:{_C['green'] if not issues else _C['txt']}; font-size:13px; "
            f"font-weight:600; background:{_C['sidebar']}; padding:12px 16px; "
            f"border-bottom:1px solid {_C['border']};"
        )
        outer.addWidget(hdr)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        content = QWidget(); content.setStyleSheet(f"background:{_C['panel']};")
        vlo = QVBoxLayout(content)
        vlo.setContentsMargins(14, 12, 14, 12); vlo.setSpacing(8)
        if not issues:
            ok = QLabel("主鍵唯一、子表 FK 都對得到母表，沒有空白主鍵/FK。")
            ok.setStyleSheet(f"color:{_C['txt2']}; font-size:12px; background:transparent;")
            ok.setWordWrap(True)
            vlo.addWidget(ok)
        else:
            for sev, msg in issues:
                color = _C["red"] if sev == "error" else _C["yellow"]
                row = QLabel(f'<span style="color:{color}">●</span>&nbsp;&nbsp;'
                             f'<span style="color:{_C["txt"]}">{msg}</span>')
                row.setWordWrap(True)
                row.setStyleSheet("background:transparent; font-size:12px;")
                vlo.addWidget(row)
        vlo.addStretch(1)
        scroll.setWidget(content)
        outer.addWidget(scroll, 1)

        bb = QWidget(); bb.setStyleSheet(f"background:{_C['sidebar']}; border-top:1px solid {_C['border']};")
        bl = QHBoxLayout(bb); bl.setContentsMargins(16, 10, 16, 10); bl.addStretch(1)
        btn = _mk_btn("關閉", "primary"); btn.setFixedHeight(32); btn.clicked.connect(dlg.accept)
        bl.addWidget(btn)
        outer.addWidget(bb)
        dlg.exec()

    def _show_config_dialog(self, table_name):
        _btn_ss = (
            f"background:{_C['input']}; border:1px solid {_C['border']}; "
            f"color:{_C['txtAcc']}; border-radius:5px; padding:3px 10px; text-align:left;"
        )

        # ── Helper: enum options editor button ────────────────────────────────
        def _make_opts_btn(parent_dlg, cur_opts, col_label, df_source=None, col_name_str=None):
            """Return (button, opts_store) where opts_store[0] is the live list."""
            opts_store = [list(cur_opts)]

            def _label():
                n = len(opts_store[0])
                return f"選項: {n}個  ✎" if n else "設定選項…"

            btn = QPushButton(_label())
            btn.setStyleSheet(_btn_ss)

            def _open():
                od = QDialog(parent_dlg)
                od.setWindowTitle(f"Enum 選項 — {col_label}")
                od.resize(340, 420)
                od.setStyleSheet(APP_QSS)
                ov = QVBoxLayout(od)
                ov.setContentsMargins(16, 16, 16, 16); ov.setSpacing(8)

                hint = QLabel("雙擊選項可編輯；拖曳可排序")
                hint.setStyleSheet(f"color:{_C['txt3']}; font-size:11px; background:transparent;")
                ov.addWidget(hint)

                lw = QListWidget()
                lw.setStyleSheet(
                    f"background:{_C['input']}; border:1px solid {_C['border']}; "
                    f"border-radius:5px; color:{_C['txt']};"
                )
                lw.addItems([str(o) for o in opts_store[0]])
                lw.setDragDropMode(QAbstractItemView.InternalMove)
                lw.setSelectionMode(QAbstractItemView.SingleSelection)
                ov.addWidget(lw, 1)

                inp_row = QWidget(); inp_row.setStyleSheet("background:transparent;")
                il = QHBoxLayout(inp_row); il.setContentsMargins(0, 0, 0, 0); il.setSpacing(6)
                inp = QLineEdit(); inp.setPlaceholderText("輸入新選項名稱")
                add_btn = _mk_btn("+ 新增", "primary"); add_btn.setFixedHeight(30)
                def _add():
                    t = inp.text().strip()
                    if t and not any(lw.item(i).text() == t for i in range(lw.count())):
                        lw.addItem(t); inp.clear()
                add_btn.clicked.connect(_add); inp.returnPressed.connect(_add)
                il.addWidget(inp, 1); il.addWidget(add_btn)
                ov.addWidget(inp_row)

                # Auto-collect button: scan df_source column for unique values
                if df_source is not None and col_name_str and col_name_str in df_source.columns:
                    auto_btn = _mk_btn("⟳ 從資料自動收集", "secondary"); auto_btn.setFixedHeight(30)
                    def _auto_collect():
                        existing = {lw.item(i).text() for i in range(lw.count())}
                        vals = df_source[col_name_str].dropna().astype(str).unique()
                        added = 0
                        for v in sorted(vals):
                            v = v.strip()
                            if v and v not in existing:
                                lw.addItem(v)
                                existing.add(v)
                                added += 1
                        if added == 0:
                            hint.setText("（所有現有值已包含在選項中）")
                        else:
                            hint.setText(f"已新增 {added} 個選項")
                    auto_btn.clicked.connect(_auto_collect)
                    ov.addWidget(auto_btn)

                del_btn = _mk_btn("刪除選取項目", "danger"); del_btn.setFixedHeight(30)
                def _del():
                    for it in lw.selectedItems():
                        lw.takeItem(lw.row(it))
                del_btn.clicked.connect(_del)
                ov.addWidget(del_btn)

                bb2 = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                bb2.accepted.connect(od.accept); bb2.rejected.connect(od.reject)
                ov.addWidget(bb2)

                if od.exec() == QDialog.Accepted:
                    opts_store[0] = [lw.item(ii).text() for ii in range(lw.count())]
                    btn.setText(_label())

            btn.clicked.connect(_open)
            return btn, opts_store

        def _col_row(col, cfg_cols, parent_dlg, df_source=None):
            """Return (row_widget, combo, opts_store, note_edit) for one column."""
            rw  = QWidget(); rw.setStyleSheet("background:transparent;")
            rlo = QHBoxLayout(rw)
            rlo.setContentsMargins(0, 0, 0, 0); rlo.setSpacing(6)
            cb = _NoscrollCombo()
            cb.addItems(["string", "int", "float", "bool", "enum", "text_ref", "array"])
            cur_type = cfg_cols.get(col, {}).get("type", "string")
            cb.setCurrentText(cur_type)

            cur_opts = cfg_cols.get(col, {}).get("options", [])
            opts_btn, opts_store = _make_opts_btn(
                parent_dlg, cur_opts, col,
                df_source=df_source, col_name_str=col
            )
            opts_btn.setVisible(cur_type == "enum")
            cb.currentTextChanged.connect(lambda t, ob=opts_btn: ob.setVisible(t == "enum"))

            note_edit = QTextEdit()
            note_edit.setAcceptRichText(False)
            note_edit.setFixedHeight(50)
            note_edit.setPlaceholderText("欄位備註（可換行；滑鼠停留欄位標題時顯示）")
            note_edit.setPlainText(cfg_cols.get(col, {}).get("note", ""))

            suggest_combo = _NoscrollCombo()
            suggest_combo.setToolTip("依此鄰欄的值過濾建議清單（空白為不啟用）")
            suggest_combo.setMaximumWidth(130)
            suggest_combo.addItem("(無建議)", "")
            if df_source is not None:
                for sc in df_source.columns:
                    if str(sc) != col:
                        suggest_combo.addItem(str(sc), str(sc))
            cur_sg = cfg_cols.get(col, {}).get("suggest_from", "")
            ix = suggest_combo.findData(cur_sg) if cur_sg else 0
            suggest_combo.setCurrentIndex(max(ix, 0))

            rlo.addWidget(cb, 0, Qt.AlignTop)
            rlo.addWidget(opts_btn, 1, Qt.AlignTop)
            rlo.addWidget(note_edit, 2)
            rlo.addWidget(suggest_combo, 0, Qt.AlignTop)
            return rw, cb, opts_store, note_edit, suggest_combo

        # ── Dialog ────────────────────────────────────────────────────────────
        dlg = QDialog(self)
        dlg.setWindowTitle(f"配置 — {table_name}")
        dlg.setMinimumWidth(560)
        dlg.resize(720, 680)
        dlg.setStyleSheet(APP_QSS)

        cfg      = self.manager.config.get(table_name, {})
        df       = self.manager.tables[table_name]
        cols     = list(df.columns)
        cols_cfg = cfg.get("columns", {})

        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        content = QWidget(); content.setStyleSheet(f"background:{_C['panel']};")
        vlo = QVBoxLayout(content)
        vlo.setContentsMargins(16, 16, 16, 16)
        vlo.setSpacing(12)
        scroll.setWidget(content)
        outer.addWidget(scroll, 1)

        # ── Main table keys ───────────────────────────────────────────────────
        def _sec(text):
            lbl = QLabel(text.upper())
            lbl.setStyleSheet(
                f"color:{_C['txt3']}; font-size:10px; font-weight:600; "
                f"letter-spacing:1px; background:transparent;"
            )
            return lbl

        def _form_row(label_text, widget):
            w = QWidget(); w.setStyleSheet("background:transparent;")
            h = QHBoxLayout(w); h.setContentsMargins(0,0,0,0); h.setSpacing(10)
            lbl = QLabel(label_text)
            lbl.setFixedWidth(160)
            lbl.setStyleSheet(f"color:{_C['txt2']}; background:transparent;")
            h.addWidget(lbl); h.addWidget(widget, 1)
            return w

        vlo.addWidget(_sec("主表設定"))
        pk_var  = _NoscrollCombo(); pk_var.addItems(cols)
        cls_var = _NoscrollCombo(); cls_var.addItems(cols)
        pk_var.setCurrentText(cfg.get("primary_key",         cols[0] if cols else ""))
        cls_var.setCurrentText(cfg.get("classification_key", cols[0] if cols else ""))
        vlo.addWidget(_form_row("Primary Key",         pk_var))
        vlo.addWidget(_form_row("Classification Key",  cls_var))

        # ── Browse helper (used by image folder + text-ref) ──────────────────
        def _browse_row(edit, is_folder=False):
            row = QWidget(); row.setStyleSheet("background:transparent;")
            rl = QHBoxLayout(row); rl.setContentsMargins(0,0,0,0); rl.setSpacing(4)
            btn = QPushButton("…"); btn.setFixedWidth(32)
            btn.setStyleSheet(
                f"background:{_C['card']}; border:1px solid {_C['border']}; "
                f"color:{_C['txt']}; border-radius:5px;"
            )
            def _browse():
                base = os.path.dirname(self.manager.json_path) if self.manager.json_path else ""
                if is_folder:
                    from PySide6.QtWidgets import QFileDialog as _QFD
                    p = _QFD.getExistingDirectory(dlg, "選擇資料夾", base)
                else:
                    from PySide6.QtWidgets import QFileDialog as _QFD
                    p, _ = _QFD.getOpenFileName(dlg, "選擇檔案", base, "JSON (*.json)")
                if p:
                    try: p = os.path.relpath(p, base) if base else p
                    except ValueError: pass
                    edit.setText(p)
            btn.clicked.connect(_browse)
            rl.addWidget(edit, 1); rl.addWidget(btn)
            return row

        # Image base folder (first row)
        img_folder_edit = QLineEdit(cfg.get("image_preview", {}).get("base_folder", ""))
        img_folder_edit.setPlaceholderText("圖片根目錄（相對路徑或絕對路徑）")
        vlo.addWidget(_form_row("Image 資料夾路徑", _browse_row(img_folder_edit, is_folder=True)))

        img_ext_edit = QLineEdit(cfg.get("image_preview", {}).get("ext", ""))
        img_ext_edit.setPlaceholderText("副檔名，例如 .png")
        vlo.addWidget(_form_row("Image 副檔名", img_ext_edit))

        # ── Image path segments builder (second row) ───────────────────────────
        _img_segs_container = QWidget(); _img_segs_container.setStyleSheet("background:transparent;")
        _img_segs_lo = QVBoxLayout(_img_segs_container)
        _img_segs_lo.setContentsMargins(0, 0, 0, 0); _img_segs_lo.setSpacing(2)
        _img_segs_rows = []  # list of (type_cb, val_stack, col_combo, lit_edit, row_w)

        _img_rows_lo = QVBoxLayout()
        _img_segs_lo.addLayout(_img_rows_lo)

        def _add_img_seg_row(seg_type="col", seg_value=""):
            row_w = QWidget(); row_w.setStyleSheet("background:transparent;")
            rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(4)
            type_cb = _NoscrollCombo()
            type_cb.addItems(["欄位", "字串"])
            type_cb.setFixedWidth(58)
            col_combo = _NoscrollCombo()
            col_combo.addItems(cols)
            lit_edit = QLineEdit()
            lit_edit.setPlaceholderText("字串值")
            val_stack = QStackedWidget()
            val_stack.setStyleSheet("background:transparent;")
            val_stack.addWidget(col_combo)   # index 0
            val_stack.addWidget(lit_edit)    # index 1
            if seg_type == "lit":
                type_cb.setCurrentIndex(1)
                val_stack.setCurrentIndex(1)
                lit_edit.setText(seg_value)
            else:
                type_cb.setCurrentIndex(0)
                val_stack.setCurrentIndex(0)
                if seg_value in cols:
                    col_combo.setCurrentText(seg_value)
            type_cb.currentIndexChanged.connect(val_stack.setCurrentIndex)
            del_btn = QPushButton("−")
            del_btn.setFixedSize(24, 24)
            del_btn.setStyleSheet(
                f"background:{_C['card']}; border:1px solid {_C['border']}; "
                f"color:{_C['red']}; border-radius:4px; font-weight:700;"
            )
            rl.addWidget(type_cb)
            rl.addWidget(val_stack, 1)
            rl.addWidget(del_btn)
            entry = (type_cb, val_stack, col_combo, lit_edit, row_w)
            _img_segs_rows.append(entry)
            _img_rows_lo.addWidget(row_w)
            def _del_seg(e=entry, w=row_w):
                if e in _img_segs_rows:
                    _img_segs_rows.remove(e)
                w.hide()
                w.deleteLater()
            del_btn.clicked.connect(_del_seg)

        # Load existing segments (backward compat with old "col" key)
        _cur_img_segs = cfg.get("image_preview", {}).get("path_segments", [])
        if not _cur_img_segs:
            _old_img_col = cfg.get("image_preview", {}).get("col", "")
            if _old_img_col:
                _cur_img_segs = [{"type": "col", "col": _old_img_col}]
        for _s in _cur_img_segs:
            if _s.get("type") == "col":
                _add_img_seg_row("col", _s.get("col", ""))
            else:
                _add_img_seg_row("lit", _s.get("value", ""))

        _add_seg_btn = QPushButton("＋ 加段")
        _add_seg_btn.setFixedHeight(26)
        _add_seg_btn.setStyleSheet(
            f"background:{_C['card']}; border:1px solid {_C['border']}; "
            f"color:{_C['txt2']}; border-radius:5px; font-size:11px;"
        )
        _add_seg_btn.clicked.connect(lambda: _add_img_seg_row("col", ""))
        _img_segs_lo.addWidget(_add_seg_btn)

        vlo.addWidget(_form_row("Image 路徑結構", _img_segs_container))

        # External text-ref JSON path row
        _trs_cfg = cfg.get("text_ref_source", {})
        text_ref_edit = QLineEdit(_trs_cfg.get("json_path", ""))
        text_ref_edit.setPlaceholderText("外部文字表路徑（相對路徑或絕對路徑）")
        vlo.addWidget(_form_row("外部文字表路徑", _browse_row(text_ref_edit, is_folder=False)))

        trs_key_edit = QLineEdit(_trs_cfg.get("key_col", "TextID"))
        trs_key_edit.setPlaceholderText("key 欄名（預設 TextID）")
        trs_val_edit = QLineEdit(_trs_cfg.get("val_col", "TextContent"))
        trs_val_edit.setPlaceholderText("value 欄名（預設 TextContent）")
        vlo.addWidget(_form_row("  文字表 Key 欄", trs_key_edit))
        vlo.addWidget(_form_row("  文字表 Val 欄", trs_val_edit))

        sep1 = QFrame(); sep1.setFixedHeight(1)
        sep1.setStyleSheet(f"background:{_C['border']}; border:none;")
        vlo.addWidget(sep1)
        vlo.addWidget(_sec("主表欄位類型"))

        main_col_widgets: dict[str, tuple] = {}  # col → (cb, opts_store)
        for col in cols:
            rw, cb, opts_store, note_edit, suggest_combo = _col_row(col, cols_cfg, dlg, df_source=df)
            vlo.addWidget(_form_row(f"  {col}", rw))
            main_col_widgets[col] = (cb, opts_store, note_edit, suggest_combo)

        # ── Sub-tables ────────────────────────────────────────────────────────
        prefix = table_name + "."
        sub_keys = [k for k in self.manager.sub_tables if k.startswith(prefix)]
        sub_widgets: dict[str, dict] = {}  # tab_name → {fk_edit, col_combos}

        if sub_keys:
            sep2 = QFrame(); sep2.setFixedHeight(1)
            sep2.setStyleSheet(f"background:{_C['border']}; border:none;")
            vlo.addSpacing(4); vlo.addWidget(sep2)
            vlo.addWidget(_sec("從表設定"))

            sub_cfg_root = cfg.get("sub_tables", {})

            for full_key in sub_keys:
                tab_name     = full_key[len(prefix):]
                sub_df       = self.manager.sub_tables[full_key]
                sub_cfg      = sub_cfg_root.get(tab_name, {})
                sub_cols_cfg = sub_cfg.get("columns", {})

                shdr = QLabel(f"▸  {tab_name}")
                shdr.setStyleSheet(
                    f"color:{_C['txtAcc']}; font-size:12px; font-weight:600; "
                    f"background:transparent; padding-top:6px;"
                )
                vlo.addWidget(shdr)

                fk_edit = QLineEdit()
                fk_edit.setPlaceholderText("foreign_key 欄位名稱")
                fk_edit.setText(sub_cfg.get("foreign_key", ""))
                vlo.addWidget(_form_row("  Foreign Key", fk_edit))

                st_note_edit = QLineEdit()
                st_note_edit.setPlaceholderText("這張子表的用途說明（滑鼠停在分頁標題時顯示）")
                st_note_edit.setText(sub_cfg.get("note", ""))
                vlo.addWidget(_form_row("  說明 / 備註", st_note_edit))

                col_combos: dict[str, tuple] = {}
                for scol in list(sub_df.columns):
                    rw, cb, opts_store, note_edit, suggest_combo = _col_row(scol, sub_cols_cfg, dlg, df_source=sub_df)
                    vlo.addWidget(_form_row(f"    {scol}", rw))
                    col_combos[scol] = (cb, opts_store, note_edit, suggest_combo)

                sub_widgets[tab_name] = {"fk_edit": fk_edit, "note_edit": st_note_edit,
                                         "col_combos": col_combos}

        else:
            # Inform user if no sub-tables detected
            no_sub = QLabel("（此表格在 JSON 中無巢狀陣列資料，故無從表）")
            no_sub.setStyleSheet(f"color:{_C['txt3']}; font-size:11px; background:transparent;")
            vlo.addWidget(no_sub)

        vlo.addStretch(1)

        # ── Buttons ───────────────────────────────────────────────────────────
        bb_w = QWidget()
        bb_w.setStyleSheet(
            f"background:{_C['sidebar']}; border-top:1px solid {_C['border']};"
        )
        bb_lo = QHBoxLayout(bb_w)
        bb_lo.setContentsMargins(16, 10, 16, 10)
        bb_lo.setSpacing(8)
        bb_lo.addStretch(1)
        btn_ok  = _mk_btn("套用", "primary"); btn_ok.setFixedHeight(34)
        btn_can = _mk_btn("取消");             btn_can.setFixedHeight(34)
        btn_ok.clicked.connect(dlg.accept)
        btn_can.clicked.connect(dlg.reject)
        bb_lo.addWidget(btn_can); bb_lo.addWidget(btn_ok)
        outer.addWidget(bb_w)

        if dlg.exec() != QDialog.Accepted:
            return

        # ── Apply ─────────────────────────────────────────────────────────────
        cfg["primary_key"]        = pk_var.currentText()
        cfg["classification_key"] = cls_var.currentText()
        img_folder_val = img_folder_edit.text().strip()
        _new_segs = []
        for (tcb, vstk, ccb, ledit, rw) in _img_segs_rows:
            if tcb.currentIndex() == 0:
                _new_segs.append({"type": "col", "col": ccb.currentText()})
            else:
                _new_segs.append({"type": "lit", "value": ledit.text()})
        img_ext_val = img_ext_edit.text().strip()
        if _new_segs:
            cfg["image_preview"] = {"path_segments": _new_segs}
            if img_folder_val:
                cfg["image_preview"]["base_folder"] = img_folder_val
            if img_ext_val:
                cfg["image_preview"]["ext"] = img_ext_val
        else:
            cfg.pop("image_preview", None)

        text_ref_path = text_ref_edit.text().strip()
        if text_ref_path:
            cfg["text_ref_source"] = {
                "json_path": text_ref_path,
                "key_col":   trs_key_edit.text().strip() or "TextID",
                "val_col":   trs_val_edit.text().strip() or "TextContent",
            }
        else:
            cfg.pop("text_ref_source", None)
        def _build_col_entry(t, opts_store, note="", suggest_from=""):
            entry = {"type": t}
            if t == "enum" and opts_store[0]:
                entry["options"] = opts_store[0]
            note = (note or "").strip()
            if note:
                entry["note"] = note
            sg = (suggest_from or "").strip()
            if sg:
                entry["suggest_from"] = sg
            return entry

        cfg.setdefault("columns", {})
        for col, (cb, opts_store, note_edit, suggest_combo) in main_col_widgets.items():
            cfg["columns"][col] = _build_col_entry(
                cb.currentText(), opts_store, note_edit.toPlainText(), suggest_combo.currentData()
            )

        cfg.setdefault("sub_tables", {})
        for tab_name, data in sub_widgets.items():
            st = cfg["sub_tables"].setdefault(tab_name, {})
            fk = data["fk_edit"].text().strip()
            if fk:
                st["foreign_key"] = fk
            st_note = data["note_edit"].text().strip()
            if st_note:
                st["note"] = st_note
            else:
                st.pop("note", None)
            st.setdefault("columns", {})
            for scol, (scb, opts_store, note_edit, suggest_combo) in data["col_combos"].items():
                st["columns"][scol] = _build_col_entry(
                    scb.currentText(), opts_store, note_edit.toPlainText(), suggest_combo.currentData()
                )

        self.manager.config[table_name] = cfg
        self.manager.save_config()
        editor = self._editors.get(table_name)
        if editor:
            editor.reload_after_config()
        self.show_snackbar("配置已套用")

    # ── Window close ──────────────────────────────────────────────────────────

    def closeEvent(self, event):
        if self.manager.dirty:
            ans = QMessageBox.question(
                self, "未儲存變更", "有未儲存的變更，是否儲存？",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            if ans == QMessageBox.Cancel:
                event.ignore(); return
            if ans == QMessageBox.Save:
                self.save_file()
        event.accept()


# ── Entry ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_QSS)
    window = App()
    window.show()
    sys.exit(app.exec())
