"""Cell delegates: combo-boxes for the drop-down columns, plus a stream-lookup
combo whose values come from the loaded HMB."""
from __future__ import annotations

from PySide6.QtWidgets import QStyledItemDelegate, QComboBox

import engine_api as api


class ComboDelegate(QStyledItemDelegate):
    """Editable combo for a fixed value list (searchable)."""
    def __init__(self, values, parent=None, editable=True):
        super().__init__(parent)
        self._values = list(values)
        self._editable = editable

    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.setEditable(self._editable)
        cb.addItems([""] + [v for v in self._values if v != ""])
        return cb

    def setEditorData(self, editor, index):
        cur = index.data() or ""
        i = editor.findText(str(cur))
        if i >= 0:
            editor.setCurrentIndex(i)
        else:
            editor.setEditText(str(cur))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), 0x2)  # Qt.EditRole


class StreamDelegate(ComboDelegate):
    """Combo whose value list is refreshed from the loaded HMB streams."""
    def __init__(self, stream_provider, parent=None):
        super().__init__([], parent, editable=True)
        self._provider = stream_provider

    def createEditor(self, parent, option, index):
        self._values = self._provider() or []
        return super().createEditor(parent, option, index)


def install_delegates(view, model, stream_provider):
    """Attach the right delegate to each drop-down column of the grid view."""
    dd = api.dropdown_lists()
    headers = model.headers
    for col_name, values in dd.items():
        if col_name in headers:
            view.setItemDelegateForColumn(headers.index(col_name),
                                          ComboDelegate(values, view))
    if "Stream Lookup" in headers:
        view.setItemDelegateForColumn(headers.index("Stream Lookup"),
                                      StreamDelegate(stream_provider, view))
