from PyQt5 import QtWidgets

from File import is_valid_excel
from PyQt5.QtCore import Qt


class LineEditInjector:
    def __init__(self, line_edit, auto_inject=True):
        self.lineEdit = line_edit
        if auto_inject:
            self.inject_drag_file()

    def inject_drag_file(self):
        self.lineEdit.setDragEnabled(True)
        self.lineEdit.dragEnterEvent = self._drag_enter_event
        self.lineEdit.dragMoveEvent = self._drag_move_event
        self.lineEdit.dropEvent = self._drop_event

    def _drag_enter_event(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def _drag_move_event(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def _drop_event(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].isLocalFile():  # si solo contiene un archivo
                path = str(urls[0].toLocalFile())
                if path.endswith(".xlsx"):
                    if is_valid_excel(path):
                        event.accept()
                        camino = str(urls[0].toLocalFile())
                        self.lineEdit.setText(camino)
                        self.lineEdit.setStyleSheet(None)
                    else:
                        event.ignore()
                        QtWidgets.QMessageBox.critical(None, "ERROR",
                                                       "Este excel parece que no contiene las estadísticas de "
                                                       "Facebook :(",
                                                       QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
                else:
                    event.ignore()
                    QtWidgets.QMessageBox.critical(None, "ERROR",
                                                   "El archivo debe ser una libro excel con la extensión .xlsx",
                                                   QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
            else:
                event.ignore()
        else:
            event.ignore()
