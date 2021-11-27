import sys
from pathlib import Path

from PyQt5 import QtWidgets, QtCore, QtGui
from openpyxl import load_workbook

from py_from_ui.MainFrame import Ui_MainWindow
from utiles.DragAndDropLineEdit import LineEditInjector

from File import open_file_explorer, obtener_posts, save, is_valid_excel
from Config import Config
from resources import resources

try:
    # Include in try/except block if you're also targeting Mac/Linux
    from PyQt5.QtWinExtras import QtWin

    myappid = 'mycompany.myproduct.subproduct.version'
    QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class Main(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QWidget.__init__(self, parent=None)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("SJ Stats")
        self.cnf = Config()

        self.get_spinners_initial_value()
        self.install_drag_and_drop_event()
        self.load_events()

        self.ui.lineEditPath.setToolTip("Puede arrastrar el archivo hasta aquí")

        self.ui.spinnerTop.setFocus()
        self.ui.spinnerTop.setToolTip("Top de post a mostrar en el excel a exportar")

        self.ui.buttomExportar.setIcon(QtGui.QIcon(":/icons/export"))
        self.ui.buttomExportar.setText("Exportar")
        self.ui.buttomExportar.setToolTip("Exporta un resumen de las estadísticas de Facebook")

        self.ui.butttomOpenFile.setToolTip("Buscar el excel con el explorador de archivos")
        self.ui.butttomOpenFile.setText("")
        self.ui.butttomOpenFile.setIcon(QtGui.QIcon(":/icons/search"))

    # le añado la propiedad de poder arrastrar los archivos aquí
    def install_drag_and_drop_event(self):
        LineEditInjector(self.ui.lineEditPath)
        self.ui.lineEditPath.installEventFilter(self)

    def get_spinners_initial_value(self):
        self.comentarios_old_value = self.ui.spinnerComentarios.value()
        self.compartido_old_value = self.ui.spinnerCompartidos.value()
        self.reacciones_old_value = self.ui.spinnerReacciones.value()

        self.interacciones_old_value = self.ui.spinnerInteracciones.value()
        self.match_audience_old_value = self.ui.spinnerMatchAudience.value()
        self.engagement_old_value = self.ui.spinnerEngagement.value()
        self.alcance_old_value = self.ui.spinnerAlcance.value()

    def load_events(self):
        self.ui.butttomOpenFile.clicked.connect(self.open_file)
        self.ui.buttomExportar.clicked.connect(self.export_file)

        self.ui.spinnerCompartidos.valueChanged.connect(self.validar_spinners_interacciones)
        self.ui.spinnerReacciones.valueChanged.connect(self.validar_spinners_interacciones)
        self.ui.spinnerComentarios.valueChanged.connect(self.validar_spinners_interacciones)

        self.ui.spinnerInteracciones.valueChanged.connect(self.validar_spinners_completo)
        self.ui.spinnerEngagement.valueChanged.connect(self.validar_spinners_completo)
        self.ui.spinnerMatchAudience.valueChanged.connect(self.validar_spinners_completo)
        self.ui.spinnerAlcance.valueChanged.connect(self.validar_spinners_completo)

    def validar_spinners_completo(self):
        interacciones = round(self.ui.spinnerInteracciones.value(), 2)
        match_audience = round(self.ui.spinnerMatchAudience.value(), 2)
        engagement = round(self.ui.spinnerEngagement.value(), 2)
        alcance = round(self.ui.spinnerAlcance.value(), 2)

        if interacciones + match_audience + engagement + alcance > 1:
            spinner_changed, value = self.get_spinner_completo_change()
            spinner_changed.setValue(value)
            QtWidgets.QMessageBox.critical(self, "ERROR",
                                           "La suma de los valores no puede ser mayor que 1",
                                           QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
        else:
            if interacciones != 0:
                self.interacciones_old_value = interacciones
            if match_audience != 0:
                self.match_audience_old_value = match_audience
            if engagement != 0:
                self.engagement_old_value = engagement
            if alcance != 0:
                self.alcance_old_value = alcance

    def get_spinner_completo_change(self):
        interacciones = round(self.ui.spinnerInteracciones.value(), 2)
        match_audience = round(self.ui.spinnerMatchAudience.value(), 2)
        engagement = round(self.ui.spinnerEngagement.value(), 2)
        alcance = round(self.ui.spinnerAlcance.value(), 2)

        if interacciones != self.interacciones_old_value:
            return self.ui.spinnerInteracciones, self.interacciones_old_value
        elif match_audience != self.match_audience_old_value:
            return self.ui.spinnerMatchAudience, self.match_audience_old_value
        elif engagement != self.engagement_old_value:
            return self.ui.spinnerEngagement, self.engagement_old_value
        elif alcance != self.alcance_old_value:
            return self.ui.spinnerAlcance, self.alcance_old_value

    def validar_spinners_interacciones(self):
        compartidos = round(self.ui.spinnerCompartidos.value(), 2)
        comentarios = round(self.ui.spinnerComentarios.value(), 2)
        reacciones = round(self.ui.spinnerReacciones.value(), 2)
        if comentarios + compartidos + reacciones > 1:
            spinner_changed, value = self.get_spinner_interacciones_change()
            spinner_changed.setValue(value)
            QtWidgets.QMessageBox.critical(self, "ERROR",
                                           "La suma de los valores no puede ser mayor que 1",
                                           QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
        else:
            if reacciones != 0:
                self.reacciones_old_value = reacciones
            if compartidos != 0:
                self.compartido_old_value = compartidos
            if comentarios != 0:
                self.comentarios_old_value = comentarios

    def get_spinner_interacciones_change(self):
        compartidos = round(self.ui.spinnerCompartidos.value(), 2)
        comentarios = round(self.ui.spinnerComentarios.value(), 2)
        reacciones = round(self.ui.spinnerReacciones.value(), 2)

        if comentarios != self.comentarios_old_value:
            return self.ui.spinnerComentarios, self.comentarios_old_value
        elif compartidos != self.compartido_old_value:
            return self.ui.spinnerCompartidos, self.compartido_old_value
        elif reacciones != self.reacciones_old_value:
            return self.ui.spinnerReacciones, self.reacciones_old_value

    def open_file(self):
        filename = open_file_explorer()
        if len(filename) > 0:
            if is_valid_excel(filename):
                self.ui.lineEditPath.setText(filename)
                self.ui.lineEditPath.setStyleSheet(None)
            else:
                QtWidgets.QMessageBox.critical(self, "ERROR",
                                               "Este excel parece que no contiene las estadísticas de "
                                               "Facebook :(",
                                               QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)

    def export_file(self):
        self.cnf.TOP = round(self.ui.spinnerTop.value(), 2)
        self.cnf.SHARED_WEIGHT = round(self.ui.spinnerCompartidos.value(), 2)
        self.cnf.LIKE_WEIGHT = round(self.ui.spinnerReacciones.value(), 2)
        self.cnf.COMMENT_WEIGHT = round(self.ui.spinnerComentarios.value(), 2)
        self.cnf.INTERACCIONES_WEIGHT = round(self.ui.spinnerInteracciones.value(), 2)
        self.cnf.MATCH_AUDIENCE_WEIGHT = round(self.ui.spinnerMatchAudience.value(), 2)
        self.cnf.ENGAGEMENT_WEIGHT = round(self.ui.spinnerEngagement.value(), 2)
        self.cnf.ALCANCE_WEIGHT = round(self.ui.spinnerAlcance.value(), 2)

        file_path = self.ui.lineEditPath.text()
        if file_path and is_valid_excel(file_path):
            top = self.ui.spinnerTop.value()
            self.save_workbook(file_path, top=top)
        else:
            QtWidgets.QMessageBox.critical(self, "ERROR",
                                           "Tiene que introducir un excel válido de Facebook :)",
                                           QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
            self.ui.lineEditPath.setStyleSheet("QLineEdit {border-width: 1px; border-style: solid; "
                                               "border-color: red red red red;}")
            self.ui.lineEditPath.setFocus()

    def save_workbook(self, path, top=3):
        parent_directory = Path(path).parent
        wb = load_workbook(path)
        posts = obtener_posts(wb)
        try:
            save(posts, parent_directory, top=top)
            QtWidgets.QMessageBox.information(self, "Info",
                                              "Archivo exportado correctamente!",
                                              QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok
                                              )
        except:
            QtWidgets.QMessageBox.critical(self, "ERROR",
                                           "Al parecer ya fue exportado el archivo excel y está abierto, por favor "
                                           "ciérrelo, y vuélvalo a intentar",
                                           QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)

    def eventFilter(self, source, event):
        if not event.type() == QtCore.QEvent.WindowDeactivate:
            if event.type() == QtCore.QEvent.FocusOut and source is self.ui.lineEditPath:
                path = self.ui.lineEditPath.text().strip()
                if len(path) != 0:
                    if not is_valid_excel(path):
                        self.ui.lineEditPath.setStyleSheet("QLineEdit {border-width: 1px; border-style: solid; "
                                                           "border-color: red red red red;}")
                    else:
                        self.ui.lineEditPath.setStyleSheet(None)
                else:
                    self.ui.lineEditPath.setStyleSheet("QLineEdit {border-width: 1px; border-style: solid; "
                                                       "border-color: red red red red;}")
            # return true here to bypass default behaviour
        return super(Main, self).eventFilter(source, event)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    app.setWindowIcon(QtGui.QIcon(":/icons/logo"))
    myapp = Main()
    myapp.show()
    sys.exit(app.exec_())
