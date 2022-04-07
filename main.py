import os
import sys
import shutil
import pandas as pd

from PyQt5.QtWidgets import (
    QTableWidget, QApplication, QGridLayout, QWidget, QTableWidgetItem,
    QMainWindow, QStatusBar, QToolBar, QLineEdit, QFileDialog, QComboBox,
    QLabel, QAbstractScrollArea, QShortcut, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from parseXML import parseXML
from lxml.etree import XMLSyntaxError
from clipboard import setClipboard


class MainWindow(QMainWindow):
    def __init__(self, parent=None, dropfile=None):
        super().__init__(parent)

        open_shortcut = QShortcut(QKeySequence(self.tr("Ctrl+O")), self)
        open_shortcut.activated.connect(self.fileSelect)

        self.setWindowTitle("XML para Excel")
        self.resize(800, 600)
        self.setAcceptDrops(True)
        self._createMenu()
        self._addWidgets()
        self._createStatusBar()
        self._createToolbar()
        self.show()
        if dropfile:
            self._filepath.setText(dropfile)
            try:
                self.readXML()
            except XMLSyntaxError:
                self._filepath.setText("")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            path = e.mimeData().urls()[0].path()[1:]
            if path.endswith('.xml'):
                e.accept()
            else:
                e.ignore()
        else:
            e.ignore()

    def dropEvent(self, e):
        path = e.mimeData().urls()[0].path()[1:]
        self._filepath.setText(path)
        self.readXML()

    def _createMenu(self):
        self.menu = self.menuBar().addMenu("&Menu")
        self.menu.addAction('&Abrir', self.fileSelect)
        self.menu.addAction('&Converter', self.readXML)
        self.menu.addAction('&Exportar', self.saveAs)
        self.menu.addAction('&Fechar', self.close)

    def _addWidgets(self):
        self._filepath = QLineEdit(self)
        self._filepath.setPlaceholderText("Insira o caminho para o arquivo "
                                   "ou um texto XML")
        self._filepath.setMaxLength(999999)

        self._mainWidget = MainWidget(self)
        self.setCentralWidget(self._mainWidget)

    def _createStatusBar(self):
        self.status = QStatusBar()
        self.setStatusBar(self.status)

    def _createToolbar(self):
        tools = QToolBar()
        self.addToolBar(tools)
        tools.setMovable(False)
        tools.setContextMenuPolicy(Qt.ContextMenuPolicy.PreventContextMenu)
        tools.addAction('Selecionar', self.fileSelect)
        tools.addAction('Converter', self.readXML)
        tools.addAction('Exportar', self.saveAs)
        tools.addWidget(self._filepath)

    def fileSelect(self):
        dlg = QFileDialog()
        dlg.ExistingFile
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        filepath = dlg.getOpenFileName(
            self, self.tr("Selecione o XML de retorno"), "",
            self.tr("Arquivos XML (*.xml)")
        )
        if filepath[0]:
            self._filepath.setText(filepath[0])
            self.readXML()

    def readXML(self):
        path = self._filepath.text()
        try:
            self.dfCodigos, self.dfGuias, self.dfProcedimentos, self.dfRelatorio = parseXML(path)
        except XMLSyntaxError:
            dlg = QMessageBox()
            dlg.setText('Erro ao ler o arquivo XML!')
            dlg.exec()
            return
        self.dataframes = [
            self.dfRelatorio,
            self.dfCodigos,
            self.dfGuias,
            self.dfProcedimentos,
        ]
        self._mainWidget._visualize.setEnabled(True)
        self._mainWidget._visualize.setCurrentIndex(2)
        self._mainWidget._visualize.setCurrentIndex(0)

    def saveAs(self):
        df = self.dataframes
        if df:
            dlg = QFileDialog()
            dlg.AcceptSave
            filename = dlg.getSaveFileName(
                self, self.tr("Salvar como"), "", self.tr("Excel (*.xlsx)")
            )
            if filename[0]:
                temp = os.getenv("temp") + "/temp.xlsx"
                with pd.ExcelWriter(temp) as writer:
                    df[0].to_excel(writer, sheet_name="Relatório")
                    df[1].to_excel(writer, sheet_name="Motivos de glosa")
                    df[2].to_excel(writer, sheet_name="Guias glosadas")
                    df[3].to_excel(writer, sheet_name="Procedimentos glosados")

                shutil.move(temp, filename[0])
                self.dbr_file = filename[0]
                self.status.showMessage("File saved at: %s" % filename[0])


class MainWidget(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.tableHeaders = {
            0: [
                'Código Glosa',
                'Descrição',
                'Nome',
                'Número Guia',
                'Data',
                'Valor Glosa',
                'Protocolo',
            ],
            1: [
                "Protocolo",
                "Código",
                "Descrição",
                "Valor Glosa",
            ],
            2: [
                "Protocolo",
                "Número Guia",
                "Data Inicial",
                "Data Final",
                "Matrícula",
                "Nome",
                "Valor Glosa",
                "Código Glosa",
                "Descrição",
            ],
            3: [
                "Protocolo",
                "Número Guia",
                "Data",
                "Matrícula",
                "Senha",
                "Nome",
                "Procedimento",
                "Valor Glosa",
                "Código Glosa",
                "Descrição",
            ]
        }

        open_shortcut = QShortcut(QKeySequence(self.tr("Ctrl+C")), self)
        open_shortcut.activated.connect(self.copySum)

        self.layout = QGridLayout(self)
        self._addWidgets()
        self.show()

    def _addWidgets(self):
        self._visualize = QComboBox(self)
        self._visualize.addItems([
            "Relatório",
            "Motivos de glosa",
            "Guias glosadas",
            "Procedimentos glosados",
        ])
        self._visualize.setEnabled(False)
        self._glosas = QLabel("Glosas: R$ ", parent=self)
        self._glosas.hide()
        self.sum = QLabel("Total selecionado: R$ ")
        self.table = QTableWidget(0, 4, parent=self)
        self.table.setHorizontalHeaderLabels(
            self.tableHeaders[0])
        self.table.setSizeAdjustPolicy(
            QAbstractScrollArea.AdjustToContents)

        self.layout.addWidget(self._visualize, 0, 0)
        self.layout.addWidget(self._glosas, 0, 1, 1, 2)
        self.layout.addWidget(self.table, 1, 0, 1, 3)
        self.layout.addWidget(self.sum, 2, 0)

        self.setLayout(self.layout)

        self._visualize.currentIndexChanged.connect(self.changeVisualization)
        self.table.itemSelectionChanged.connect(self.calcSum)

    def changeVisualization(self):
        index = self.sender().currentIndex()
        headers = self.tableHeaders[index]
        df = self.parent.dataframes[index]

        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(headers)
        row = 0
        for line in df.iterrows():
            column = 0
            for value in line[1]:
                self.table.setItem(
                    row, column, QTableWidgetItem(str(value))
                )
                column += 1
            row += 1
        self.table.resizeColumnsToContents()

    def calcSum(self):
        try:
            col = self.table.selectedIndexes()[0].column()
            header = self.table.horizontalHeaderItem(col).text()
            if header == "Valor Glosa":
                values = self.table.selectedItems()
                val = sum([float(x.text()) for x in values])
                self.sum.setText(f"Total selecionado: R$ {val:.2f}")
        except IndexError:
            pass

    def copySum(self):
        col = self.table.selectedIndexes()[0].column()
        header = self.table.horizontalHeaderItem(col).text()
        if header != "Valor Glosa":
            tblrange = self.table.selectedRanges()[0]
            rows = tblrange.rowCount()
            cols = tblrange.columnCount()
            selection = self.table.selectedItems()
            text = [x.text().replace(".", ",") for x in selection]
            text = ['\t'.join(text[i*cols:(i+1)*cols]) for i in range(rows)]
            setClipboard('\n'.join(text))
            return
        text = self.sum.text()
        text = text.replace("Total selecionado: R$ ", "")
        text = text.replace(".", ",")
        setClipboard(text)


if __name__ == '__main__':
    dropfile = sys.argv[1] if len(sys.argv) > 1 else None
    app = QApplication([])
    window = MainWindow(None, dropfile)
    window.show()
    app.exec_()
