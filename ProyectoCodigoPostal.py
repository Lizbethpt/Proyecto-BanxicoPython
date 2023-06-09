import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTreeWidgetItem, QTreeWidget, QGraphicsScene
from PyQt5 import QtCore, QtGui, QtWidgets
from docx import Document
from openpyxl import Workbook
import win32com.client as win32
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(594, 464)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.treeWidget = QTreeWidget(self.centralwidget)
        self.treeWidget.setGeometry(QtCore.QRect(10, 160, 281, 241))
        self.treeWidget.setObjectName("treeWidget")
        self.BtnAbrir = QtWidgets.QPushButton(self.centralwidget)
        self.BtnAbrir.setGeometry(QtCore.QRect(10, 10, 75, 23))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.BtnAbrir.setFont(font)
        self.BtnAbrir.setObjectName("BtnAbrir")
        self.BtnGuardaraExcel = QtWidgets.QPushButton(self.centralwidget)
        self.BtnGuardaraExcel.setGeometry(QtCore.QRect(90, 10, 111, 23))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.BtnGuardaraExcel.setFont(font)
        self.BtnGuardaraExcel.setObjectName("BtnGuardaraExcel")
        self.BtnGuardarenPDF = QtWidgets.QPushButton(self.centralwidget)
        self.BtnGuardarenPDF.setGeometry(QtCore.QRect(230, 40, 111, 23))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.BtnGuardarenPDF.setFont(font)
        self.BtnGuardarenPDF.setObjectName("BtnGuardarenPDF")
        self.BtnGuardarenWord = QtWidgets.QPushButton(self.centralwidget)
        self.BtnGuardarenWord.setGeometry(QtCore.QRect(90, 40, 111, 23))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.BtnGuardarenWord.setFont(font)
        self.BtnGuardarenWord.setObjectName("BtnGuardarenWord")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Codigo postales"))
        self.BtnAbrir.setText(_translate("MainWindow", "Abrir"))
        self.BtnGuardaraExcel.setText(_translate("MainWindow", "Guardar a Excel"))
        self.BtnGuardarenPDF.setText(_translate("MainWindow", "Guardar pdf"))
        self.BtnGuardarenWord.setText(_translate("MainWindow", "Guardar en word"))


class MiVentana(QMainWindow):
    d1 = []

    def __init__(self):
        super().__init__()

        # Configurar la interfaz gráfica generada por pyuic5
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.BtnAbrir.clicked.connect(self.abrir)
        self.ui.BtnGuardarenPDF.clicked.connect(self.guardar_en_pdf)
        self.ui.BtnGuardarenWord.clicked.connect(self.guardar_en_word)
        self.ui.BtnGuardaraExcel.clicked.connect(self.guardar_en_excel)
        self.fig = Figure(figsize=(5, 4))
        self.canvas = FigureCanvas(self.fig)
        self.scene = QGraphicsScene(self)
        self.scene.addWidget(self.canvas)

    def guardar_en_word(self):
        document = Document()

        for row in self.d1:
            table = document.add_table(rows=1, cols=len(row))
            table.autofit = False

            for i, cell in enumerate(row):
                table.cell(0, i).text = str(cell)

        document.save('datos.docx')

    def guardar_en_excel(self):
        workbook = Workbook()
        sheet = workbook.active

        for row_data in self.d1:
            sheet.append(row_data)

        workbook.save('datos.xlsx')

    def guardar_en_pdf(self):
        # Datos de ejemplo
        data = self.d1
        pdf_filename = 'tabla_datos.pdf'

        custom_size = (8 * 180, 80 * 20)  # Convertir pulgadas a puntos (1 pulgada = 72 puntos)
        pdf = SimpleDocTemplate(pdf_filename, pagesize=custom_size)
        # Crear la tabla
        table = Table(data)

        # Estilo de la tabla
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        # Aplicar el estilo a la tabla
        table.setStyle(style)

        # Crear la lista de elementos a agregar al PDF
        elements = [table]

        # Generar el PDF
        pdf.build(elements)

        print(f"El archivo '{pdf_filename}' ha sido creado.")

    def abrir(self):
        ruta_archivo, _ = QFileDialog.getOpenFileName(None, "Seleccionar archivo", "", "Archivos de texto (*.txt)")

        with open(ruta_archivo, 'r', encoding='cp1252') as archivo:
            columnas = archivo.readline().strip()
            columna = columnas.split('|')

            # Configurar las columnas del QTreeWidget
            self.ui.treeWidget.setColumnCount(len(columna))
            self.ui.treeWidget.setHeaderLabels(columna)

            # Agregar un elemento de nivel superior con los títulos de las columnas
            titulo_item = QTreeWidgetItem(self.ui.treeWidget, columna)
            self.ui.treeWidget.addTopLevelItem(titulo_item)
            for linea in archivo:
                datos = linea.strip().split('|')

                item = QTreeWidgetItem(self.ui.treeWidget, datos)
                self.d1.append(datos[:])  # Agregar una copia de los datos a d1
                self.ui.treeWidget.addTopLevelItem(item)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ventana = MiVentana()
    ventana.show()
    sys.exit(app.exec_())



