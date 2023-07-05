import os
import subprocess
import os.path
import PyPDF2
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QListWidget, QPushButton, QFileDialog
from PyQt5.QtCore import Qt, QEvent, QObject
from PyQt5.QtGui import QPalette, QColor, QIcon
from PyPDF2 import PdfFileReader, PdfMerger
from docx import Document
from docx.shared import Inches
import zipfile
from pdf2docx import Converter
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QDialog, QFileDialog, QHBoxLayout,
                             QLabel, QListWidget, QListWidgetItem, QMainWindow,
                             QMessageBox, QPushButton, QSplitter, QStyle,
                             QToolBar, QToolButton, QVBoxLayout, QWidget,
                             QAbstractItemView)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.archivos_seleccionados = []

        self.layout = QVBoxLayout()

        self.label = QLabel("Arrastra y suelta archivos PDF aquí:")
        self.layout.addWidget(self.label)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        self.list_widget.setStyleSheet(
            "QListWidget { background-color: black; color: white; }")
        self.list_widget.setDragDropMode(
            QAbstractItemView.InternalMove)  # agregar esta linea
        self.layout.addWidget(self.list_widget)

        self.boton_combinar = QPushButton("Combinar")
        self.boton_combinar.setStyleSheet(
            "background-color: green; color: white;")
        self.boton_combinar.clicked.connect(self.combinar_archivos)
        self.layout.addWidget(self.boton_combinar)

        self.label_contador = QLabel("Archivos cargados: 0")
        self.label_contador.setStyleSheet("color: white;")
        self.layout.addWidget(self.label_contador)

        self.setLayout(self.layout)

        self.setAcceptDrops(True)

        self.set_estilo_ventana()

        # Instala un filtro de eventos en la lista de archivos
        self.list_widget.installEventFilter(self)
        self.boton_word = QPushButton("Convertir a Word")
        self.boton_word.setStyleSheet(
            "background-color: blue; color: white;")
        self.boton_word.clicked.connect(self.convert_to_word)
        self.layout.addWidget(self.boton_word)

        self.layout.addWidget(self.boton_combinar)
        self.layout.addWidget(self.boton_word)
        self.layout.addWidget(self.label_contador)
        self.setLayout(self.layout)

    def convert_to_word(self):
        if len(self.archivos_seleccionados) == 0:
            return

        # Obtener el primer archivo PDF seleccionado
        archivo_pdf = self.archivos_seleccionados[0]

        # Generar el nombre del archivo de Word de salida
        nombre, _ = os.path.splitext(archivo_pdf)
        nuevo_archivo = nombre + ".docx"

        # Crear un objeto Converter
        convertidor = Converter(archivo_pdf)

        # Convertir el PDF a Word
        convertidor.convert(nuevo_archivo, start=0, end=None)

        # Cerrar el convertidor
        convertidor.close()

        print("Archivo convertido exitosamente.")

    def dragMoveEvent(self, e):  # agregar esta funcion
        if e.mimeData().hasUrls():
            e.setDropAction(Qt.MoveAction)
            e.accept()
            self.list_widget.setCurrentIndex(self.list_widget.indexAt(e.pos()))
        else:
            e.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        archivos

    def set_estilo_ventana(self):
        # Cambiar el fondo de la ventana a negro
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(0, 0, 0))
        self.setPalette(palette)

        # Cambiar el color de las letras a blanco
        self.setStyleSheet("QLabel { color: white; }")

        # Cambiar el color de los bordes a verde
        self.setStyleSheet("QWidget { border: 2px solid green; }")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        archivos = [url.toLocalFile() for url in urls if url.isLocalFile()]
        self.archivos_seleccionados.extend(archivos)
        self.actualizar_list_widget()

    def actualizar_list_widget(self):
        self.list_widget.clear()
        for archivo in self.archivos_seleccionados:
            nombre_archivo = os.path.basename(archivo)
            self.list_widget.addItem(nombre_archivo)

        self.actualizar_contador()

    def actualizar_contador(self):
        contador = len(self.archivos_seleccionados)
        self.label_contador.setText(f"Archivos cargados: {contador}")

    def combinar_archivos(self):
        if len(self.archivos_seleccionados) == 0:
            return

        merger = PdfMerger()

        # Append files to the merger
        for archivo in self.archivos_seleccionados:
            merger.append(archivo)

        ruta_combinado, _ = QFileDialog.getSaveFileName(
            self, "Guardar archivo combinado", "", "Archivos PDF (*.pdf)")

        # Save the merged file
        if ruta_combinado:
            merger.write(ruta_combinado)
            merger.close()
            print("Archivos combinados exitosamente.")

    def limpiar_archivos(self):
        self.archivos_seleccionados = []
        self.actualizar_list_widget()

    def eventFilter(self, obj, event):
        if obj is self.list_widget and event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Delete or event.key() == Qt.Key_Backspace:
                self.eliminar_archivos_seleccionados()
                return True
        elif obj is self.list_widget and event.type() == QEvent.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                index = self.list_widget.indexAt(event.pos())
                if index.isValid():
                    self.list_widget.startDrag(Qt.MoveAction)
                    return True
        elif obj is self.list_widget and event.type() == QEvent.DragEnter:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True
        elif obj is self.list_widget and event.type() == QEvent.DragMove:
            event.acceptProposedAction()
            return True
        elif obj is self.list_widget and event.type() == QEvent.Drop:
            urls = event.mimeData().urls()
            archivos = [url.toLocalFile() for url in urls if url.isLocalFile()]
            self.archivos_seleccionados.extend(archivos)
            self.actualizar_list_widget()
            event.acceptProposedAction()
            return True
        return super().eventFilter(obj, event)

    def eliminar_archivos_seleccionados(self):
        indices = [index.row() for index in self.list_widget.selectedIndexes()]
        indices.sort(reverse=True)
        for indice in indices:
            del self.archivos_seleccionados[indice]
        self.actualizar_list_widget()


if __name__ == "__main__":
    import webbrowser

    app = QApplication([])
    window = MainWindow()
    window.setWindowTitle("PDF Uni Pro")
    # Ajusta la geometría de la ventana (x, y, ancho, alto)
    window.setGeometry(100, 100, 400, 800)
    window.show()

    def abrir_link_donaciones():
        webbrowser.open("https://paypal.me/luismellizo")

    boton_donaciones = QPushButton("Donar", window)
    boton_donaciones.setStyleSheet(
        "QPushButton { background-color: yellow; color: black; font-size: 12px; }")
    boton_donaciones.setFixedSize(60, 20)
    boton_donaciones.move(10, window.height() - boton_donaciones.height() - 10)
    boton_donaciones.clicked.connect(abrir_link_donaciones)
    window.layout.addWidget(boton_donaciones)

    app.exec_()
