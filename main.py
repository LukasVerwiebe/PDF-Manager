########################################################
# Python script mit der Funktion PDF Dateien zu Mergen
# Autor: Lukas Verwiebe
########################################################

# Importe für das script
import sys
import os
import io
import aspose.words as aw
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QListWidget, \
                            QVBoxLayout, QHBoxLayout, QGridLayout, QDialog, QFileDialog, QMessageBox, \
                            QAbstractItemView
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon
from PyPDF2 import PdfFileMerger, PdfMerger, PdfReader, PdfWriter


# Klasse für die Funktionen der PDF-Liste
class ListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent=None)
        self.setAcceptDrops(True)
        self.setStyleSheet('font-size: 25px;')
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    # Drag-and-Drop Funktion für das PDF-Listen Feld
    # Dateien können in die Felder gezogen werden, dabei wird dann der Dateipfad dort eingetragen
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            return super().dragEnterEvent(event)

    # Bewegen der Dateien innerhalb der PDF-Liste für die Festlegung der Reihenfolge
    # in welcher die Dateien in der neuen PDF-Datei angefügt werden sollen
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            return super().dragMoveEvent(event)

    # Drop Event die Dateinamen werden in einem Array festgehalten
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            pdffiles = []

            for url in event.mimeData().urls():
                if url.isLocalFile():
                    # Nur Dateien mit der Dateiendung .pdf werden in die Liste eingelesen
                    if url.toString().endswith('.pdf'):
                        pdffiles.append(str(url.toLocalFile()))
            self.addItems(pdffiles)
        else:
            return super().dropEvent(event)


# Klasse für die Funktionen des Feldes für die neu Erstellte PDF Datei
class OutputField(QLineEdit):
    def __init__(self):
        super().__init__()
        # Höhe des Feldes
        self.height = 55
        # Schriftgöße des Feldes
        self.setStyleSheet('font-size: 20px')
        self.setFixedHeight(self.height)

    # Drag-and-Drop Funktion für das Feld
    # eine Datei kann in das Feld gezogen werden, dabei wird dann der Dateipfad dort eingetragen
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    # Drag-and-Drop Funktion für das Feld
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            if event.mimeData().urls():
                self.setText(event.mimeData().urls()[0].toLocalFile())
        else:
            event.ignore()


# Klasse für die PushButton der Anwendung
class Button(QPushButton):
    def __init__(self, label_text):
        super().__init__()
        self.setText(label_text)
        # Festlegung des Stils der Button
        self.setStyleSheet('''
            font-size: 20px;
            width: 180px;
            height: 50px;
        ''')


# Klasse für das Ausführen der Anwendung und Aufbau des PyQt Fensters
class PDFApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PDF Datei-Manager')
        self.resize(1200, 600)

        # Initializer der Button Instanzen
        self.buttonSplitSelect = None
        self.outputFile = None
        self.buttonBrowseOutputFile = None
        self.buttonReadSelect = None
        self.buttonReset = None
        self.buttonClose = None
        self.buttonMerge = None
        self.buttonDeleteSelect = None
        self.pdfListWidget = None
        # Aufruf des User-Interface
        self.initUI()

        # Button: zum Speichern der neuen PDF-Datei
        self.buttonBrowseOutputFile.clicked.connect(self.populateFileName)

    # Zusammenstellung der UI
    def initUI(self):
        # Die Boxen für das Layout der Seite
        mainLayout = QVBoxLayout()
        outputFolderRow = QHBoxLayout()
        buttonLayout = QHBoxLayout()

        # Das Feld für die Ausgabe Datei
        self.outputFile = OutputField()
        outputFolderRow.addWidget(self.outputFile)

        # Button zum Öffnen des Datei-Explorers für den Speicher Vorgang
        self.buttonBrowseOutputFile = Button('&Speichern unter')
        self.buttonBrowseOutputFile.setFixedHeight(self.outputFile.height)
        outputFolderRow.addWidget(self.buttonBrowseOutputFile)

        # List Widget für das Einfügen der PDF-Dateien
        self.pdfListWidget = ListWidget(self)

        # Button zum Aufsplitten eines Dokuments aus der List Widget
        self.buttonSplitSelect = Button('&Aufsplitten')
        self.buttonSplitSelect.clicked.connect(self.splitPDF)
        buttonLayout.addWidget(self.buttonSplitSelect)

        # Button zum Auslesen eines Eintrags aus der List Widget
        self.buttonReadSelect = Button('&Auslesen')
        self.buttonReadSelect.clicked.connect(self.extractText)
        buttonLayout.addWidget(self.buttonReadSelect)

        # Button zum Löschen eines Eintrags aus der List Widget
        self.buttonDeleteSelect = Button('&Löschen')
        self.buttonDeleteSelect.clicked.connect(self.deleteSelected)
        buttonLayout.addWidget(self.buttonDeleteSelect, 1, Qt.AlignRight)

        # Button für den PDF Merge Vorgang
        self.buttonMerge = Button('&Merge')
        self.buttonMerge.clicked.connect(self.mergeFile)
        buttonLayout.addWidget(self.buttonMerge)

        # Button zum Schließen der Anwendung
        self.buttonClose = Button('&Schließen')
        self.buttonClose.clicked.connect(QApplication.quit)
        buttonLayout.addWidget(self.buttonClose)

        # Button zum zurücksetzten aller bisherigen Einträge
        self.buttonReset = Button('&Zurücksetzten')
        self.buttonReset.clicked.connect(self.clearQueue)
        buttonLayout.addWidget(self.buttonReset)

        # Hinzufügen der Layout-Boxen
        mainLayout.addLayout(outputFolderRow)
        mainLayout.addWidget(self.pdfListWidget)
        mainLayout.addLayout(buttonLayout)

        self.setLayout(mainLayout)

    # Funktion zum Löschen eines Eintrags aus der List Widget
    def deleteSelected(self):
        #
        for item in self.pdfListWidget.selectedItems():
            self.pdfListWidget.takeItem(self.pdfListWidget.row(item))

    # Funktion zum Leeren aller Felder
    def clearQueue(self):
        self.pdfListWidget.clear()
        self.outputFile.setText('')

    # Funktion zum Anzeigen von Nachrichten
    def dialogMessage(self, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle('PDF Manager')
        dlg.setIcon(QMessageBox.Information)
        dlg.setText(message)
        dlg.show()

    # Funktion zum Festlegen des Speicherpfads der Datei als Standard wird das Format pdf eingestellt
    # andere Einstellungen werden dadurch unterbunden
    def _getSaveFilePath(self):
        file_save_path, _ = QFileDialog.getSaveFileName(self, 'Save PDF file', os.getcwd(), 'PDF file (*.pdf')
        return file_save_path

    # Eintragen des Dateipfads für die neue PDF-Datei in das vorgesehene Feld auf dem Layout
    def populateFileName(self):
        path = self._getSaveFilePath()
        if path:
            self.outputFile.setText(path)

    # Funktion für den PDF Merge Vorgang
    def mergeFile(self):
        # Erst prüfen, ob eine Ziel-Datei bereits angegeben wurde,
        # wenn nicht, wird zuerst der Datei-Explorer geöffnet, um diese festzulegen
        if not self.outputFile.text():
            self.populateFileName()
            return

        # Wenn Dateien zum Mergen angegeben wurden, wird der Vorgang gestartet
        if self.pdfListWidget.count() > 0:
            pdfMerger = PdfMerger()

            try:
                # Die Liste mit den Dateien wird durchgegangen und dem pdfmerger hinzugefügt
                for i in range(self.pdfListWidget.count()):
                    pdfMerger.append(self.pdfListWidget.item(i).text())

                # Mit den hinzugefügten Dateien wird der Schreib Vorgang für die neue PDF-Datei gestartet
                pdfMerger.write(self.outputFile.text())
                pdfMerger.close()

                # Nach dem Abschluss des Vorgangs wird die PDF-Liste geleert
                self.pdfListWidget.clear()
                # Zum Abschluss wird eine Meldung angezeigt, dass der Vorgang erfolgreich war
                self.dialogMessage('Der PDF Merge ist Vollständig')
            except Exception as e:
                self.dialogMessage(e)
        # Wenn keine Dateien für den Merge Vorgang angegeben sind, wird eine Meldung angezeigt
        else:
            self.dialogMessage('Die Warteschlange ist Leer!')

    # Funktion die aus einer PDF-Datei den Text ausließt. Der Text wird in eine TXT Datei geschrieben
    def extractText(self):
        # Lesen aus der PDF-Datei, die zuerst in der PDF-Liste vorkommt.
        # Für die nächste muss der Vorgang neu gestartet werden.
        if self.pdfListWidget.count() > 0:
            # Der PDF-Reader wird mit der ersten datei aus der PDF-Liste initialisiert
            pdfReader = PdfReader(self.pdfListWidget.item(0).text())

            # Die Datei out.txt erstellen, wenn nicht vorhanden und öffnen
            with open('out.txt', 'w') as file:
                # Die PDF-Datei Seitenweise durchgehen und auslesen
                for page_num in range(len(pdfReader.pages)):
                    pageObj = pdfReader.pages[page_num]

                    try:
                        # Den Text aus der Seite extrahieren
                        txt = pageObj.extract_text()
                        file.write('\n')
                        file.write(''.center(100, '-'))
                        file.write('\n')

                        # Mit der nächsten Seite weitermachen
                        file.write('Seite: {0}\n'.format(page_num + 1))
                        #file.write('\n')
                        file.write(''.center(100, '-'))
                        file.write('\n')
                        file.write(txt)
                    except Exception as e:
                        self.dialogMessage(e)

                # Schließen des Schreibvorgangs
                file.close()

            # Erfolgsmeldung:
            self.dialogMessage('Das Auslesen der ersten Datei ist Vollständig.')

    # Funktion zum Aufsplitten einer PDF-Datei. Die Datei wird anhand der Seiten aufgeteilt.
    def splitPDF(self):
        # Prüfen, ob eine Datei eingefügt wurde
        if self.pdfListWidget.count() > 0:
            # Die erste Datei aus der Liste einlesen als file
            with open(self.pdfListWidget.item(0).text(), 'rb') as file:
                # Die Datei in den Reader laden
                pdfReader = PdfReader(file)

                total_pages = len(pdfReader.pages)
                # Die Datei wird in einzelne neue PDF Dateien aufgeteilt.
                # Für jede Seite wird eine eigene PDF-Datei erstellt.
                for index, page in enumerate(pdfReader.pages):
                    pdfWriter = PdfWriter()
                    pdfWriter.add_page(page)

                    # Die neuen PDF-Dateien werden nummeriert
                    with open('page_{0}.pdf'.format(index + 1), 'wb') as output:
                        pdfWriter.write(output)

            # Erfolgsmeldung:
            self.dialogMessage('Das Aufsplitten der Datei wahr erfolgreich.')


# Instanz der Applikation
app = QApplication(sys.argv)
# Stil der Applikation: Fusion
app.setStyle('fusion')

# Aufruf der Anwendung mittels der Klasse: PDFApp
pdfApp = PDFApp()
pdfApp.show()

# Exit Prozess
sys.exit(app.exec_())

