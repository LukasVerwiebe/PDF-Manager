# PDF-Manager
Datei-Manager für PDF-Dateien mit dem Dateien aneinandergefügt, ausgelesen und aufgesplittet werden können.
Das Programm wurde mit Python erstellt, dabei wurden die folgenden Bibliotheken verwendet:
- PyQt5
- PyPDF2

## Funktionen:
### Drag & Drop
- Zum Speichern einer neuen PDF-Datei kann neben dem Auswählen eines Speicherorts auch eine Datei per Drag & Drop in das Feld gezogen werden dort wird dann der Dateipfad angezeigt.
- Die PDF-Dateien welche z.B. zusammengefügt werden sollen müssen per Drag & Drop in das Listen Feld in der Mitte der Applikation gezogen werden, die Dateien werden dort dann angezeigt zusammen mit dem Dateipfad die Reihenfolge der Dateien kann ebenfalls per Drag & Drop verändert werden. Die Dateien werden nach ihrer Reiehenfolge in der Liste eingelesen und verarbeitet.
- Mittels des Button 'Löschen' kann ein einzelner Eintrag aus der Liste entfernt werden.
- Mittels des Button 'Zurücksetzten' werden alle Einträge aus der Liste und dem Speicherfeld entfernt.

 ### Zusammenfügen
 - Die eingefügten Dateien könen über den Button 'Merge' zusammengefügt werden, dieses werden dabei nach iherer Reihenfolge zusammengefügt werden.

### Auslesen
- Die erste Datei in der Liste kann mittels des Button 'Auslesen' in eine Textdatei ausgelesen werden. Die Seiten werden durch einen Trenner in der Datei angezeigt.

### Aufsplitten
- Die erste Datei in der Liste kann mittels des Button 'Aufsplitten' in einzelne Dateien aufgeteilt werden. Dabei wird für jede Seite eine eigene PDF-Datei erstellt. 

## Oberfläche:
![image](https://github.com/LukasVerwiebe/PDF-Manager/assets/63674539/e48d2b6a-9fd0-4e4d-9d81-642fd42108f6)
