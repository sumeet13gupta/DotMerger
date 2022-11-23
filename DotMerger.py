import sys, os
from PyQt5.QtWidgets import QApplication, QMainWindow, QListWidget, QListWidgetItem, QPushButton, QInputDialog, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QUrl
from PyPDF4 import PdfFileMerger
import comtypes.client
from comtypes.gen import PowerPoint
from comtypes.client import Constants, CreateObject

class ListBoxWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.resize(600, 600)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            links = []
            for url in event.mimeData().urls():
                # https://doc.qt.io/qt-5/qurl.html
                if url.isLocalFile():
                    links.append(str(url.toLocalFile()))
                else:
                    links.append(str(url.toString()))
            self.addItems(links)
        else:
            event.ignore()

class AppDemo(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1200, 600)
        self.setWindowTitle("DotMerger - merge pdfs and ppts as pdfs")
        self.setWindowIcon(QIcon('wep.png'))
        self.listbox_view = ListBoxWidget(self)

        self.btn = QPushButton('Generate pdf', self)
        self.btn.setGeometry(850, 400, 200, 50)
        self.btn.clicked.connect( self.OnClicked)                                           #lambda: print(self.getSelectedItem()))

        
    def OnClicked(self):
        print("clicked")
        print(self.getSelectedItem())
        list_of_files = self.getAllItem()
        op = self.takeopath()
        nm = self.takeName()
        self.simple_merger(list_of_files, op, nm)

    def getSelectedItem(self):
        item = QListWidgetItem(self.listbox_view.currentItem())
        return item.text()

    def getAllItem(self):
        all_file_list = [self.listbox_view.item(i).text() for i in range(self.listbox_view.count())]
        
        print(all_file_list)
        return all_file_list
    #Added func
    def takeName(self):
        Name, done1 = QInputDialog.getText(self, 'Input Dialog', 'Enter Output file name:')
        return Name
    def takeopath(self):
        dialog = QFileDialog()
        path = dialog.getExistingDirectory(self, 'Select an directory')
        # path, done1 = QInputDialog.getText(self, 'Input Dialog', 'Enter Output path:')
        return path
    # Merger
    def simple_merger(self, pdf_list: list, output_path: str, output_name: str):
        merger = PdfFileMerger(strict=False)
        for pdf_file in pdf_list:
            if os.path.splitext(pdf_file)[1] == '.pptx' or os.path.splitext(pdf_file)[1] == '.ppt' :
                opname = os.path.splitext(pdf_file)[0] +".pdf"
                print(opname)
                opname = self.PPTtoPDF(pdf_file,opname,32)
                merger.append(opname)
            else:
                merger.append(pdf_file)

        myFile = output_path +"/"+ output_name +".pdf"
        print(myFile)
        merger.write(myFile)
        merger.close()
    # converter

    def PPTtoPDF(self, inputFileName, outputFileName, formatType = 32):
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        constant = comtypes.client.Constants(powerpoint)

        if outputFileName[-3:] != 'pdf':
            outputFileName = outputFileName + ".pdf"
        print("Can't open : ", inputFileName)
        deck = powerpoint.Presentations.Open(inputFileName)
        deck.SaveAs(outputFileName, PowerPoint.ppSaveAsPDF) # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()   
        return outputFileName


if __name__ == '__main__':
    app = QApplication(sys.argv)

    demo = AppDemo()
    demo.show()

    sys.exit(app.exec_())

