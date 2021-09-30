import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import *
from PyQt5.QtCore import *
import docx2txt


class RTE(QMainWindow):
    def __init__(self):
        super(RTE, self).__init__()
        """ O contato visual todo que teremos. A principal aba se dizendo """
        self.editor = QTextEdit()
        self.fontSizeBox = QSpinBox()

        font = QFont('Times', 12)
        self.editor.setFont(font)
        self.path = ""
        self.setCentralWidget(self.editor)
        self.setWindowTitle('Rich Text Editor')
        self.showMaximized()
        self.create_menu_bar()
        self.create_tool_bar()
        self.editor.setFontPointSize(12)

    def create_menu_bar(self):
        """ File menu """
        menuBar = QMenuBar(self)

        file_menu = QMenu("File", self)
        menuBar.addMenu(file_menu)

        save_action = QAction('Save', self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        open_action = QAction('Open', self)
        open_action.triggered.connect(self.file_open)
        file_menu.addAction(open_action)

        rename_action = QAction('Rename', self)
        rename_action.triggered.connect(self.file_saveas)
        file_menu.addAction(rename_action)

        pdf_action = QAction("Save as PDF", self)
        pdf_action.triggered.connect(self.save_pdf)
        file_menu.addAction(pdf_action)

        edit_menu = QMenu("Edit", self)
        menuBar.addMenu(edit_menu)

        """ Uma outra forma de copiar, além do atalho CRTL+C e do botão copy na barra de ferramentas"""
        copy_action = QAction('Copy', self)
        copy_action.triggered.connect(self.editor.copy)
        edit_menu.addAction(copy_action)

        """ Uma outra forma de colar, além do atalho CRTL+V e do botão paste na barra de ferramentas """
        paste_action = QAction('Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)

        clear_action = QAction('Clear', self)
        clear_action.triggered.connect(self.editor.clear)
        edit_menu.addAction(clear_action)

        select_action = QAction('Select All', self)
        select_action.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(select_action)

        view_menu = QMenu("View", self)
        menuBar.addMenu(view_menu)

        fullscr_action = QAction('Full Screen View', self)
        fullscr_action.triggered.connect(lambda: self.showFullScreen())
        view_menu.addAction(fullscr_action)

        normalscr_action = QAction('Normal View', self)
        normalscr_action.triggered.connect(lambda: self.showNormal())
        view_menu.addAction(normalscr_action)

        minscr_action = QAction('Minimize', self)
        minscr_action.triggered.connect(lambda: self.showMinimized())
        view_menu.addAction(minscr_action)

        self.setMenuBar(menuBar)

    def create_tool_bar(self):
        """Barra de feramentas. Onde fica os botôes para agilizar o procedimento... """
        toolbar = QToolBar()

        save_action = QAction(QIcon('save.png'), 'Save', self)
        save_action.triggered.connect(self.saveFile)
        toolbar.addAction(save_action)

        toolbar.addSeparator()

        undoButton = QAction(QIcon('undo.png'), 'Undo', self)
        """ Dizer o que acontece se o botao for clicado """
        undoButton.triggered.connect(self.editor.undo)
        """ Botão de desfazer a barra de ferramentas """
        toolbar.addAction(undoButton)

        redoButton = QAction(QIcon('redo.png'), 'Redo', self)
        redoButton.triggered.connect(self.editor.redo)
        toolbar.addAction(redoButton)

        copyButton = QAction(QIcon('copy.png'), 'Copy', self)
        copyButton.triggered.connect(self.editor.copy)
        toolbar.addAction(copyButton)

        cutButton = QAction(QIcon('cut.png'), 'Cut', self)
        cutButton.triggered.connect(self.editor.cut)
        toolbar.addAction(cutButton)

        pasteButton = QAction(QIcon('paste.png'), 'Paste', self)
        pasteButton.triggered.connect(self.editor.paste)
        toolbar.addAction(pasteButton)

        toolbar.addSeparator()

        """ Variaçoes da fonte """
        self.fontBox = QComboBox(self)
        self.fontBox.addItems(
            ["Algerian", "Arial Black", "Arial", "Calibri", "Corbel", "Elephant", "GEORGIA", "Helvetica", "Impact",
             "Papyrus", "Times New Roman" ])
        self.fontBox.activated.connect(self.setFont)
        toolbar.addWidget(self.fontBox)

        self.fontSizeBox.setValue(12)
        self.fontSizeBox.valueChanged.connect(self.setFontSize)
        toolbar.addWidget(self.fontSizeBox)

        toolbar.addSeparator()

        leftAllign = QAction(QIcon('left-align.png'), 'left Allign', self)
        leftAllign.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignLeft))
        toolbar.addAction(leftAllign)

        centerAllign = QAction(QIcon('center-align.png'), 'Center Allign', self)
        centerAllign.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignCenter))
        toolbar.addAction(centerAllign)

        rightAllign = QAction(QIcon('right-align.png'), 'Right Allign', self)
        rightAllign.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignRight))
        toolbar.addAction(rightAllign)

        toolbar.addSeparator()

        boldBtn = QAction(QIcon('bold.png'), 'Bold', self)
        boldBtn.triggered.connect(self.boldText)
        toolbar.addAction(boldBtn)

        underlineBtn = QAction(QIcon('underline.png'), 'underline', self)
        underlineBtn.triggered.connect(self.underlineText)
        toolbar.addAction(underlineBtn)

        italicBtn = QAction(QIcon('italic.png'), 'italic', self)
        italicBtn.triggered.connect(self.italicText)
        toolbar.addAction(italicBtn)

        toolbar.addSeparator()

        zoomInAction = QAction(QIcon('zoom-in.png'),'Zoom in', self)
        zoomInAction.triggered.connect(self.editor.zoomIn)
        toolbar.addAction(zoomInAction)

        zoomOutAction = QAction(QIcon('zoom-out.png'), 'Zoom out',self)
        zoomOutAction.triggered.connect(self.editor.zoomOut)
        toolbar.addAction(zoomOutAction)


        self.addToolBar(toolbar)

    def setFontSize(self):
        """ O local para escolher o tamanho da fonte """
        value = self.fontSizeBox.value()
        self.editor.setFontPointSize(value)

    def setFont(self):
        """ O local para escolher os diferentes tipos de fonte """
        font = self.fontBox.currentText()
        self.editor.setCurrentFont(QFont(font))

    def italicText(self):
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not (state))

    def underlineText(self):
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not (state))

    def boldText(self):
        if self.editor.fontWeight != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)

    def saveFile(self):
        """ Botão de Salvar. Sem ser o salvar pelo 'File' no menu_bar. Salva apenas no formato .Txt """
        print(self.path)
        if self.path == '':
            self.file_saveas()
        text = self.editor.toPlainText()
        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_open(self):
        """ Abrir pelo 'File' no menu_bar """
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "","Text documents (*.txt)")

        try:
            # with open(self.path, 'r') as f:
            #    text = f.read()
            text = docx2txt.process(self.path)  # docx2txt
            # doc = Document(self.path)         # if using docx
            # text = ''
            # for line in doc.paragraphs:
            #    text += line.text
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        """ Salvar pelo 'File' no menu_bar. Além dessa opção, tem a opçao de salvar pelo botão  """
        print(self.path)
        if self.path == '':
            # If we do not have a path, we need to use Save As.
            self.file_saveas()

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "","Text documents (*.txt)")
        if self.path == '':
            return
        text = self.editor.toPlainText()
        try:
            with open(path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)

    def save_pdf(self):
        """ Uma opçao a parte pra salvar em PDF além de poder salvar em Txt """
        f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "Adobe Acrobat Document (*.pdf)")
        print(f_name)

        if f_name != '':  # if name not empty
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(f_name)
            self.editor.document().print_(printer)


app = QApplication(sys.argv)
window = RTE()
window.show()
sys.exit(app.exec_())

input()