from PySide6.QtCore import Qt, QMimeData, QByteArray, Signal
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QMessageBox, QLabel, QPushButton, QWidget, QApplication, QMainWindow, QTextEdit, QVBoxLayout, \
    QHBoxLayout, QFileDialog, QInputDialog
import threading, re, time, tempfile, pyperclip
import os
from PySide6.QtGui import QGuiApplication

try:
    import win32com.client as win32
    import win32gui
    import pythoncom
    win32comsupport = True
    from pynput import keyboard
except Exception as err:
    win32comsupport = False
    print(f'Win32com not supported: {err} \nWordHook disabled.')

ver = " Bleeding Edge 1.3.2"

class DraggableWidget(QWidget):
    def __init__(self, parent=None):
        super(DraggableWidget, self).__init__(parent)
        self.moving = False
        self.offset = None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.moving = True
            self.offset = event.position().toPoint()

    def mouseMoveEvent(self, event):
        if self.moving:
            self.move((event.globalPosition().toPoint() - self.offset))

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.moving = False
class MainWindow(QMainWindow):
    update_equation_edit_signal = Signal(str)
    copy_svg_thread_safe_signal = Signal(str)
    thread_safe_svg_paste_signal = Signal(str)
    doneWidgetAutoShowSignal = Signal(bool)

    def closeEvent(self, event):
        self.doneWidget.close()
        super().closeEvent(event)

    def infoDialog(self, message):
        dialog = QMessageBox()
        dialog.setWindowTitle("ⓘ Info")
        dialog.setText(message)
        dialog.exec()

    def toggleShow(self, arg):
        if arg:
            self.doneWidget.setWindowOpacity(1)
        else:
            self.doneWidget.setWindowOpacity(0)

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setWindowTitle(f"MathJax To Go - {ver}")
        self.update_equation_edit_signal.connect(self.update_equation_edit)
        self.clipboard = QApplication.clipboard()
        self.doneWidgetAutoShowSignal.connect(self.toggleShow)

        # Properties and equation init
        self.svgData = ""
        self.equation = r"\Large \text{you gonna type something or what?}"
        self.autoCopy = False
        self.physicsEnabled = False
        self.colorsv2Enabled = False
        self.displayStyle = True
        self.equation_edit = QTextEdit()
        self.equation_edit.setPlaceholderText("Type Equation Here")
        self.equation_edit.setAcceptRichText(False)
        self.customCDN = False
        self.wordHookStatus = False
        self.replaceFlag = False
        self.copy_svg_thread_safe_signal.connect(self.copySvg)
        self.thread_safe_svg_paste_signal.connect(self.experimentalSvgFileInsertion)

        # Create layout
        self.layout = QVBoxLayout()
        self.topLayout = QHBoxLayout()
        self.interactiveWindowLayout = QHBoxLayout()
        self.interactiveWindowLayout.addWidget(self.equation_edit)
        self.optionInsertionLayout = QVBoxLayout()
        self.interactiveWindowLayout.addLayout(self.optionInsertionLayout)
        self.optionLowerLayout = QHBoxLayout()

        # Load webengine
        self.mathjax_script = r'<script type="text/javascript" async src = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg-full.js"> </script>'
        self.view = QWebEngineView()
        self.load_mathjax()
        self.view.loadFinished.connect(self.update_mathjax)
        self.interactiveWindowLayout.addWidget(self.view)

        # Wordhook Preliminaries
        try:
            if win32comsupport:
                self.wordHookButton = QPushButton("Hook to MS Word")
                self.wordHookButton.setToolTip('Hook to MS Word for real-time editing'
                                               '\nType TeX equations between $$ to display a real-time widget, then'
                                               'press Enter to remove the widget and replace the text with the rendered SVG.')
                self.wordHookButton.setStyleSheet("background-color: darkred")
                self.wordHookButton.clicked.connect(self.wordHook)
                self.wordHookPlaceButton = QPushButton("Done")
                self.wordHookPlaceButton.setStyleSheet("background-color: #222288")
                self.wordHookPlaceButton.clicked.connect(lambda: setattr(self, 'replaceFlag', True))
                self.topLayout.addWidget(self.wordHookPlaceButton)
                self.wordHookPlaceButton.hide()
                self.topLayout.addWidget(self.wordHookButton)
                self.doneWidget = DraggableWidget()
                self.doneWidget.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
                screen = QGuiApplication.primaryScreen().geometry()
                x = 0
                y = (screen.height() / 2) - 150
                self.doneWidget.move(x, y)
                self.doneWidgetButton = QPushButton("Done", self.doneWidget)
                self.doneWidgetButton.setStyleSheet("background-color: #222288")
                self.doneWidgetButton.clicked.connect(lambda: setattr(self, 'replaceFlag', True))
                self.doneWidgetCloseButton = QPushButton("Close", self.doneWidget)
                self.doneWidgetCloseButton.setStyleSheet("background-color: darkred")
                self.doneWidgetCloseButton.clicked.connect(self.stop_word_hook)
                self.doneWidgetControlHelpButton = QPushButton("ⓘ Help") # ⓘ
                self.doneWidgetControlHelpButton.setStyleSheet("background-color: darkgreen")
                self.doneWidgetControlHelpButton.clicked.connect(lambda: self.infoDialog("- Click and drag in any empty space to move the widget."
                                                                                         "\n- You can use Ctrl+Scroll to zoom in/out in the SVG view."
                                                                                         "\n- Click the Done button to replace the typed equation with the rendered SVG."
                                                                                         "\n - You can also type \\done anywhere within the equation to trigger the replacement."
                                                                                         "\n- Click the Close button to exit WordHook"
                                                                                         "\n- Use the provided window size controls to adjust the view to your preference."
                                                                                         "\n- Click the Set Default button to save the current window size and position as default."
                                                                                         "\n- This way, WordHook will always start at the same position and size you saved."))
                self.smallView = QWebEngineView()
                def smallViewSizeChange(dir, val):
                    if dir == 'x':
                        self.smallView.setFixedSize(self.smallView.width() + val, self.smallView.height())
                        self.doneWidget.setFixedSize(self.doneWidget.width() + val, self.doneWidget.height())
                    elif dir == 'y':
                        self.smallView.setFixedSize(self.smallView.width(), self.smallView.height() + val)
                        self.doneWidget.setFixedSize(self.doneWidget.width(), self.doneWidget.height() + val)
                self.doneWidgetSizeUpButton = QPushButton("⇧", self.doneWidget)
                self.doneWidgetSizeUpButton.setStyleSheet("background-color: darkgray")
                self.doneWidgetSizeUpButton.setMaximumWidth(40)
                self.doneWidgetSizeUpButton.clicked.connect(lambda: smallViewSizeChange('y', -20))
                self.doneWidgetSizeDownButton = QPushButton("⇩", self.doneWidget)
                self.doneWidgetSizeDownButton.setStyleSheet("background-color: darkgray")
                self.doneWidgetSizeDownButton.setMaximumWidth(40)
                self.doneWidgetSizeDownButton.clicked.connect(lambda: smallViewSizeChange('y', 20))
                self.doneWidgetSizeLeftButton = QPushButton("⇦", self.doneWidget)
                self.doneWidgetSizeLeftButton.setStyleSheet("background-color: darkgray")
                self.doneWidgetSizeLeftButton.setMaximumWidth(40)
                self.doneWidgetSizeLeftButton.clicked.connect(lambda: smallViewSizeChange('x', -20))
                self.doneWidgetSizeRightButton = QPushButton("⇨", self.doneWidget)
                self.doneWidgetSizeRightButton.setStyleSheet("background-color: darkgray")
                self.doneWidgetSizeRightButton.setMaximumWidth(40)
                self.doneWidgetSizeRightButton.clicked.connect(lambda: smallViewSizeChange('x', 20))
                self.doneWidgetSizeLabel = QLabel("Click and drag here to move view")
                def doneWidgetSetDefault():
                    with open('./MJ2GSavedValues.ini', 'w') as f:
                        f.write(f'!doneWidgetWidth:{self.doneWidget.width()}')
                        f.write(f'\n!doneWidgetHeight:{self.doneWidget.height()}')
                        f.write(f'\n!doneWidgetX:{self.doneWidget.x()}')
                        f.write(f'\n!doneWidgetY:{self.doneWidget.y()}')
                try:
                    with open('./MJ2GSavedValues.ini', 'r') as f:
                        for line in f:
                            if line.startswith('!doneWidgetWidth:'):
                                self.doneWidget.setFixedWidth(int(line.split(':')[1]))
                            elif line.startswith('!doneWidgetHeight:'):
                                self.doneWidget.setFixedHeight(int(line.split(':')[1]))
                            elif line.startswith('!doneWidgetX:'):
                                self.doneWidget.move(int(line.split(':')[1]), self.doneWidget.y())
                            elif line.startswith('!doneWidgetY:'):
                                self.doneWidget.move(self.doneWidget.x(), int(line.split(':')[1]))
                except FileNotFoundError:
                    pass
                except Exception as e:
                    print(f'Error loading saved values: {e}')
                self.doneWidgetSetDefaultButton = QPushButton("Set Default", self.doneWidget)
                self.doneWidgetSetDefaultButton.setStyleSheet("background-color: #222288")
                self.doneWidgetSetDefaultButton.clicked.connect(doneWidgetSetDefault)
                def toggleWidgetAutoShow():
                    self.doneWidgetAutoShow = not self.doneWidgetAutoShow
                    if self.doneWidgetAutoShow:
                        self.doneWidget.setWindowOpacity(0)
                    self.doneWidgetAutoShowButton.setStyleSheet("background-color: darkgreen" if self.doneWidgetAutoShow else "background-color: darkred")
                self.doneWidgetAutoShowButton = QPushButton("Auto-Show", self.doneWidget)
                self.doneWidgetAutoShowButton.setStyleSheet("background-color: darkred")
                self.doneWidgetAutoShow = False
                self.doneWidgetAutoShowButton.clicked.connect(toggleWidgetAutoShow)
                doneWidgetLayout = QHBoxLayout()
                doneWidgetViewPortLayout = QVBoxLayout()
                doneWidgetViewPortLayout.addStretch()
                doneWidgetViewPortLayout.addWidget(self.doneWidgetButton)
                doneWidgetViewPortLayout.addWidget(self.smallView)
                doneWidgetViewPortLayout.addWidget(self.doneWidgetControlHelpButton)
                doneWidgetViewPortLayout.addStretch()
                doneWidgetHorizontalControlLayout = QHBoxLayout()
                doneWidgetHorizontalControlLayout.addWidget(self.doneWidgetSizeLeftButton)
                doneWidgetHorizontalControlLayout.addWidget(self.doneWidgetSizeRightButton)
                doneWidgetHorizontalControlLayout.addStretch()
                doneWidgetControlLayout = QVBoxLayout()
                doneWidgetControlLayout.addWidget(self.doneWidgetCloseButton)
                doneWidgetUpButtonPadLayout = QHBoxLayout()
                doneWidgetUpButtonPadLayout.addStretch()
                doneWidgetUpButtonPadLayout.addWidget(self.doneWidgetSizeUpButton)
                doneWidgetUpButtonPadLayout.addStretch()
                doneWidgetDownButtonPadLayout = QHBoxLayout()
                doneWidgetDownButtonPadLayout.addStretch()
                doneWidgetDownButtonPadLayout.addWidget(self.doneWidgetSizeDownButton)
                doneWidgetDownButtonPadLayout.addStretch()
                doneWidgetControlLayout.addLayout(doneWidgetUpButtonPadLayout)
                doneWidgetControlLayout.addLayout(doneWidgetHorizontalControlLayout)
                doneWidgetControlLayout.addLayout(doneWidgetDownButtonPadLayout)
                doneWidgetControlLayout.addWidget(self.doneWidgetSetDefaultButton)
                doneWidgetControlLayout.addWidget(self.doneWidgetAutoShowButton)
                doneWidgetControlLayout.addStretch()
                doneWidgetLayout.addLayout(doneWidgetControlLayout)
                doneWidgetLayout.addLayout(doneWidgetViewPortLayout)
                self.doneWidget.setLayout(doneWidgetLayout)

        except Exception as err:
            print(f'Win32com has failed or is not supported: {err} \nWordHook disabled.')

        # Controls label
        self.controlsLabel = QLabel("ⓘ WebEngine: Click and drag to pan, control-scroll to zoom")
        self.topLayout.addStretch()
        self.topLayout.addWidget(self.controlsLabel)

        # Add svg copy button
        self.copyButton = QPushButton("Copy SVG")
        self.copyButton.setStyleSheet("background-color: darkgreen")
        self.copyButton.clicked.connect(self.copySvg)

        # Add svg auto-copy button
        self.autoCopyButton = QPushButton("Auto-Copy")
        self.autoCopyButton.setStyleSheet("background-color: darkred")
        self.autoCopyButton.clicked.connect(self.toggleAutoCopy)

        # Add button to save svg as file
        self.saveButton = QPushButton("Save SVG")
        self.saveButton.setStyleSheet("background-color: #222288")
        self.saveButton.clicked.connect(self.saveSvg)

        # Add button to toggle using physics package
        self.usePhysicsButton = QPushButton("Physics")
        self.usePhysicsButton.setStyleSheet("background-color: darkred")
        self.usePhysicsButton.clicked.connect(self.togglePhysics)

        # Add button to toggle colorsv2 pkg
        self.useColorsv2Button = QPushButton("Colorsv2")
        self.useColorsv2Button.setStyleSheet("background-color: darkred")
        self.useColorsv2Button.clicked.connect(self.toggleColorsv2)

        # Add preamble label
        self.preambleLabel = QLabel("ⓘ Optional Preamble: ")
        # Add hover details for preamble label
        self.preambleLabel.setToolTip('By default, tex2svg adds in all TeX packages except physics and colorsv2.'
                                      '\n You can choose to use them.'
                                      '\n If you use \\color or anything that uses commands from the color package,'
                                      '\n the tex autoloader will automatically load it.')
        # Add developer label :3
        self.developerLabel = QLabel("ⓘ Developed with love by github.com/eljokun")

        # Add button toggles to layout
        self.optionLowerLayout.addWidget(self.preambleLabel)
        self.optionLowerLayout.addWidget(self.usePhysicsButton)
        self.optionLowerLayout.addWidget(self.useColorsv2Button)
        self.optionLowerLayout.addStretch()
        self.optionLowerLayout.addWidget(self.developerLabel)
        self.optionLowerLayout.addStretch()
        self.optionLowerLayout.addWidget(self.saveButton)
        self.optionLowerLayout.addWidget(self.autoCopyButton)
        self.optionLowerLayout.addWidget(self.copyButton)

        # Add button to make window always on top
        self.alwaysOnTopButton = QPushButton("Always On Top")
        self.alwaysOnTopButton.setStyleSheet("background-color: darkred")
        self.alwaysOnTopButton.clicked.connect(self.toggleAlwaysOnTop)
        self.optionInsertionLayout.addWidget(self.alwaysOnTopButton)

        # Add option to switch CDN
        self.cdnButton = QPushButton("Switch CDN")
        self.cdnButton.setStyleSheet("background-color: #222288")
        self.cdnButton.clicked.connect(self.switchCDN)
        self.optionInsertionLayout.addWidget(self.cdnButton)

        # Display style toggle
        self.displayStyleButton = QPushButton("Display Style")
        self.displayStyleButton.setStyleSheet("background-color: darkgreen")
        self.displayStyleButton.clicked.connect(self.toggleDisplayStyle)
        self.optionInsertionLayout.addWidget(self.displayStyleButton)

        # Clear contents
        self.clearButton = QPushButton("Clear")
        self.clearButton.setStyleSheet("background-color: #442222")
        self.clearButton.clicked.connect(lambda: self.equation_edit.clear())
        self.optionInsertionLayout.addWidget(self.clearButton)

        # Insert label
        self.insertLabel = QLabel("ⓘ Insert")
        self.insertLabel.setToolTip('Insert LaTeX classics at your caret position.')
        self.optionInsertionLayout.addWidget(self.insertLabel)

        # Add dfrac
        self.addDFracButton = QPushButton("dfrac")
        self.addDFracButton.setStyleSheet("background-color: #444444")
        self.addDFracButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\dfrac{ }{ }"))
        self.optionInsertionLayout.addWidget(self.addDFracButton)

        # Add text
        self.addTextButton = QPushButton("text")
        self.addTextButton.setStyleSheet("background-color: #444444")
        self.addTextButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\text{  }"))
        self.optionInsertionLayout.addWidget(self.addTextButton)

        # Add cases(system)
        self.addCasesButton = QPushButton("cases (system)")
        self.addCasesButton.setStyleSheet("background-color: #444444")
        self.addCasesButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\begin{cases}    \end{cases}"))
        self.optionInsertionLayout.addWidget(self.addCasesButton)

        # Add partial derivative
        self.addPartialDerivativeButton = QPushButton("partial derivative")
        self.addPartialDerivativeButton.setStyleSheet("background-color: #444444")
        self.addPartialDerivativeButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\dfrac{\partial }{\partial }"))
        self.optionInsertionLayout.addWidget(self.addPartialDerivativeButton)

        # Add tex array button
        self.addTexArrayButton = QPushButton("array")
        self.addTexArrayButton.setStyleSheet("background-color: #444444")
        self.addTexArrayButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\begin{array}{c}  \end{array}"))
        self.optionInsertionLayout.addWidget(self.addTexArrayButton)

        # Add aligned
        self.addAlignedButton = QPushButton("aligned")
        self.addAlignedButton.setStyleSheet("background-color: #444444")
        self.addAlignedButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\begin{aligned}  \end{aligned}"))
        self.optionInsertionLayout.addWidget(self.addAlignedButton)

        # Add limit
        self.addLimitButton = QPushButton("lim")
        self.addLimitButton.setStyleSheet("background-color: #444444")
        self.addLimitButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\lim_{x \to }"))
        self.optionInsertionLayout.addWidget(self.addLimitButton)

        # Add sum and limits
        self.addSumButton = QPushButton("sum")
        self.addSumButton.setStyleSheet("background-color: #444444")
        self.addSumButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\sum\limits_{ }^{ }"))
        self.optionInsertionLayout.addWidget(self.addSumButton)

        # Add 3x3 matrix
        self.addMatrixButton = QPushButton("matrix")
        self.addMatrixButton.setStyleSheet("background-color: #444444")
        self.addMatrixButton.clicked.connect(lambda: self.addTextAtCursorPosition(r"\left[\begin{matrix} \end{matrix}\right]"))
        self.optionInsertionLayout.addWidget(self.addMatrixButton)

        # Add underbrace
        self.addUnderbraceButton = QPushButton("underbrace selection")
        self.addUnderbraceButton.setStyleSheet("background-color: #444444")
        self.addUnderbraceButton.clicked.connect(lambda: self.wrapSelectedText(r"\underbrace{", "}_{ }"))

        # Finalize insertions layout
        self.optionInsertionLayout.addWidget(self.addUnderbraceButton)
        self.optionInsertionLayout.addStretch()

        # Add layouts to main layout
        self.layout.addLayout(self.topLayout)
        self.layout.addLayout(self.interactiveWindowLayout)
        self.layout.addLayout(self.optionLowerLayout)

        # Confirm layout and initialize central widget
        central_widget = QWidget()
        central_widget.setLayout(self.layout)
        self.setCentralWidget(central_widget)
        self.update_mathjax()
        self.equation_edit.textChanged.connect(self.update_mathjax)
    def toggleDisplayStyle(self):
        self.displayStyle = not self.displayStyle
        self.displayStyleButton.setStyleSheet("background-color: darkgreen" if self.displayStyle else "background-color: darkred")
        self.update_mathjax()

    def toggleAlwaysOnTop(self):
        wasMaximized = self.isMaximized()
        if self.windowFlags() & Qt.WindowStaysOnTopHint:
            self.setWindowFlags((self.windowFlags() & ~Qt.WindowStaysOnTopHint) | Qt.WindowCloseButtonHint)
            self.alwaysOnTopButton.setStyleSheet("background-color: darkred")
        else:
            self.setWindowFlags((self.windowFlags() | Qt.WindowStaysOnTopHint) | Qt.WindowCloseButtonHint)
            self.alwaysOnTopButton.setStyleSheet("background-color: darkgreen")
        self.show()
        if wasMaximized:
            self.showMaximized()

    def switchCDN(self):
        default = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg-full.js"
        if not self.customCDN:
            cdn, confirm = QInputDialog.getText(self, "Switch CDN", "Enter CDN URL:", text=default)
            if confirm:
                self.customCDN = True
                self.mathjax_script = f'<script type="text/javascript" async src = "{cdn}"> </script>'
                self.cdnButton.setText("Reset CDN")
                self.load_mathjax()
        else:
            self.customCDN = False
            self.mathjax_script = f'<script type="text/javascript" async src = "{default}"> </script>'
            self.cdnButton.setText("Switch CDN")
            self.load_mathjax()
        self.cdnButton.setStyleSheet("background-color: darkred" if self.customCDN else "background-color: #222288")


    def togglePhysics(self):
        self.physicsEnabled = not self.physicsEnabled
        self.usePhysicsButton.setStyleSheet(
            "background-color: darkgreen" if self.physicsEnabled else "background-color: darkred")
        self.load_mathjax()

    def toggleColorsv2(self):
        self.colorsv2Enabled = not self.colorsv2Enabled
        self.useColorsv2Button.setStyleSheet(
            "background-color: darkgreen" if self.colorsv2Enabled else "background-color: darkred")
        self.load_mathjax()

    def getSvg(self, callback):
        if self.wordHookStatus:
            self.smallView.page().toHtml(callback)
        else:
            self.view.page().toHtml(callback)
    def extractSvgFromHTML(self, html):
        start = html.find('<svg')
        end = html.find('</svg>', start)
        self.svgData = html[start:end + 6].replace('currentColor', 'black')

    def experimentalSvgFileInsertion(self):
        def callback(html):
            pythoncom.CoInitialize()
            temp_file_path = None
            try:
                self.extractSvgFromHTML(html)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".svg") as exptemp:
                    exptemp.write(self.svgData.encode())
                    temp_file_path = os.path.abspath(exptemp.name)
            except Exception as e:
                print(f'Error copying SVG data: {e}')
            if temp_file_path:
                word = win32.gencache.EnsureDispatch('Word.Application')
                try:
                    word.ActiveDocument.InlineShapes.AddPicture(temp_file_path)
                except Exception as e:
                    print(f'Error inserting SVG file: {e}')
            os.unlink(temp_file_path)
            pythoncom.CoUninitialize()
        self.getSvg(callback)




    def copySvg(self):
        def callback(html):
            try:
                self.extractSvgFromHTML(html)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".svg") as temp:
                    temp.write(self.svgData.encode())
                    temp_file_path = temp.name
                pyperclip.copy(temp_file_path)
                with open(temp_file_path, 'r') as f:
                    mimedata = QMimeData()
                    mimedata.setData('image/svg+xml', QByteArray(f.read().encode()))
                    self.clipboard.setMimeData(mimedata)
                threading.Timer(1, os.remove, args=[temp_file_path]).start()
                print(f'Copied SVG data: {self.clipboard.text()}')
            except Exception as e:
                print(f'Error copying SVG data: {e}')

        self.getSvg(callback)

    def saveSvg(self):
        def callback(html):
            self.extractSvgFromHTML(html)
            # File dialog
            savefile, _ = QFileDialog.getSaveFileName(self, 'Save SVG', '', 'SVG files (*.svg)')
            if savefile and not len(self.equation)==0:
                with open(savefile, 'w') as f:
                    f.write(self.svgData)
        self.getSvg(callback)

    def toggleAutoCopy(self):
        self.autoCopy = not self.autoCopy
        self.autoCopyButton.setStyleSheet("background-color: green" if self.autoCopy else "background-color: darkred")
        self.copyButton.setEnabled(not self.autoCopy)
        self.copyButton.setStyleSheet("background-color: gray" if self.autoCopy else "background-color: darkgreen")

    def addTextAtCursorPosition(self, text):
        self.equation_edit.textCursor().insertText(text)

    def wrapSelectedText(self, left, right):
        cursor = self.equation_edit.textCursor()
        selected_text = cursor.selectedText()
        cursor.insertText(left + selected_text + right)
        self.equation_edit.setTextCursor(cursor)

    # MathJax loading for webengine
    # Preamble and allat stuff goes here
    # The idea is to load the script then for every text change update the math content, schedule mathjax render and
    # render/extract (copy if enabled) svg.
    def load_mathjax(self):
        base_packages = ['base', 'ams', 'bbox', 'boldsymbol', 'braket', 'cancel', 'color', 'enclose', 'extpfeil',
                         'html', 'mhchem', 'newcommand', 'noerrors', 'unicode', 'verb', 'autoload', 'require',
                         'configmacros', 'tagformat', 'action', 'bbox', 'boldsymbol', 'colorv2', 'enclose', 'extpfeil',
                         'html', 'mhchem', 'newcommand', 'noerrors', 'unicode', 'verb', 'autoload', 'require',
                         'configmacros', 'tagformat', 'action']

        if self.physicsEnabled:
            base_packages.append('physics')
        else:
            try:
                base_packages.remove('physics')
            except ValueError:
                pass

        # Convert the list of packages to a string
        packages_str = ', '.join(f"'{pkg}'" for pkg in base_packages)

        # Define the loader packages
        loader_packages = ['[tex]/physics'] if self.physicsEnabled else ['[tex]/autoload']
        loader_packages_str = ', '.join(f"'{pkg}'" for pkg in loader_packages)

        html = """
        <!DOCTYPE html>
        <html>
        <style>
        body {{
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                background-color: darkgray;
                overflow: auto;
            }}
        </style>
        <head>
            <script type="text/x-mathjax-config">
                MathJax = {{
                    loader: {{
                        load: [{loader_packages}]
                    }},
                    svg: {{
                        scale: 1,
                        minScale: .1,
                        fontCache: 'global',
                        stroke: 'black',
                    }},
                    tex: {{
                        displayMath: [['$$','$$']],
                        packages: {{'[+]': [{packages}]}}
                    }}
                }};
            </script>
            {mathjax_script}
        </head>
        <body>
            <p id="math-content">
            </p>
            <script>
                var svg;
                var isDragging = false;
                var previousMousePosition;

                document.addEventListener('mousedown', function(event) {{
                    svg = document.querySelector('svg');
                    if (svg) {{
                        isDragging = true;
                        previousMousePosition = {{ x: event.clientX, y: event.clientY }};
                    }}
                }});

                document.addEventListener('mousemove', function(event) {{
                    if (isDragging && svg) {{
                        var dx = event.clientX - previousMousePosition.x;
                        var dy = event.clientY - previousMousePosition.y;
                        var transform = svg.getAttribute('transform') || '';
                        transform += ' translate(' + dx + ' ' + dy + ')';
                        svg.setAttribute('transform', transform);
                        previousMousePosition = {{ x: event.clientX, y: event.clientY }};
                    }}
                }});
                document.addEventListener('mouseup', function(event) {{
                    isDragging = false;
                }});
            </script>
        </body>
        </html>
        """.format(packages=packages_str, loader_packages=loader_packages_str, mathjax_script=self.mathjax_script)

        self.update_mathjax()

        self.view.setHtml(html)

        if self.wordHookStatus:
            self.smallView.setHtml(html)

    def update_mathjax(self):
        plainTextEquation = self.equation_edit.toPlainText()

        def formatted(plainTxtEq):
            if not plainTxtEq:
                if not self.wordHookStatus:
                    plainTxtEq = r"\Large \text{you gonna type something or what?}"
                else:
                    plainTxtEq = (r"")
            return plainTxtEq.replace("\\", "\\\\").replace("\n","\\n").replace("'", "\\'")

        physicsPreamble = formatted(r"\require{physics} ") if self.physicsEnabled else r""

        displayStylePreamble = formatted(r"\displaystyle ") if self.displayStyle else r""

        self.equation = f"{displayStylePreamble}{physicsPreamble}{formatted(plainTextEquation)}"

        script = r"""
        var element = document.getElementById('math-content');
        var svg = MathJax.tex2svg('{}').outerHTML;
        element.innerHTML = svg;
        """.format(self.equation)

        if self.autoCopy:
            self.copySvg()

        if self.wordHookStatus:
            self.smallView.page().runJavaScript(script)

        else:
            self.view.page().runJavaScript(script)
    def update_equation_edit(self, text):
        self.equation_edit.setText(text)
    def start_word_hook(self):
        self.wordHookStatus = True
        self.controlsLabel.hide()
        self.load_mathjax()
        self.view.hide()
        self.showMinimized()
        self.doneWidget.show()
        self.wordHookButton.setStyleSheet("background-color: darkgreen")
        self.wordHookButton.setText("[OVERRIDE] Unhook from MS Word")
        self.wordHookButton.setToolTip('Unhook from MS Word')
        if self.autoCopy:
            self.toggleAutoCopy()
        self.equation_edit.hide()
        if self.alwaysOnTopButton.styleSheet() == "background-color: darkred":
            self.toggleAlwaysOnTop()
        for i in range(self.optionInsertionLayout.count()):
            widget = self.optionInsertionLayout.itemAt(i).widget()
            if widget is not None:
                widget.hide()
        for i in range(self.optionLowerLayout.count()):
            widget = self.optionLowerLayout.itemAt(i).widget()
            if widget is not None:
                widget.hide()
        self.word_polling_thread = threading.Thread(target=self.poll_word_content)
        self.word_polling_thread.daemon = True
        self.word_polling_thread.start()

    def poll_word_content(self):
        pythoncom.CoInitialize()
        word = win32.gencache.EnsureDispatch('Word.Application')
        while self.wordHookStatus:
            try:
                wordDoc = word.ActiveDocument
            except Exception as e:
                print(f'No active word document, or error: {e}')
                self.stop_word_hook()
                return
            try:
                word_content = wordDoc.Range().Text
            except Exception as e:
                print(f'Error reading word content: {e}')
                self.stop_word_hook()
                return
            matches = re.findall(r'\$\$(.*?)\$\$', word_content, re.DOTALL)
            if matches:
                if self.doneWidgetAutoShow:
                    self.doneWidgetAutoShowSignal.emit(True)
                if r"\done" in matches[0]:
                    self.replaceFlag = True
                    matches[0] = matches[0].replace(r"\done", "")
                    self.update_equation_edit_signal.emit(matches[0])
                if self.replaceFlag:
                    if not matches[0]:
                        continue
                    self.copy_svg_thread_safe_signal.emit("1")
                    time.sleep(0.1)
                    start_pos = word_content.find('$$')
                    end_pos = word_content.find('$$', start_pos + 3)
                    if start_pos != -1 and end_pos != -1:
                        wordDoc.Range(start_pos, end_pos + 3).Select()
                        word.Selection.Delete()
                    time.sleep(0.2)
                    self.thread_safe_svg_paste_signal.emit("1")
                    self.update_equation_edit_signal.emit("")
                    self.replaceFlag = False
                    matches.clear()
                    time.sleep(0.1)
                else:
                    if self.doneWidgetAutoShow:
                        self.update_equation_edit_signal.emit(matches[0])
                matches.clear()
            else:
                if self.doneWidgetAutoShow:
                    self.doneWidgetAutoShowSignal.emit(False)
            # Refresh rate for word doc content polling here change this if ur lagging
            time.sleep(1 / 20)
        pythoncom.CoUninitialize()
    def stop_word_hook(self):
        self.wordHookStatus = False
        self.doneWidget.hide()
        self.view.show()
        self.controlsLabel.show()
        self.wordHookButton.setStyleSheet("background-color: darkred")
        self.wordHookButton.setText("Hook to MS Word")
        self.wordHookButton.setToolTip('Hook to MS Word for real-time editing')
        self.equation_edit.clear()
        self.equation_edit.show()
        for i in range(self.optionInsertionLayout.count()):
            widget = self.optionInsertionLayout.itemAt(i).widget()
            if widget is not None:
                widget.show()
        for i in range(self.optionLowerLayout.count()):
            widget = self.optionLowerLayout.itemAt(i).widget()
            if widget is not None:
                widget.show()
        if self.alwaysOnTopButton.styleSheet() == "background-color: darkgreen":
            self.toggleAlwaysOnTop()
        self.show()
    def wordHook(self):
        if not self.wordHookStatus:
            self.start_word_hook()
        else:
            self.stop_word_hook()
        self.update_mathjax()


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    # Hide console
    window.setAttribute(Qt.WA_MacShowFocusRect, False)
    window.show()
    app.exec()
