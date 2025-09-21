import cgitb
import ctypes
import time
import keyboard
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QCoreApplication, QMetaObject, QRect, QObject
from PyQt5.QtGui import QFont, QCursor, QKeySequence, QIcon, QPixmap
from PyQt5.QtWidgets import QWidget, QLineEdit, QToolButton, QLabel, QApplication, QShortcut
import sys
import matplotlib.pyplot as plt
from matplotlib import mathtext
import win32con
from ctypes import windll, byref
from ctypes.wintypes import MSG
import tkinter.messagebox
import pygetwindow as gw
import win32clipboard
from PIL import Image
import io
import os
import base64
from win10toast import ToastNotifier
import tkinter

VERSION = "1.1"
FILEDIR = "Latexer"

LatexerIcon = b'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'
if not os.path.exists(FILEDIR):
    os.mkdir(FILEDIR)
with open(FILEDIR + "/LatexerIcon.ico", "wb") as f:
    f.write(base64.b64decode(LatexerIcon))

tk = tkinter.Tk()
tk.title("Latexer")
tk.withdraw()
tk.iconbitmap(FILEDIR + "/LatexerIcon.ico")

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Latexer")
sys.excepthook = cgitb.Hook(1, None, 5, sys.stderr, 'text')


def copyImg(path):
    image = Image.open(path)

    output = io.BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


# matplotlib.use("pgf")
# pgf_config = {
#     "font.family": 'serif',
#     "font.size": 20,
#     "pgf.rcfonts": False,
#     "text.usetex": True,
#     "pgf.preamble": " ".join([
#         r"\usepackage{unicode-math}",
#         # r"\setmathfont{XITS Math}",
#         r"\setmainfont{Times New Roman}",
#         r"\usepackage{xeCJK}",
#         r"\xeCJKsetup{CJKmath=true}",
#         r"\setCJKmainfont{SimSun}",
#     ]),
# }
# rcParams.update(pgf_config)


# plt.rcParams['text.usetex'] = True
# plt.rcParams['text.latex.preamble'] = r'\usepackage{amsmath}'

config = {
    "font.family": 'serif',
    "font.size": 20,
    "mathtext.fontset": 'stix',
    "font.serif": ['SimSun'],
}
plt.rcParams.update(config)

history = []


class HotKey(QThread):
    showWindow = pyqtSignal(int)
    error = pyqtSignal()

    def __init__(self):
        super(HotKey, self).__init__()
        self.main_key = 69

    def run(self):
        user32 = windll.user32
        if not user32.RegisterHotKey(None, 1, win32con.MOD_ALT, self.main_key):
            self.error.emit()
            return
        else:
            notifier = ToastNotifier()
            notifier.show_toast('Latexer', '程序已启动，按 Alt + E 呼出面板。输入 //help 查看帮助。', duration=5,
                                threaded=True,
                                icon_path=FILEDIR + "/LatexerIcon.ico")
        while True:
            try:
                msg = MSG()
                if user32.GetMessageA(byref(msg), None, 0, 0) != 0:
                    if msg.message == win32con.WM_HOTKEY:
                        if msg.wParam == win32con.MOD_ALT:
                            self.showWindow.emit(msg.lParam)
            finally:
                user32.UnregisterHotKey(None, 1)
                if not user32.RegisterHotKey(None, 1, win32con.MOD_ALT, self.main_key):
                    self.error.emit()
                    return


class Widget(QWidget):
    showWindow = pyqtSignal(int)

    def __init__(self):
        super(Widget, self).__init__()
        self.setObjectName("Form")
        self.resize(280, 259)
        ico_path = FILEDIR + "/LatexerIcon.ico"
        icon = QIcon()
        icon.addPixmap(QPixmap(ico_path), QIcon.Normal, QIcon.Off)
        font = QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.setFont(font)
        self.lineEdit = QLineEdit(self)
        self.lineEdit.setGeometry(QRect(10, 10, 222, 31))
        font = QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.toolButton = QToolButton(self)
        self.toolButton.setGeometry(QRect(240, 10, 31, 31))
        self.toolButton.setObjectName("toolButton")
        self.label = QLabel(self)
        self.label.setGeometry(QRect(10, 50, 0, 0))
        font = QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(9)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")

        self.label.setStyleSheet("background:rgba(255,255,255,.7);color:red;border-radius:3px;padding:3px;")
        self.lineEdit.setStyleSheet("border-radius:3px;")
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Drawer)
        self.toolButton.clicked.connect(self.convert)
        self.status = False
        self.lineEdit.setAttribute(Qt.WA_InputMethodEnabled, False)
        self.historyPos = 0
        QShortcut(QKeySequence("return"), self.toolButton, self.convert)
        QShortcut(QKeySequence("enter"), self.toolButton, self.convert)
        QShortcut(QKeySequence("escape"), self, self.switchWindow)
        QShortcut(QKeySequence("up"), self.lineEdit, lambda: self.lookHistory(-1))
        QShortcut(QKeySequence("down"), self.lineEdit, lambda: self.lookHistory(1))

        self.retranslateUi()
        QMetaObject.connectSlotsByName(self)

    def lookHistory(self, dir):
        if not history:
            return
        if dir == 1 and self.historyPos >= -1:
            return
        if dir == -1 and len(history) >= - self.historyPos + 1:
            self.historyPos -= 1
        if dir == 1 and self.historyPos < -1:
            self.historyPos += 1
        self.lineEdit.setText(history[self.historyPos])
        print(self.historyPos, history)

    def setLabel(self, text):
        self.label.setText(text)
        labelWidth = 261
        labelHeight = self.label.heightForWidth(labelWidth)
        print(labelHeight)
        if labelHeight > 500:
            labelWidth = 800
            labelHeight = self.label.heightForWidth(labelWidth)
        self.label.setGeometry(QRect(10, 50, labelWidth, labelHeight))
        self.resize(280 + labelWidth, 259 + labelHeight)

    def convert(self):
        self.toolButton.setEnabled(False)
        self.label.setText("")
        self.label.setGeometry(QRect(10, 50, 0, 0))
        latex = self.lineEdit.text().strip().strip("$").strip()
        if latex[0:2] == "//":
            op = latex[2:]
            if op == "info":
                tkinter.messagebox.showinfo("关于", "Latexer | Version " + str(
                    VERSION) + "\nLicensed Under The MIT License\n\nAuthor: HShiDianLu.\nGithub: https://github.com/HShiDianLu/Latexer")
            elif op == "help":
                tkinter.messagebox.showinfo("热键与指令说明",
                                            "Alt + K - 打开主面板\n上下箭头 - 查看历史纪录\n\n//info - 打开关于页面\n//help - 打开此页面\n//exit - 退出程序")
            elif op == "exit":
                sys.exit()
            else:
                self.setLabel("无法识别的指令: " + op + "。\n输入 //help 查看帮助。")
            self.toolButton.setEnabled(True)
            return
        latex = "$" + latex + "$"
        if latex.replace(" ", "") == "$$":
            self.switchWindow()
            self.toolButton.setEnabled(True)
            return
        filename = FILEDIR + "/latex_" + str(time.time()).replace(".", "") + ".png"
        try:
            mathtext.math_to_image(latex, filename, dpi=140)
        except Exception as e:
            print(str(e))
            self.setLabel("Latex 渲染失败:\n" + str(e))
            self.toolButton.setEnabled(True)
            return
        self.toolButton.setEnabled(True)
        self.switchWindow()
        copyImg(filename)
        keyboard.press_and_release('ctrl+v')

    def hotKeyCallback(self, data):
        self.switchWindow()

    def switchWindow(self):
        print("showWindow")
        if not self.status:
            self.show()
            self.move(QCursor.pos().x() + 10, QCursor.pos().y())
            self.status = True
            self.lineEdit.setText("$  $")
            self.lineEdit.setCursorPosition(2)
            gw.getWindowsWithTitle('Latexer UI')[0].activate()
            self.lineEdit.setFocus()
            self.label.setText("")
            self.label.setGeometry(QRect(10, 50, 0, 0))
            self.historyPos = 0
        else:
            latex = self.lineEdit.text().strip().strip("$").strip()
            latex = "$" + latex + "$"
            if latex.replace(" ", "") != "$$" and ((not history) or self.lineEdit.text() != history[self.historyPos]):
                history.append(self.lineEdit.text())
            self.close()
            self.status = False

    def retranslateUi(self):
        _translate = QCoreApplication.translate
        self.setWindowTitle(_translate("Form", "Latexer UI"))
        self.toolButton.setText(_translate("Form", "OK"))


def errorCallback():
    tkinter.messagebox.showerror("错误",
                                 "全局热键注册失败。这可能是由于：\n1. 已经有一个正在运行的程序了。\n2. 热键冲突，请检查有无占用了 Alt + E 热键的程序。\n\n按确定退出程序。")
    sys.exit()


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    ui = Widget()
    listener = HotKey()
    listener.showWindow.connect(ui.hotKeyCallback)
    listener.error.connect(errorCallback)
    listener.start()
    sys.exit(app.exec_())
