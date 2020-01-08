
import os
import time
import sys
from pathlib import Path
import zipfile
import win32api
import win32con
import win32file
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QFileDialog, QPushButton, QLabel, QMainWindow, QDesktopWidget, \
    QListWidget, QGridLayout, QMessageBox
import threading
import flash


class App(QMainWindow):

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)

        self.is_flashing = False
        self.is_file_to_burn_zip = False
        self.file_path_to_burn = ""
        self.image_path = None
        self.drive_name = ""
        self.burn_status = ""

        self.initUI()

    def initUI(self):

        self.mwidget = QMainWindow(self)

        self.setWindowFlags(Qt.CustomizeWindowHint | Qt.FramelessWindowHint)
        self.setStyleSheet("background-color: white;border: 1px solid rgb(0,0,0)")

        app_logo = self.resource_path('data_files/app.ico')
        self.setWindowIcon(QtGui.QIcon(app_logo))
        self.setFixedSize(800, 600)
        self.center()
        self.oldPos = self.pos()

        self.set_txt()
        self.set_drive_button()
        self.set_img_button()
        self.set_flash_image()
        self.set_app_button()
        self.set_exit_button()
        self.set_txt_burn_time()
        self.set_burn_img_button()
        self.set_txt_img('')
        self.set_txt_drive('')

        self.show()

    def set_txt(self):

        self.ltime = QLabel(self)
        self.img_labl = QLabel(self)
        self.drive_labl = QLabel(self)

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()
        try:
            self.listwidget.hide()
        except Exception:
            pass

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)

        self.move(qr.topLeft())

    def center_window2(self):

        qr = self.frameGeometry()

        cod = list(qr.getCoords())
        end_qr = (cod[0]+355, cod[1]+305, cod[2]-1000, cod[3]-450)
        # 860, 400, 200, 150
        self.listwidget.setGeometry(end_qr[0], end_qr[1], 100, 120)


    def img_path(self):

        file_path = (QFileDialog.getOpenFileName(self, filter="*.zip; *.img"))[0]
        self.image_path = file_path

        if Path(file_path).suffix == ".zip":
            self.is_file_to_burn_zip = True
        else:
            self.is_file_to_burn_zip = False

        return (file_path)

    def read_img_from_zip(self):

        file_dir_path = os.path.dirname(self.image_path)

        if Path(self.image_path).suffix == ".zip":
            try:
                with zipfile.ZipFile(self.image_path, "r") as theZip:
                    fileNames = theZip.namelist()
                    for fileName in fileNames:
                        if fileName.endswith('img'):
                            theZip.extract(fileName, file_dir_path)

                    img_path = file_dir_path + '/' + fileName
            except Exception:
                pass

            if not os.path.exists(img_path):
                msg = QMessageBox()
                msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Didn't find file to burn")
                msg.setStyleSheet(
                    "QMessageBox{background: rgb(192,192,192);  border: none;font-family: Arial; font-style: normal;  font-size: 10pt; color: #000000 ; }");
                msg.exec_()
                return

        return img_path

    def img_button_pressed(self):
        img_path = self.img_path()
        self.set_txt_img(img_path)
        self.image_path = img_path

    def set_img_button(self):

        self.img_button = QPushButton('Choose File', self)
        self.img_button.clicked.connect(self.img_button_pressed)
        self.img_button.setStyleSheet('background-color: rgb(51,153,255);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
        self.img_button.resize(150,60)
        self.img_button.move(300, 135)
        self.img_button.installEventFilter(self)

    def set_drive_button(self):
        self.drive_button = QPushButton('Choose Drive', self)
        self.drive_button.clicked.connect(self.drive_button_pressed)
        self.drive_button.setStyleSheet('background-color: rgb(51,153,255);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
        self.drive_button.resize(150,60)
        self.drive_button.move(300, 245)
        self.drive_button.installEventFilter(self)

    def set_flash_image(self):
        self.flash_image = QPushButton('', self)
        self.flash_image.setStyleSheet("border: none;")
        flash_image = self.resource_path('data_files/app_flash.png')
        self.flash_image.setIcon(QtGui.QIcon(flash_image))
        self.flash_image.setIconSize(QtCore.QSize(100, 300))
        self.flash_image.move(130, 30)
        self.flash_image.resize(120, 500)

    def eventFilter(self, obj, event):

        if obj == self.drive_button and event.type() == QtCore.QEvent.HoverEnter:
            self.onHovered(self.drive_button)
        elif obj == self.drive_button and event.type() == QtCore.QEvent.HoverLeave:
            self.NotonHovered(self.drive_button)
        elif obj == self.burn_button and event.type() == QtCore.QEvent.HoverEnter:
            self.onHoveredBurn()
        elif obj == self.burn_button and event.type() == QtCore.QEvent.HoverLeave:
            self.NotonHoveredBurn()
        elif obj == self.img_button and event.type() == QtCore.QEvent.HoverEnter:
            self.onHovered(self.img_button)
        elif obj == self.img_button and event.type() == QtCore.QEvent.HoverLeave:
            self.NotonHovered(self.img_button)
        elif obj == self.app_button and event.type() == QtCore.QEvent.HoverEnter:
            self.onupHovered()
        elif obj == self.app_button and event.type() == QtCore.QEvent.HoverLeave:
            self.NotupHovered()

        return super(App, self).eventFilter(obj, event)

    def onHovered(self, button):
        button.setStyleSheet('background-color: rgb(102,178,255);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')

    def NotonHovered(self, button):
        button.setStyleSheet('background-color: rgb(51,153,255);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')

    def onHoveredBurn(self):
        if not self.is_flashing:
            self.burn_button.setStyleSheet('background-color: rgb(255,153,153);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
        else:
            self.burn_button.setStyleSheet('background-color: rgb(224,224,224);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')

    def NotonHoveredBurn(self):
        if not self.is_flashing:
            self.burn_button.setStyleSheet('background-color: rgb(255,102,102);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
        else:
            self.burn_button.setStyleSheet('background-color: rgb(224,224,224);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')


    def onupHovered(self):
        QApplication.setOverrideCursor(QCursor(Qt.PointingHandCursor))

    def NotupHovered(self):
        QApplication.setOverrideCursor(QCursor(Qt.ArrowCursor))

    def set_burn_img_button(self):
        self.burn_button = QPushButton('Flash', self)
        self.burn_button.clicked.connect(self.burn_button_pressed)
        self.burn_button.setStyleSheet('background-color: rgb(255,102,102);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
        self.burn_button.resize(150,60)
        self.burn_button.move(300, 355)
        self.burn_button.installEventFilter(self)

    def drive_button_pressed(self):
        if self.set_drive_items():
            self.listwidget.show()
            self.listwidget.itemClicked.connect(self.drive_select_changed)
        else:
            msg = QMessageBox()
            msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
            msg.setIcon(QMessageBox.Critical)
            msg.setText("\nPlease connect external drive to your computer")
            msg.setStyleSheet(
                "QMessageBox{background: rgb(249,249,249);  border: none;font-family: Arial; font-style: normal;  font-size: 8pt; color: #000000 ; }")
            msg.exec_()
            return

    def drive_select_changed(self):
        self.drive_name = self.listwidget.currentItem().text()
        self.set_txt_drive(self.drive_name)
        self.listwidget.hide()

    def set_drive_items(self):

        all_drives = self.get_drives_list()
        if not all_drives:
            return False
        self.layout = QGridLayout()
        self.setLayout(self.layout)
        self.listwidget = QListWidget()
        self.listwidget.setWindowTitle(" ")
        self.listwidget.setWindowFlags(Qt.CustomizeWindowHint | Qt.FramelessWindowHint)
        self.listwidget.setStyleSheet("border: 0px solid rgb(0,0,0)")

        self.center_window2()

        for count, edrive in enumerate(all_drives):
            self.listwidget.insertItem(count, edrive)

        return True

    def burn_button_pressed(self):

        if not self.drive_name:
            msg = QMessageBox()
            msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
            msg.setIcon(QMessageBox.Critical)
            msg.setText("\nPlease choose Drive before flashing")
            msg.setStyleSheet("QMessageBox{background: rgb(249,249,249);  border: none;font-family: Arial; font-style: normal;  font-size: 8pt; color: #000000 ; }")
            msg.exec_()
            return

        if not self.image_path:
            msg = QMessageBox()
            msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
            msg.setIcon(QMessageBox.Critical)
            msg.setText("\nPlease choose a file to burn before flashing")
            msg.setStyleSheet("QMessageBox{background: rgb(249,249,249);  border: none;font-family: Arial; font-style: normal;  font-size: 8pt; color: #000000 ; }")
            msg.exec_()
            return

        self.is_flashing = True
        self.drive_button.setDisabled(True)
        self.img_button.setDisabled(True)
        self.burn_button.setDisabled(True)
        self.burn_button.setStyleSheet('background-color: rgb(224,224,224);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')

        t = threading.Thread(target=self.win_flash, args=(self.drive_name, self.image_path, ))
        t.start()

    def set_exit_button(self):

        self.exit_button = QPushButton(self)
        self.exit_button.setStyleSheet("border: none;")
        exit_button = self.resource_path('data_files/ExitButton.png')
        self.exit_button.setIcon(QtGui.QIcon(exit_button))
        self.exit_button.setIconSize(QtCore.QSize(24,24))
        self.exit_button.clicked.connect(self.exit_button_pressed)
        self.exit_button.move(740, 15)
        self.exit_button.resize(50, 30)


    def set_logo_button(self):

        self.logo_button = QPushButton(self)
        self.logo_button.setStyleSheet("border: 0px solid rgb(0,0,0)")
        logo_button = self.resource_path('data_files/icon.ico')
        self.logo_button.setIcon(QtGui.QIcon(logo_button))
        self.logo_button.setIconSize(QtCore.QSize(50,50))
        self.logo_button.move(100, 100)
        self.logo_button.resize(50, 50)

    def set_app_button(self):

        self.app_button = QPushButton(self)
        self.app_button.setStyleSheet("border: none;")
        app_button = self.resource_path('data_files/app.png')
        self.app_button.setIcon(QtGui.QIcon(app_button))
        self.app_button.setIconSize(QtCore.QSize(800, 135))
        self.app_button.clicked.connect(self.app_button_pressed)
        self.app_button.resize(800,140)
        self.app_button.setMaximumWidth(780)
        self.app_button.move(15, 450)
        self.app_button.installEventFilter(self)

    def app_button_pressed(self):
        pass
        #webbrowser.open('website domain')

    def exit_button_pressed(self):
        sys.exit(app.exec_())

    def set_txt_img(self, img_path):

        file_name = Path(img_path).name
        if len(file_name) > 9:
            file_name = (Path(file_name).stem)[:4] + "..." + (Path(file_name).stem)[-2:] + Path(img_path).suffix

        #self.img_labl = QLabel(self)
        self.img_labl.setText(file_name)
        self.img_labl.setStyleSheet("border: 0px solid rgb(0,0,0)")
        self.img_labl.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Black))
        self.img_labl.move(560, 155)
        self.img_labl.adjustSize()

    def set_txt_drive(self, drive_name):
        if drive_name:
            self.drive_labl.setText("Chosen Drive - {}".format(drive_name))
            self.drive_labl.setStyleSheet("border: 0px solid rgb(0,0,0)")
            self.drive_labl.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Black))
            self.drive_labl.move(560, 265)
            self.drive_labl.adjustSize()

        else:
            self.drive_labl.setText("")
            self.drive_labl.setStyleSheet("border: 0px solid rgb(0,0,0)")
            self.drive_labl.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Black))
            self.drive_labl.move(560, 265)
            self.drive_labl.adjustSize()


    def set_txt_burn_time(self):
        self.ltime.setText("")
        self.ltime.setStyleSheet("border: 0px solid rgb(0,0,0)")
        self.ltime.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Black))
        self.ltime.move(560, 375)
        self.ltime.adjustSize()

    def set_new(self):
        self.new.setText("bis")
        self.new.setStyleSheet("border: 0px solid rgb(0,0,0)")
        self.new.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Black))
        self.new.move(200, 200)
        self.new.adjustSize()


    def get_drives_list(self):
        drives = [i for i in win32api.GetLogicalDriveStrings().split('\x00') if i]
        rdrives = [d for d in drives if win32file.GetDriveType(d) == win32con.DRIVE_REMOVABLE]
        return rdrives

    def win_flash(self, drive, image): #data_callback, progress_callback):

        # there are images left to burn, pick the first one
        image = image
        drive = drive
        logfile = "C:\\Users\\Public\\burntime.log"
        if not os.path.exists(os.path.dirname(logfile)):
            logfile = "C:\\flashtime.log"

        try:
            os.remove(logfile)
        except Exception:
            pass

        if self.is_file_to_burn_zip:
            self.ltime.setText("unzipping img...")
            self.ltime.move(560, 375)
            self.ltime.adjustSize()
            image = self.read_img_from_zip()

        # remove the trailing \ on the drive name
        if "\\" in drive:
            drive = drive[:-1]


        try:
            p = threading.Thread(target=flash.start, args=(image, drive, logfile,))
            p.start()
            while True:
                try:
                    lf = open(logfile, 'rt')
                    break
                except Exception:
                    self.ltime.setText("Start flashing...")
                    self.ltime.move(560, 375)
                    self.ltime.adjustSize()
                    time.sleep(1)

            while p.isAlive():
                #  capturing progress via a file
                l = lf.readline().strip("\n")

                if len(l) != 0:
                    self.ltime.setText(str(l) + " completed")
                    self.ltime.move(560, 375)
                    self.ltime.adjustSize()

                    # check for errors
                    if l.startswith("ERROR"):
                        print("Error detected:\n{}".format(l))
                        return False

            self.ltime.setText("Flash completed!")
            self.ltime.move(560, 375)
            self.ltime.adjustSize()
            #  close the log file
            if lf:
                lf.close()

            # check for return code
            if p.returncode == 0:
                # All ok, pop the image from the list
                os.remove(logfile)
                self.burn_button.setStyleSheet('background-color: rgb(255,102,102);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
                self.is_flashing = False
                self.drive_button.setDisabled(False)
                self.img_button.setDisabled(False)
                self.burn_button.setDisabled(False)
                return "Done"
            else:
                os.remove(logfile)
                self.burn_button.setStyleSheet('background-color: rgb(255,102,102);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
                self.is_flashing = False
                self.drive_button.setDisabled(False)
                self.img_button.setDisabled(False)
                self.burn_button.setDisabled(False)
                return "Oops!"

        except OSError as e:
            os.remove(logfile)
            self.burn_button.setStyleSheet('background-color: rgb(255,102,102);border-style: outset;border-width: 2px;border-radius: 200px;border-color: beige;color: rgb(255,255,255);font: bold 16px;min-width: 10em;padding: 6px;')
            self.drive_button.setDisabled(False)
            self.img_button.setDisabled(False)
            self.burn_button.setDisabled(False)
            self.is_flashing = False
            #logging.debug("Failed to execute program '%s': %s" % (cmd, str(e)))
            raise

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())