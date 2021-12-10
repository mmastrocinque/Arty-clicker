import sys
import time

import numpy as np
import pyautogui
import win32com.client
import win32con
import win32gui
import win32ui
import cv2 as cv
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout


class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setWindowFlags(
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.X11BypassWindowManagerHint
        )
        self.setGeometry(
            QtWidgets.QStyle.alignedRect(
                QtCore.Qt.LeftToRight, QtCore.Qt.AlignCenter,
                QtCore.QSize(200, 200),
                QtWidgets.qApp.desktop().availableGeometry()
            ))
        fireBtn = QPushButton('Fire', self)
        fireBtn.resize(100, 32)
        fireBtn.move(50, 50)
        fireBtn.clicked.connect(fireHandler)
        quitBtn = QPushButton('Stop', self)
        quitBtn.resize(100, 32)
        quitBtn.move(50, 100)
        quitBtn.clicked.connect(QtWidgets.qApp.quit)


def fireHandler():
    win32gui.EnumWindows(winEnumHandler, None)
    print('BANG')


def screenshot(hwnd=None):
    if hwnd:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('')
        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
        x, y, x1, y1 = win32gui.GetClientRect(hwnd)
        x, y = win32gui.ClientToScreen(hwnd, (x, y))
        x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
        im = pyautogui.screenshot(region=(x, y, x1, y1))
        return im
    else:
        print('No window given!')
        exit(1)


def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if 'factorio 1' in win32gui.GetWindowText(hwnd).lower():
            print(hex(hwnd), win32gui.GetWindowText(hwnd))
            img = screenshot(hwnd)
            img = np.array(img)
            # img = cv.cvtColor(img,cv.COLOR_RGBA2BGR)

            # hsv = cv.cvtColor(img,cv.COLOR_BGR2HSV)
            lower_range = np.array([153, 15, 15])
            upper_range = np.array([255, 24, 25])
            mask = cv.inRange(img, lower_range, upper_range)
            kernel = np.ones((3, 3), np.uint8)
            mask_erode = cv.erode(mask, kernel, iterations=2)
            masked = cv.bitwise_and(img, img, mask=mask_erode)
            contours, hierarchy = cv.findContours(
                image=mask_erode, mode=cv.RETR_TREE, method=cv.CHAIN_APPROX_NONE)

            image_copy = masked.copy()
            coordsToShoot = []
            for c in contours:
                if cv.contourArea(c) > 0:
                    M = cv.moments(c)
                    cX = int(M["m10"] / M["m00"])
                    cY = int(M["m01"] / M["m00"])
                    coordsToShoot.append((cX,cY))
                    # draw the contour and center of the shape on the image
                    cv.drawContours(image=image_copy, contours=[c], contourIdx=-1, color=(0, 255, 0), thickness=2, lineType=cv.LINE_AA)
                    cv.circle(image_copy, (cX, cY), 7, (255, 255, 255), -1)
                    cv.putText(image_copy, "center", (cX - 20, cY - 20),
                    cv.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 2)
            print(coordsToShoot)

            x, y, x1, y1 = win32gui.GetClientRect(hwnd)
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))

            

            cv.destroyAllWindows()
            cv.imshow('Computer Vision', img)
            cv.imshow('None approximation', image_copy)

            # cv.imshow("mask",mask)
            cv.imshow("mask_erode", mask_erode)
            cv.imshow("masked", masked)

            time.sleep(2)
            for c in coordsToShoot:
                pyautogui.click(c[0] + x ,y + c[1])

        else:
            return


def window_capture(hwnd):
    if hwnd:

        x, y, x1, y1 = win32gui.GetClientRect(hwnd)
        x, y = win32gui.ClientToScreen(hwnd, (x, y))
        x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))

        wDC = win32gui.GetWindowDC(hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, x1, y1)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0, 0), (x1, y1), dcObj, (0, 0), win32con.SRCCOPY)

        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (y1, x1, 4)

        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())

        return img


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = MainWindow()
    mywindow.show()
    app.exec_()
