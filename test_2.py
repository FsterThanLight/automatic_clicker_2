import math
import sys
import time

import cv2
from PyQt5.QtCore import QPoint, QRectF
from PyQt5.QtCore import QRect, Qt, pyqtSignal, QSettings
from PyQt5.QtGui import QBrush, QWindow
from PyQt5.QtGui import QPixmap, QPainter, QPen, QFont, QColor
from PyQt5.QtWidgets import QApplication, QLabel
from numpy import uint8, array
from pynput.mouse import Controller


#
#
def get_opposite_color(color: QColor):
    return QColor(255 - color.red(), 255 - color.green(), 255 - color.blue())


def get_line_interpolation(p1, p2):  # 线性插值
    res = []
    dy = p1[1] - p2[1]
    dx = p1[0] - p2[0]
    n = max(abs(dy), abs(dx))
    nx = dx / n
    ny = dy / n
    for i in range(n):
        res.append([p2[0] + i * nx, p2[1] + i * ny])
    return res


class Finder:  # 选择智能选区
    def __init__(self, parent):
        self.h = self.w = 0
        self.rect_list = self.contours = []
        self.area_threshold = 200
        self.parent = parent
        self.img = None

    def find_contours_setup(self):

        try:
            self.area_threshold = self.parent.parent.ss_areathreshold.value()
        except Exception as e:
            print('find_contours_setup', e)
            self.area_threshold = 200
        self.h, self.w, _ = self.img.shape

        gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)  # 灰度化
        th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 5, 2)  # 自动阈值
        self.contours = cv2.findContours(th, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)[-2]
        self.find_contours()

    def find_contours(self):
        draw_img = cv2.drawContours(self.img.copy(), self.contours, -1, (0, 255, 0), 1)
        self.rect_list = [[0, 0, self.w, self.h]]
        for i in self.contours:
            x, y, w, h = cv2.boundingRect(i)
            area = cv2.contourArea(i)
            if area > self.area_threshold and w > 10 and h > 10:
                self.rect_list.append([x, y, x + w, y + h])
        print('contours:', len(self.contours), 'left', len(self.rect_list))

    def find_targetrect(self, point):
        target_rect = [0, 0, self.w, self.h]
        target_area = 1920 * 1080
        for rect in self.rect_list:
            if point[0] in range(rect[0], rect[2]):
                if point[1] in range(rect[1], rect[3]):
                    area = (rect[3] - rect[1]) * (rect[2] - rect[0])
                    if area < target_area:
                        target_rect = rect
                        target_area = area
        return target_rect

    def clear_setup(self):
        self.h = self.w = 0
        self.rect_list = self.contours = []
        self.img = None


class MaskLayer(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setMouseTracking(True)

    def paintEvent(self, e):
        super().paintEvent(e)
        if self.parent.on_init:
            print('oninit return')
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        if self.parent.painter_tools["perspective_cut_on"]:  # 透视裁剪工具
            # painter.setPen(QPen(self.parent.pencolor, 3, Qt.SolidLine))
            color = get_opposite_color(self.parent.pencolor)
            for i in range(len(self.parent.perspective_cut_pointlist)):
                painter.setPen(QPen(color, 10, Qt.SolidLine))
                painter.drawPoint(
                    QPoint(self.parent.perspective_cut_pointlist[i][0], self.parent.perspective_cut_pointlist[i][1]))
                painter.setPen(QPen(color, 3, Qt.SolidLine))
                if i < len(self.parent.perspective_cut_pointlist) - 1:
                    painter.drawLine(self.parent.perspective_cut_pointlist[i][0],
                                     self.parent.perspective_cut_pointlist[i][1],
                                     self.parent.perspective_cut_pointlist[i + 1][0],
                                     self.parent.perspective_cut_pointlist[i + 1][1])
                else:
                    painter.drawLine(self.parent.perspective_cut_pointlist[i][0],
                                     self.parent.perspective_cut_pointlist[i][1],
                                     self.parent.mouse_posx, self.parent.mouse_posy)
                    painter.drawLine(self.parent.perspective_cut_pointlist[0][0],
                                     self.parent.perspective_cut_pointlist[0][1],
                                     self.parent.mouse_posx, self.parent.mouse_posy)
            # 画网格
            painter.setPen(QPen(QColor(120, 180, 120, 180), 1, Qt.SolidLine))
            if len(self.parent.perspective_cut_pointlist) >= 2:
                p0 = self.parent.perspective_cut_pointlist[0]
                p1 = self.parent.perspective_cut_pointlist[1]
                pp1 = pp0 = (self.parent.mouse_posx, self.parent.mouse_posy)
                if len(self.parent.perspective_cut_pointlist) > 2:
                    pp1 = self.parent.perspective_cut_pointlist[2]
                dx1 = pp1[0] - p1[0]
                dy1 = pp1[1] - p1[1]
                dx0 = pp0[0] - p0[0]
                dy0 = pp0[1] - p0[1]
                maxs = max(math.sqrt(dy0 ** 2 + dx0 ** 2), math.sqrt(dy1 ** 2 + dx1 ** 2))
                if maxs > 25:
                    n = maxs // 25
                    ddx0 = dx0 / (n + 1)
                    ddy0 = dy0 / (n + 1)
                    ddx1 = dx1 / (n + 1)
                    ddy1 = dy1 / (n + 1)
                    for i in range(int(n) + 1):
                        painter.drawLine(pp0[0] - i * ddx0, pp0[1] - i * ddy0, pp1[0] - i * ddx1, pp1[1] - i * ddy1)
            if len(self.parent.perspective_cut_pointlist) >= 3:
                p0 = self.parent.perspective_cut_pointlist[1]
                p1 = self.parent.perspective_cut_pointlist[2]
                pp1 = (self.parent.mouse_posx, self.parent.mouse_posy)
                pp0 = self.parent.perspective_cut_pointlist[0]

                dx1 = pp1[0] - p1[0]
                dy1 = pp1[1] - p1[1]
                dx0 = pp0[0] - p0[0]
                dy0 = pp0[1] - p0[1]
                maxs = max(math.sqrt(dy0 ** 2 + dx0 ** 2), math.sqrt(dy1 ** 2 + dx1 ** 2))
                if maxs > 25:
                    n = maxs // 25
                    ddx0 = dx0 / (n + 1)
                    ddy0 = dy0 / (n + 1)
                    ddx1 = dx1 / (n + 1)
                    ddy1 = dy1 / (n + 1)
                    for i in range(int(n) + 1):
                        painter.drawLine(pp0[0] - i * ddx0, pp0[1] - i * ddy0, pp1[0] - i * ddx1, pp1[1] - i * ddy1)

        elif self.parent.painter_tools["polygon_ss_on"]:  # 多边形截图
            color = get_opposite_color(self.parent.pencolor)
            for i in range(len(self.parent.polygon_ss_pointlist)):
                painter.setPen(QPen(color, 3, Qt.SolidLine))
                if i < len(self.parent.polygon_ss_pointlist) - 1:
                    painter.drawLine(self.parent.polygon_ss_pointlist[i][0],
                                     self.parent.polygon_ss_pointlist[i][1],
                                     self.parent.polygon_ss_pointlist[i + 1][0],
                                     self.parent.polygon_ss_pointlist[i + 1][1])
                else:

                    painter.drawLine(self.parent.polygon_ss_pointlist[i][0],
                                     self.parent.polygon_ss_pointlist[i][1],
                                     self.parent.mouse_posx, self.parent.mouse_posy)
                    painter.setPen(QPen(QColor(200, 200, 200, 222), 2, Qt.DashDotLine))
                    painter.drawLine(self.parent.polygon_ss_pointlist[0][0],
                                     self.parent.polygon_ss_pointlist[0][1],
                                     self.parent.mouse_posx, self.parent.mouse_posy)

        elif not (self.parent.painter_tools['selectcolor_on'] or self.parent.painter_tools['bucketpainter_on']):
            # 正常显示选区
            rect = QRect(min(self.parent.x0, self.parent.x1), min(self.parent.y0, self.parent.y1),
                         abs(self.parent.x1 - self.parent.x0), abs(self.parent.y1 - self.parent.y0))

            painter.setPen(QPen(Qt.green, 2, Qt.SolidLine))
            painter.drawRect(rect)
            painter.drawRect(0, 0, self.width(), self.height())
            painter.setPen(QPen(QColor(0, 150, 0), 8, Qt.SolidLine))
            painter.drawPoint(
                QPoint(self.parent.x0, min(self.parent.y1, self.parent.y0) + abs(self.parent.y1 - self.parent.y0) // 2))
            painter.drawPoint(
                QPoint(min(self.parent.x1, self.parent.x0) + abs(self.parent.x1 - self.parent.x0) // 2, self.parent.y0))
            painter.drawPoint(
                QPoint(self.parent.x1, min(self.parent.y1, self.parent.y0) + abs(self.parent.y1 - self.parent.y0) // 2))
            painter.drawPoint(
                QPoint(min(self.parent.x1, self.parent.x0) + abs(self.parent.x1 - self.parent.x0) // 2, self.parent.y1))
            painter.drawPoint(QPoint(self.parent.x0, self.parent.y0))
            painter.drawPoint(QPoint(self.parent.x0, self.parent.y1))
            painter.drawPoint(QPoint(self.parent.x1, self.parent.y0))
            painter.drawPoint(QPoint(self.parent.x1, self.parent.y1))

            x = y = 100
            if self.parent.x1 > self.parent.x0:
                x = self.parent.x0 + 5
            else:
                x = self.parent.x0 - 72
            if self.parent.y1 > self.parent.y0:
                y = self.parent.y0 + 15
            else:
                y = self.parent.y0 - 5
            painter.setPen(QPen(Qt.darkGreen, 2, Qt.SolidLine))
            painter.drawText(x, y,
                             '{}x{}'.format(abs(self.parent.x1 - self.parent.x0), abs(self.parent.y1 - self.parent.y0)))

            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(0, 0, 0, 120))
            painter.drawRect(0, 0, self.width(), min(self.parent.y1, self.parent.y0))
            painter.drawRect(0, min(self.parent.y1, self.parent.y0), min(self.parent.x1, self.parent.x0),
                             self.height() - min(self.parent.y1, self.parent.y0))
            painter.drawRect(max(self.parent.x1, self.parent.x0), min(self.parent.y1, self.parent.y0),
                             self.width() - max(self.parent.x1, self.parent.x0),
                             self.height() - min(self.parent.y1, self.parent.y0))
            painter.drawRect(min(self.parent.x1, self.parent.x0), max(self.parent.y1, self.parent.y0),
                             max(self.parent.x1, self.parent.x0) - min(self.parent.x1, self.parent.x0),
                             self.height() - max(self.parent.y1, self.parent.y0))

        if not (self.parent.painter_tools['drawcircle_on'] or self.parent.painter_tools['drawrect_bs_on'] or
                self.parent.painter_tools['pen_on'] or self.parent.painter_tools['eraser_on'] or
                self.parent.painter_tools['drawtext_on'] or self.parent.painter_tools['backgrounderaser_on']
                or self.parent.painter_tools['drawpix_bs_on'] or self.parent.move_rect):

            select_color_mode = True if self.parent.painter_tools['selectcolor_on'] or self.parent.painter_tools[
                'bucketpainter_on'] else False  # 取色器或油漆桶

            if self.parent.mouse_posx > self.width() - 140:
                enlarge_box_x = self.parent.mouse_posx - 140
            else:
                enlarge_box_x = self.parent.mouse_posx + 20
            if self.parent.mouse_posy > self.height() - 140:
                enlarge_box_y = self.parent.mouse_posy - 120
            else:
                enlarge_box_y = self.parent.mouse_posy + 20
            enlarge_rect = QRect(enlarge_box_x, enlarge_box_y, 120, 120)
            painter.setPen(QPen(QColor(255, 192, 203), 1, Qt.SolidLine))
            painter.drawRect(enlarge_rect)
            painter.setBrush(QBrush(QColor(255, 182, 193, 180)))
            painter.drawRect(QRect(enlarge_box_x, enlarge_box_y - 43, enlarge_rect.width(), 43))
            painter.setBrush(Qt.NoBrush)
            # painter.drawRect(QRect(enlarge_box_x, enlarge_box_y-42, enlarge_rect.width(), 42))
            color = QColor(self.parent.qimg.pixelColor(self.parent.mouse_posx, self.parent.mouse_posy))
            RGB_color = [color.red(), color.green(), color.blue()]
            HSV_color = cv2.cvtColor(array([[RGB_color]], dtype=uint8), cv2.COLOR_RGB2HSV).tolist()[0][0]
            painter.drawText(enlarge_box_x, enlarge_box_y - 6,
                             ' POS:({},{}) {}'.format(self.parent.mouse_posx, self.parent.mouse_posy,
                                                      color.name().upper() if select_color_mode else ""))
            painter.drawText(enlarge_box_x, enlarge_box_y - 18,
                             " HSV:({},{},{})".format(HSV_color[0], HSV_color[1], HSV_color[2]))
            painter.drawText(enlarge_box_x, enlarge_box_y - 30,
                             " RGB:({},{},{})".format(RGB_color[0], RGB_color[1], RGB_color[2]))

            if select_color_mode:
                painter.setBrush(QBrush(color))
                painter.drawRect(QRect(enlarge_box_x - 20, enlarge_box_y, 20, 20))
                painter.setBrush(Qt.NoBrush)

            try:  # 鼠标放大镜
                painter.setCompositionMode(QPainter.CompositionMode_Source)
                rpix = QPixmap(self.width() + 120, self.height() + 120)
                rpix.fill(QColor(0, 0, 0))
                rpixpainter = QPainter(rpix)
                rpixpainter.drawPixmap(60, 60, self.parent.pixmap())
                rpixpainter.end()
                larger_pix = rpix.copy(self.parent.mouse_posx, self.parent.mouse_posy, 120, 120).scaled(
                    120 + self.parent.tool_width * 10, 120 + self.parent.tool_width * 10)
                pix = larger_pix.copy(larger_pix.width() // 2 - 60, larger_pix.height() // 2 - 60, 120, 120)
                painter.drawPixmap(enlarge_box_x, enlarge_box_y, pix)
                painter.setPen(QPen(QColor(255, 192, 203), 1, Qt.SolidLine))
                painter.drawLine(enlarge_box_x, enlarge_box_y + 60, enlarge_box_x + 120, enlarge_box_y + 60)
                painter.drawLine(enlarge_box_x + 60, enlarge_box_y, enlarge_box_x + 60, enlarge_box_y + 120)
            except Exception as e:
                print('draw_enlarge_box fail', e)

        painter.end()


class PaintLayer(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pixpng = None
        self.pixPainter = None
        self.parent = parent
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setMouseTracking(True)
        self.px = self.py = -50
        # self.pixpng = QPixmap(":/msk.jpg")

    def paintEvent(self, e):
        super().paintEvent(e)
        if self.parent.on_init:
            print('oninit return')
            return
        if 1 in self.parent.painter_tools.values():  # 如果有画笔工具打开
            painter = QPainter(self)
            color = QColor(self.parent.pencolor)
            color.setAlpha(255)

            width = self.parent.tool_width
            if self.parent.painter_tools['selectcolor_on'] or self.parent.painter_tools['bucketpainter_on']:
                width = 5
                color = QColor(Qt.white)
            painter.setPen(QPen(color, 1, Qt.SolidLine))
            rect = QRectF(self.px - width // 2, self.py - width // 2,
                          width, width)
            painter.drawEllipse(rect)  # 画鼠标圆
            painter.end()
        try:
            self.pixPainter = QPainter(self.pixmap())
            self.pixPainter.setRenderHint(QPainter.Antialiasing)
        except Exception as e:
            print(f'painter error:{e}')

        if len(self.parent.drawtext_pointlist) > 1:
            self.parent.text_box.paint = False
            text = self.parent.text_box.toPlainText()
            self.parent.text_box.clear()
            pos = self.parent.drawtext_pointlist.pop(0)
            if text:
                self.pixPainter.setFont(QFont('', self.parent.tool_width))
                self.pixPainter.setPen(QPen(self.parent.pencolor, 3, Qt.SolidLine))
                self.pixPainter.drawText(pos[0] + self.parent.text_box.document.size().height() / 8,
                                         pos[1] + self.parent.text_box.document.size().height() * 32 / 41, text)
                self.parent.backup_shortshot()
                self.parent.setFocus()
                self.update()
        try:
            self.pixPainter.end()
        except Exception as e:
            print('pixpainter end fail!', e)


class Slabel(QLabel):  # 区域截图功能
    showm_signal = pyqtSignal(str)
    recorder_recordchange_signal = pyqtSignal()
    close_signal = pyqtSignal()
    ocr_image_signal = pyqtSignal(str)
    screen_shot_result_signal = pyqtSignal(str)
    screen_shot_end_show_sinal = pyqtSignal(QPixmap)
    set_area_result_signal = pyqtSignal(list)
    getpix_result_signal = pyqtSignal(tuple, QPixmap)

    def __init__(self, parent=None):
        super().__init__()
        self.smartcursor_on = None
        self.mouse_posy = None
        self.choicing = None
        self.mouse_posx = None
        self.NpainterNmoveFlag = None
        self.originalPix = None
        self.on_init = None
        self.backup_pic_list = None
        self.backup_ssid = None
        self.finder = None
        self.settings = None
        self.mask = None
        self.paintlayer = None
        self.mode = None
        self.sshoting = None
        self.parent = parent
        # if not os.path.exists("j_temp"):
        #     os.mkdir("j_temp")

    def setup(self, mode="screenshot"):  # 初始化界面
        self.on_init = True
        self.mode = mode
        self.paintlayer = PaintLayer(self)  # 绘图层
        self.mask = MaskLayer(self)  # 遮罩层
        self.settings = QSettings('Fandes', 'jamtools')
        self.setMouseTracking(True)
        self.finder = Finder(self)  # 智能选区的寻找器
        self.init_parameters()
        self.backup_ssid = 0  # 当前备份数组的id,用于确定回退了几步
        self.backup_pic_list = []  # 备份页面的数组,用于前进/后退
        self.on_init = False

    def init_parameters(self):  # 初始化参数
        self.NpainterNmoveFlag = self.choicing = self.move_rect = self.move_y0 = self.move_x0 = self.move_x1 \
            = self.change_alpha = self.move_y1 = False
        self.x0 = self.y0 = self.rx0 = self.ry0 = self.x1 = self.y1 = self.mouse_posx = self.mouse_posy = -50
        self.bx = self.by = 0
        self.alpha = 255  # 透明度值
        self.smartcursor_on = self.settings.value("screenshot/smartcursor", True, type=bool)
        self.finding_rect = True  # 正在自动寻找选取的控制变量,就进入截屏之后会根据鼠标移动到的位置自动选取,
        self.tool_width = 5
        self.roller_area = (0, 0, 1, 1)
        self.backgrounderaser_pointlist = []  # 下面xxpointlist都是储存绘图数据的列表
        self.eraser_pointlist = []
        self.pen_pointlist = []
        self.drawpix_pointlist = []
        self.repairbackground_pointlist = []
        self.drawtext_pointlist = []
        self.perspective_cut_pointlist = []
        self.polygon_ss_pointlist = []
        self.drawrect_pointlist = [[-2, -2], [-2, -2], 0]
        self.drawarrow_pointlist = [[-2, -2], [-2, -2], 0]
        self.drawcircle_pointlist = [[-2, -2], [-2, -2], 0]
        self.painter_tools = {'drawpix_bs_on': 0, 'drawarrow_on': 0, 'drawcircle_on': 0, 'drawrect_bs_on': 0,
                              'pen_on': 0, 'eraser_on': 0, 'drawtext_on': 0,
                              'backgrounderaser_on': 0, 'selectcolor_on': 0, "bucketpainter_on": 0,
                              "repairbackground_on": 0, "perspective_cut_on": 0, "polygon_ss_on": 0}

        self.old_pen = self.old_eraser = self.old_brush = self.old_backgrounderaser = [-2, -2]
        self.left_button_push = False

    def init_slabel_ui(self):  # 初始化界面的参数

        # self.shower.hide()

        self.setToolTip("左键框选，右键返回")

    def search_in_which_screen(self):
        mousepos = Controller().position
        screens = QApplication.screens()
        secondscreen = QApplication.primaryScreen()
        for i in screens:
            rect = i.geometry().getRect()
            if mousepos[0] in range(rect[0], rect[0] + rect[2]) and mousepos[1] in range(rect[1], rect[1] + rect[3]):
                secondscreen = i
                break
        print("t", self.x(), QApplication.desktop().width(), QApplication.primaryScreen().geometry(),
              secondscreen.geometry(), mousepos)
        return secondscreen

    def screen_shot(self, pix=None, mode="screenshot"):
        self.sshoting = True
        time.process_time()
        pixRat = QWindow().devicePixelRatio()
        if type(pix) is QPixmap:
            get_pix = pix
            self.init_parameters()
        else:
            self.setup(mode)  # 初始化截屏
            if QApplication.desktop().screenCount() > 1:
                sscreen = self.search_in_which_screen()
            else:
                sscreen = QApplication.primaryScreen()
            get_pix = sscreen.grabWindow(0)
            get_pix.setDevicePixelRatio(pixRat)
        pixmap = QPixmap(get_pix.width(), get_pix.height())
        pixmap.setDevicePixelRatio(pixRat)
        pixmap.fill(Qt.transparent)  # 填充透明色,不然没有透明通道

        painter = QPainter(pixmap)
        painter.drawPixmap(0, 0, get_pix)
        painter.end()  # 一定要end
        self.originalPix = pixmap.copy()
        self.setPixmap(pixmap)
        self.mask.setGeometry(0, 0, get_pix.width(), get_pix.height())
        self.mask.show()

        self.paintlayer.setGeometry(0, 0, get_pix.width(), get_pix.height())
        self.paintlayer.setPixmap(QPixmap(get_pix.width(), get_pix.height()))
        self.paintlayer.pixmap().fill(Qt.transparent)  # 重点,不然不透明
        self.paintlayer.show()

        self.setWindowOpacity(1)
        self.showFullScreen()
        if type(pix) is not QPixmap:
            self.backup_ssid = 0
            self.backup_pic_list = [self.originalPix.copy()]

        self.init_ss_thread_fun(get_pix)
        # # 以下设置样式
        self.setStyleSheet("QPushButton{color:black;background-color:rgb(239,239,239);padding:1px 4px;}"
                           "QPushButton:hover{color:green;background-color:rgb(200,200,100);}"
                           "QGroupBox{border:none;}")

        self.setFocus()
        self.setMouseTracking(True)
        self.activateWindow()
        self.raise_()
        self.update()

        QApplication.processEvents()

    def init_ss_thread_fun(self, get_pix):  # 后台初始化截屏线程,用于寻找所有智能选区

        self.x0 = self.y0 = 0
        self.x1 = QApplication.desktop().width()
        self.y1 = QApplication.desktop().height()
        self.mouse_posx = self.mouse_posy = -150
        self.qimg = get_pix.toImage()
        temp_shape = (self.qimg.height(), self.qimg.width(), 4)
        ptr = self.qimg.bits()
        ptr.setsize(self.qimg.byteCount())
        result = array(ptr, dtype=uint8).reshape(temp_shape)[..., :3]
        self.finder.img = result
        self.finder.find_contours_setup()
        QApplication.processEvents()

    def choice(self):  # 选区完毕后显示选择按钮的函数
        self.choicing = True
        print(f'选区区域: {self.x0, self.y0, self.x1, self.y1}')

    def mouseDoubleClickEvent(self, e):  # 双击
        if e.button() == Qt.LeftButton:
            print("左键双击")

    # 鼠标点击事件
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:  # 按下了左键
            self.left_button_push = True
            if 1 in self.painter_tools.values():  # 如果有绘图工具打开了,说明正在绘图
                pass
            else:  # 否则说明正在选区或移动选区
                r = 0
                x0 = min(self.x0, self.x1)
                x1 = max(self.x0, self.x1)
                y0 = min(self.y0, self.y1)
                y1 = max(self.y0, self.y1)
                my = (y1 + y0) // 2
                mx = (x1 + x0) // 2
                # 以下为判断点击在哪里
                if not self.finding_rect and (self.x0 - 8 < event.x() < self.x0 + 8) and (
                        my - 8 < event.y() < my + 8 or y0 - 8 < event.y() < y0 + 8 or y1 - 8 < event.y() < y1 + 8):
                    self.move_x0 = True
                    r = 1

                elif not self.finding_rect and (self.x1 - 8 < event.x() < self.x1 + 8) and (
                        my - 8 < event.y() < my + 8 or y0 - 8 < event.y() < y0 + 8 or y1 - 8 < event.y() < y1 + 8):
                    self.move_x1 = True
                    r = 1
                    # print('x1')

                elif not self.finding_rect and (self.y0 - 8 < event.y() < self.y0 + 8) and (
                        mx - 8 < event.x() < mx + 8 or x0 - 8 < event.x() < x0 + 8 or x1 - 8 < event.x() < x1 + 8):
                    self.move_y0 = True
                    print('y0')
                elif not self.finding_rect and self.y1 - 8 < event.y() < self.y1 + 8 and (
                        mx - 8 < event.x() < mx + 8 or x0 - 8 < event.x() < x0 + 8 or x1 - 8 < event.x() < x1 + 8):
                    self.move_y1 = True

                elif (x0 + 8 < event.x() < x1 - 8) and (
                        y0 + 8 < event.y() < y1 - 8) and not self.finding_rect:
                    # if not self.finding_rect:
                    self.move_rect = True
                    self.setCursor(Qt.SizeAllCursor)
                    self.bx = abs(max(self.x1, self.x0) - event.x())
                    self.by = abs(max(self.y1, self.y0) - event.y())
                else:
                    self.NpainterNmoveFlag = True  # 没有绘图没有移动还按下了左键,说明正在选区,标志变量
                    # if self.finding_rect:
                    #     self.rx0 = event.x()
                    #     self.ry0 = event.y()
                    # else:
                    self.rx0 = event.x()  # 记录下点击位置
                    self.ry0 = event.y()
                    if self.x1 == -50:
                        self.x1 = event.x()
                        self.y1 = event.y()

                    # print('re')
                if r:  # 判断是否点击在了对角线上
                    if (self.y0 - 8 < event.y() < self.y0 + 8) and (
                            x0 - 8 < event.x() < x1 + 8):
                        self.move_y0 = True
                        # print('y0')
                    elif self.y1 - 8 < event.y() < self.y1 + 8 and (
                            x0 - 8 < event.x() < x1 + 8):
                        self.move_y1 = True
                        # print('y1')
            if self.finding_rect:
                self.finding_rect = False
            self.update()

    # 鼠标释放事件
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.left_button_push = False
            self.setCursor(Qt.ArrowCursor)
            self.NpainterNmoveFlag = False  # 选区结束标志置零
            self.move_rect = self.move_y0 = self.move_x0 = self.move_x1 = self.move_y1 = False
            if not self.painter_tools["perspective_cut_on"] and not self.painter_tools["polygon_ss_on"]:
                print('re')
                self.choice()

        elif event.button() == Qt.RightButton:  # 右键
            self.setCursor(Qt.ArrowCursor)

            if self.choicing:  # 退出选定的选区
                # self.botton_box.hide()
                self.choicing = False
                self.finding_rect = True
                # self.shower.hide()
                self.x0 = self.y0 = self.x1 = self.y1 = -50
            else:  # 退出截屏
                try:
                    if not QSettings('Fandes', 'jamtools').value("S_SIMPLE_MODE", False, bool):
                        self.parent.show()

                    self.parent.bdocr = False
                except:
                    print(sys.exc_info(), 2051)
                self.clear_and_hide()
            self.update()

    # 鼠标移动事件
    def mouseMoveEvent(self, event):
        if self.isVisible():
            self.mouse_posx = event.x()  # 先储存起鼠标位置,用于画笔等的绘图计算
            self.mouse_posy = event.y()
            if self.finding_rect and self.smartcursor_on:  # 如果允许智能选取并且在选选区步骤
                self.x0, self.y0, self.x1, self.y1 = self.finder.find_targetrect((self.mouse_posx, self.mouse_posy))
            else:  # 不在绘画
                minx = min(self.x0, self.x1)
                maxx = max(self.x0, self.x1)
                miny = min(self.y0, self.y1)
                maxy = max(self.y0, self.y1)  # 以上取选区的最小值和最大值
                my = (maxy + miny) // 2
                mx = (maxx + minx) // 2  # 取中间值
                if ((minx - 8 < event.x() < minx + 8) and (miny - 8 < event.y() < miny + 8)) or \
                        ((maxx - 8 < event.x() < maxx + 8) and (maxy - 8 < event.y() < maxy + 8)):
                    self.setCursor(Qt.SizeFDiagCursor)
                elif ((minx - 8 < event.x() < minx + 8) and (maxy - 8 < event.y() < maxy + 8)) or \
                        ((maxx - 8 < event.x() < maxx + 8) and (miny - 8 < event.y() < miny + 8)):
                    self.setCursor(Qt.SizeBDiagCursor)
                elif (self.x0 - 8 < event.x() < self.x0 + 8) and (
                        my - 8 < event.y() < my + 8 or miny - 8 < event.y() < miny + 8 or maxy - 8 < event.y() < maxy + 8):
                    self.setCursor(Qt.SizeHorCursor)
                elif (self.x1 - 8 < event.x() < self.x1 + 8) and (
                        my - 8 < event.y() < my + 8 or miny - 8 < event.y() < miny + 8 or maxy - 8 < event.y() < maxy + 8):
                    self.setCursor(Qt.SizeHorCursor)
                elif (self.y0 - 8 < event.y() < self.y0 + 8) and (
                        mx - 8 < event.x() < mx + 8 or minx - 8 < event.x() < minx + 8 or maxx - 8 < event.x() < maxx + 8):
                    self.setCursor(Qt.SizeVerCursor)
                elif (self.y1 - 8 < event.y() < self.y1 + 8) and (
                        mx - 8 < event.x() < mx + 8 or minx - 8 < event.x() < minx + 8 or maxx - 8 < event.x() < maxx + 8):
                    self.setCursor(Qt.SizeVerCursor)
                elif (minx + 8 < event.x() < maxx - 8) and (
                        miny + 8 < event.y() < maxy - 8):
                    if self.move_rect:
                        self.setCursor(Qt.SizeAllCursor)
                elif self.move_x1 or self.move_x0 or self.move_y1 or self.move_y0:  # 再次判断防止光标抖动
                    b = (self.x1 - self.x0) * (self.y1 - self.y0) > 0
                    if (self.move_x0 and self.move_y0) or (self.move_x1 and self.move_y1):
                        if b:
                            self.setCursor(Qt.SizeFDiagCursor)
                        else:
                            self.setCursor(Qt.SizeBDiagCursor)
                    elif (self.move_x1 and self.move_y0) or (self.move_x0 and self.move_y1):
                        if b:
                            self.setCursor(Qt.SizeBDiagCursor)
                        else:
                            self.setCursor(Qt.SizeFDiagCursor)
                    elif (self.move_x0 or self.move_x1) and not (self.move_y0 or self.move_y1):
                        self.setCursor(Qt.SizeHorCursor)
                    elif not (self.move_x0 or self.move_x1) and (self.move_y0 or self.move_y1):
                        self.setCursor(Qt.SizeVerCursor)
                    elif self.move_rect:
                        self.setCursor(Qt.SizeAllCursor)
                else:
                    self.setCursor(Qt.ArrowCursor)
                # 以上几个ifelse都是判断鼠标移动的位置和选框的关系然后设定光标形状
                # print(11)
                if self.NpainterNmoveFlag:  # 如果没有在绘图也没在移动(调整)选区,在选区,则不断更新选区的数值
                    # self.sure_btn.hide()
                    # self.roll_ss_btn.hide()
                    self.x1 = event.x()  # 储存当前位置到self.x1下同
                    self.y1 = event.y()
                    self.x0 = self.rx0  # 鼠标按下时记录的坐标,下同
                    self.y0 = self.ry0
                    if self.y1 > self.y0:  # 下面是边界修正,由于选框占用了一个像素,否则有误差
                        self.y1 += 1
                    else:
                        self.y0 += 1
                    if self.x1 > self.x0:
                        self.x1 += 1
                    else:
                        self.x0 += 1
                else:  # 说明在移动或者绘图,不过绘图没有什么处理的,下面是处理移动/拖动选区
                    if self.move_x0:  # 判断拖动标志位,下同
                        self.x0 = event.x()
                    elif self.move_x1:
                        self.x1 = event.x()
                    if self.move_y0:
                        self.y0 = event.y()
                    elif self.move_y1:
                        self.y1 = event.y()
                    elif self.move_rect:  # 拖动选框
                        dx = abs(self.x1 - self.x0)
                        dy = abs(self.y1 - self.y0)
                        if self.x1 > self.x0:
                            self.x1 = event.x() + self.bx
                            self.x0 = self.x1 - dx
                        else:
                            self.x0 = event.x() + self.bx
                            self.x1 = self.x0 - dx

                        if self.y1 > self.y0:
                            self.y1 = event.y() + self.by
                            self.y0 = self.y1 - dy
                        else:
                            self.y0 = event.y() + self.by
                            self.y1 = self.y0 - dy
            self.update()  # 更新界面
        QApplication.processEvents()

    def keyPressEvent(self, e):  # 按键按下,没按一个键触发一次
        super(Slabel, self).keyPressEvent(e)
        if e.key() == Qt.Key_Escape:  # 退出
            self.clear_and_hide()

    def clear_and_hide(self):  # 清理退出
        print("clear and hide")
        self.hide()

    # 绘制事件
    def paintEvent(self, event):  # 绘图函数,每次调用self.update时触发
        super().paintEvent(event)
        if self.on_init:
            print('oninit return')
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    s = Slabel()
    # s.smartcursor_on = False
    s.screen_shot()
    s.show()
    sys.exit(app.exec_())
