import protoapi
import protogui
import sys
import json
import os

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
from pymouse import PyMouse
from pykeyboard import PyKeyboard
from win32api import GetSystemMetrics


class clicker_base(object):
    """仅包含左键单击位置的类。其他属性的基类。
    """

    def __init__(self,
                 setting_dict: dict,
                 mouse: PyMouse = PyMouse(),
                 keyboard: PyKeyboard = PyKeyboard(),
                 delay=0.1):
        self.position = setting_dict['position']
        self.delay = delay
        self.mouse = mouse
        self.keyboard = keyboard

    def execution(self):
        x, y = self.position
        protoapi.click2(self.mouse, x, y, self.delay)


class clicker_clicker(clicker_base):
    """包含复选框的类。
    """

    def __init__(self, setting_dict: dict, status=None, **kwargs):
        super().__init__(setting_dict, **kwargs)
        self.status = status

    def execution(self, status=None):
        if (not status) and (self.status is not None):
            status = self.status
        else:
            assert Exception('Undefine parameters: status')

        x, y = self.position
        if status:
            protoapi.click2(self.mouse, x, y, self.delay)


class clicker_combo(clicker_base):
    """包含下拉框的类。
    """

    def __init__(self, setting_dict: dict, chosen=None, **kwargs):
        super().__init__(setting_dict, **kwargs)
        self.bias = setting_dict['bias']
        self.selections = [str(x) for x in setting_dict['selections']]
        self.chosen = chosen

    def execution(self, chosen=None):
        if (not chosen) and (self.chosen is not None):
            chosen = self.chosen
        else:
            assert Exception('Undefine parameters: chosen')

        x, y = self.position
        x_bias, y_bias = self.bias
        coff = self.selections.index(chosen)
        protoapi.click2(self.mouse, x, y, self.delay)
        x_fin, y_fin = x + coff * x_bias, y + coff * y_bias
        protoapi.click2(self.mouse, x_fin, y_fin, self.delay)


class clicker_sheet(clicker_base):
    """提供于表格填写的类。
    """

    def __init__(self, setting_dict, csv_fname=None, **kwargs):
        super().__init__(setting_dict, **kwargs)
        self.nxtline_button = setting_dict['nxtline_button']
        self.bias = setting_dict['bias']
        self.csv_fname = csv_fname

    def execution(self, csv_fname=None):
        if (not csv_fname) and (self.csv_fname is not None):
            csv_fname = self.csv_fname
        if (csv_fname is not None) and (csv_fname != ''):
            protoapi.solvent_table(csv_fname, self.nxtline_button, self.mouse,
                                   self.keyboard, self.position, self.bias)
        else:
            # we decided not to raise this warning to keep console clean.
            Warning('Undefine parameters: csv_name. Step Ignored.')


class clicker_text(clicker_base):
    """包含可以填写文本的类。
    """

    def __init__(self, setting_dict: dict, text=None, **kwargs):
        super().__init__(setting_dict, **kwargs)
        self.text = text

    def execution(self, text=None):
        if (not text) and (self.text is not None):
            text = self.text

        x, y = self.position
        if text != '':
            print(text, type(text))
            protoapi.click2(self.mouse, x, y, self.delay)
            self.keyboard.type_string(text)
        else:
            Warning('Empty text. Input words would be ignored.')


class waitThreader(QThread):
    trigger_waiting = pyqtSignal()

    def __init__(self, sec=10):
        super().__init__()
        self.sec = sec

    def run(self):
        protoapi.time.sleep(self.sec)
        self.trigger_waiting.emit()

class monitorThreader(QThread):
    trigger_monitor=pyqtSignal()
    def __init__(self,boxs,sec=2):
        """Monitor子线程。
        
        Arguments:
            boxs  -- n个box的dict，包含了n个希望截取的区域
            masks  -- 长度为n的列表，指示需要运行的boxs
            values  -- 用于在对象之间传递OCR结果的list
        """
        super().__init__()
        self.boxs=[value for key,value in boxs.items()]
        self.masks=[1]*len(boxs)
        self.values=['']*len(boxs)
        self.status=False
        self.sec=sec
    def run(self):
        self.values=['']*len(self.boxs)
        while True:
            if self.status:
                for i,mask in enumerate(self.masks):
                    if mask==1:
                        self.values[i]=protoapi.screenshot_ocr(self.boxs[i])
                self.trigger_monitor.emit()
            protoapi.time.sleep(self.sec)

class mainWindow(protogui.Ui_MainWindow, QMainWindow):
    def __init__(self, config, parent=None):
        super(mainWindow, self).__init__(parent)
        self.setupUi(self)
        self.waittrigger=waitThreader(sec=10)
        self.config = config
        self.montrigger=monitorThreader(config['monitor_setting'])
        self.UI_completetion()

    def UI_completetion(self):
        """完善GUI的相关行为。包含初始化数值，信号槽的链接等。
        """

        def add_combo(combo_box, selections):
            combo_box.addItems(map(str, selections['selections']))

        # GUI 优化
        add_combo(self.comboBoxSTem_2, self.config['temper_control']['column'])
        add_combo(self.comboBoxTem_3, self.config['temper_control']['sample'])
        add_combo(self.comboBoxSam,
                  self.config['PDA_detector']['sampling_rate'])
        add_combo(self.comboBoxRes,
                  self.config['PDA_detector']['3d_settings']['resolution'])
        self.setFixedSize(self.width(), self.height())
        self.ifDetector.setChecked(True)
        self.lineRangeL.setEnabled(False)
        self.lineRangeR.setEnabled(False)
        self.lineSheet.setReadOnly(True)
        self.ifTemperCtl.setChecked(True)
        self.ifPressureCtl.setChecked(True)
        self.Stop.setEnabled(False)

        # Signal slot connection
        self.ifDetector.clicked.connect(self.detector_changed)
        self.if3D.clicked.connect(self.if3d_changed)
        self.LabelPath.clicked.connect(self.select_dir)
        self.StartLeft.clicked.connect(self.running_bat)
        self.ConfigCtl.clicked.connect(self.open_cfg)
        self.Start.clicked.connect(self.mstart_clicked)
        self.Stop.clicked.connect(self.mstop_clicked)
        self.ifTemperCtl.clicked.connect(self.tctl_clicked)
        self.ifPressureCtl.clicked.connect(self.pctl_clicked)

        # Outside trigger connection
        self.waittrigger.trigger_waiting.connect(self.bat_exec)
        self.montrigger.trigger_monitor.connect(self.update_value)

    def running_bat(self):
        """
        定义要进行的脚本内容和顺序，并从GUI中加载参数。
        自定义的脚本和行为在此处改写。
        """
        # loading parameters.
        temper_column = self.comboBoxSTem_2.currentText()
        temper_sample = self.comboBoxTem_3.currentText()
        if_detector = self.ifDetector.isChecked()
        sample_rate = self.comboBoxSam.currentText()
        if3d_data = self.if3D.isChecked()
        range_left = self.lineRangeL.text()
        range_right = self.lineRangeR.text()
        sample_interval = self.comboBoxRes.currentText()
        low_pressure = self.lineLowPSI.text()
        high_pressure = self.lineHighPSI.text()
        csv_fpath = self.lineSheet.text()

        # instance every step
        config = self.config
        IMEditor_st = clicker_base(config['IMEditor_st'])
        IMEditor_cls = clicker_base(config['IMEditor_cls'])
        temper_ctl_tab = clicker_base(config['temper_control'])
        temper_column_ctl = clicker_combo(
            config['temper_control']['column'], chosen=temper_column)
        temper_sample_ctl = clicker_combo(
            config['temper_control']['sample'], chosen=temper_sample)
        PDA_detector_tab = clicker_base(config['PDA_detector'])
        if_detector_used = clicker_clicker(
            config['PDA_detector']['if_detector'], if_detector)
        PDA_sr = clicker_combo(
            config['PDA_detector']['sampling_rate'], chosen=sample_rate)
        if_3d = clicker_clicker(
            config['PDA_detector']['3d_settings']['if_3d_data'], if3d_data)
        PDA_3dsr_left = clicker_text(
            config['PDA_detector']['3d_settings']['sampling_range']['left'],
            range_left)
        PDA_3dsr_right = clicker_text(
            config['PDA_detector']['3d_settings']['sampling_range']['right'],
            range_right)
        PDA_3dres = clicker_combo(
            config['PDA_detector']['3d_settings']['resolution'],
            sample_interval)
        sampling_ctl_tab = clicker_base(config['sampling_control'])
        pressure_low = clicker_text(
            config['sampling_control']['pressure_limits']['low'],
            high_pressure)
        pressure_high = clicker_text(
            config['sampling_control']['pressure_limits']['high'],
            low_pressure)
        sampling_sheet = clicker_sheet(
            config['sampling_control']['sampling_sheet'], csv_fname=csv_fpath)

        # build working sequence
        # start
        bat_seq = [IMEditor_st]
        # temperature control
        bat_seq.extend([temper_ctl_tab, temper_column_ctl, temper_sample_ctl])
        # PDA control
        bat_seq.extend([
            PDA_detector_tab,
            if_detector_used,
        ])
        if if_detector:
            bat_seq.append(PDA_sr)
            bat_seq.append(if_3d)
            if if3d_data:
                bat_seq.extend([PDA_3dsr_left, PDA_3dsr_right, PDA_3dres])
        # sampling control
        bat_seq.extend(
            [sampling_ctl_tab, pressure_low, pressure_high, sampling_sheet])
        # finish
        bat_seq.append(IMEditor_cls)

        # execution
        self.StartLeft.setText('请在10秒内将系统焦点转移到待定应用')
        self.bat_seq = bat_seq
        self.waittrigger.start()

    def bat_exec(self):
        for event in self.bat_seq:
            event.execution()
        self.StartLeft.setText('开始填写')

    def detector_changed(self):
        if self.ifDetector.isChecked():
            self.comboBoxSam.setEnabled(True)
            self.if3D.setEnabled(True)
            self.if3d_changed()
            self.comboBoxRes.setEnabled(True)
        else:
            self.comboBoxSam.setEnabled(False)
            self.if3D.setEnabled(False)
            self.if3d_changed()
            self.comboBoxRes.setEnabled(False)

    def if3d_changed(self):
        if self.if3D.isChecked() and self.if3D.isEnabled():
            self.lineRangeL.setEnabled(True)
            self.lineRangeR.setEnabled(True)
        else:
            self.lineRangeL.setEnabled(False)
            self.lineRangeR.setEnabled(False)

    def select_dir(self):
        dir, _ = QFileDialog.getOpenFileName(self, "打开...", "",
                                             "CSV Files (*.csv)")  #起始路径
        if dir != '':
            self.lineSheet.setText(dir)

    def open_cfg(self):
        try:
            if not os.path.exists('setting.json'):
                QMessageBox.critical(self.MainWindow, '丢失了设置文件',
                                     '无法找到文件，该文件可能已被删除。')
            else:
                os.system('explorer setting.json')
        except:
            QMessageBox.critical123(self.MainWindow, '无法打开文件', '可能没有该文件的访问权限。')

    def mstart_clicked(self):
        if self.ifTemperCtl.isChecked() or self.ifPressureCtl.isChecked():
            # enable/disable
            self.Start.setEnabled(False)
            self.Stop.setEnabled(True)
            self.temperSetSam.setEnabled(False)
            self.temperSetCol.setEnabled(False)
            self.pressureSet.setEnabled(False)

            # project check, mask ready
            self.montrigger.masks[0]=self.ifTemperCtl.isChecked()
            self.montrigger.masks[1]=self.ifTemperCtl.isChecked()
            self.montrigger.masks[2]=self.ifPressureCtl.isChecked()
            self.montrigger.status=True
            self.montrigger.start()
        else:
            QMessageBox.information(self, '未勾选任何内容', '没有需要监控的项目')

    def mstop_clicked(self):
        self.montrigger.status=False
        self.montrigger.exit()
        self.Start.setEnabled(True)
        self.Stop.setEnabled(False)
        if self.ifTemperCtl.isChecked():
            self.temperSetSam.setEnabled(True)
            self.temperSetCol.setEnabled(True)
        if self.ifPressureCtl.isChecked():
            self.pressureSet.setEnabled(True)

    def tctl_clicked(self):
        self.temperSetSam.setEnabled(not self.temperSetSam.isEnabled())
        self.temperSetCol.setEnabled(not self.temperSetCol.isEnabled())

    def pctl_clicked(self):
        self.pressureSet.setEnabled(not self.pressureSet.isEnabled())
    
    def update_value(self):
        if self.ifTemperCtl:
            self.valueTempNowSam.setText(self.montrigger.values[0])
            self.valueTempNowCol.setText(self.montrigger.values[1])
            try:
                if (self.temperSetSam.text()!='') and (self.montrigger.values[0]!=''):
                    if int(self.temperSetSam.text())<int(self.montrigger.values[0]):
                        QMessageBox.Warning(self,'温度异常','进样温度异常过高')
                if (self.temperSetCol.text()!='') and (self.montrigger.values[1]!=''):
                    if int(self.temperSetCol.text())<int(self.montrigger.values[1]):
                        QMessageBox.Warning(self,'温度异常','色谱柱温度异常过高')
            except Exception as E:
                print('Value incorrect because ',E,'.Warning Ignored.')
        if self.ifPressureCtl:
            self.valuePressureNow.setText(self.montrigger.values[2])
            try:
                if (self.pressureSet.text()!='') and (self.montrigger.values[2]!=''):
                    if int(self.pressureSet.text())<int(self.montrigger.values[2]):
                        QMessageBox.Warning(self,'压力异常','色谱柱压力异常过高')

            except Exception as E:
                print('Value incorrect because ',E,'.Warning Ignored.')



def resolution_preprocessing(config):
    """根据实际分辨率和给定的理论分辨率进行设计操作的分辨率的放缩。
    
    Arguments:
        config {[type]} -- [description]
    """

    def dict_traveral(tree, x_scale, y_scale):
        for key, value in tree.items():
            if isinstance(value, dict):
                tree[key] = dict_traveral(value, x_scale, y_scale)
            if 'pos' in key:
                x_raw, y_raw = tree[key]
                tree[key] = [int(x_raw * x_scale), int(y_raw * y_scale)]
            if 'box' in key:
                x1, y1, x2, y2 = tree[key]
                tree[key] = [
                    int(x1 * x_scale),
                    int(y1 * y_scale),
                    int(x2 * x_scale),
                    int(y2 * y_scale)
                ]
        return tree

    width_real = GetSystemMetrics(0)
    height_real = GetSystemMetrics(1)
    width_setting, height_setting = config['screen_res']
    if width_real == width_setting and height_real == height_setting:
        return config
    width_scale = width_real / width_setting
    height_scale = height_real / height_setting
    config = dict_traveral(config, width_scale, height_scale)
    return config


# instance
filename_setting = 'setting.json'
config = json.load(open(filename_setting, mode='r'))

m = PyMouse()
k = PyKeyboard()
app = QApplication(sys.argv)
myWin = mainWindow(config)
myWin.show()
sys.exit(app.exec_())