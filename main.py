import os
# 指明系统声音驱动，适配win10系统
# pygame.mixer 实际上是基于 SDL（Simple DirectMedia Layer）库工作的。
# SDL 根据 SDL_AUDIODRIVER 环境变量来选择当前系统上可用的音频驱动。
os.environ.setdefault('SDL_AUDIODRIVER', 'directsound')
import sys
import time
import configparser
import pygame
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, QSharedMemory
import winreg


class ConfigManager:

    # config文件不打包到exe，与exe同目录，所以获取exe文件夹路径
    @staticmethod
    def get_base_path():
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.abspath(os.path.dirname(__file__))

    def __init__(self):
        self.config_file = os.path.join(self.get_base_path(), "config.ini")
        self.config = configparser.ConfigParser()
        self.load_config()

    def load_config(self):
        if not os.path.exists(self.config_file):
            self.config['Settings'] = {
                'AutoStart': 'true',
                'StartHour': '7',
                'EndHour': '22',
                'ChimeType': 'westminster'
            }
            self.save_config()
        else:
            self.config.read(self.config_file)

    def save_config(self):
        with open(self.config_file, 'w') as f:
            self.config.write(f)

    def get_bool(self, key, default=False):
        return self.config.getboolean('Settings', key, fallback=default)

    def get_int(self, key, default=0):
        return self.config.getint('Settings', key, fallback=default)

    def get_str(self, key, default=''):
        return self.config.get('Settings', key, fallback=default)

    def set_value(self, key, value):
        self.config.set('Settings', key, str(value))
        self.save_config()


class AutoStartManager:
    REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
    KEY_NAME = "HourlyChime"

    # 此处获取exe文件完整路径，以便加入注册表设置开机自启
    @staticmethod
    def get_exec_path():
        if getattr(sys, 'frozen', False):
            return sys.executable
        return f'"{sys.executable}" "{os.path.abspath(__file__)}"'

    @classmethod
    def is_enabled(cls):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, cls.REG_PATH, 0, winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, cls.KEY_NAME)
                return cls.get_exec_path() in value
        except Exception:
            return False

    @classmethod
    def set_enabled(cls, enable):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, cls.REG_PATH, 0, winreg.KEY_WRITE) as key:
                if enable:
                    winreg.SetValueEx(key, cls.KEY_NAME, 0, winreg.REG_SZ, cls.get_exec_path())
                else:
                    winreg.DeleteValue(key, cls.KEY_NAME)
        except Exception as e:
            print("设置开机启动失败:", e)


class ChimePlayer:
    def __init__(self, chime_file, speech_dir):
        pygame.mixer.init()
        self.chime_file = chime_file
        self.speech_dir = speech_dir

    def play(self, hour=None):
        sounds = []
        if os.path.exists(self.chime_file):
            sounds.append(self.chime_file)
        if hour is not None:
            speech_file = os.path.join(self.speech_dir, f"hourly-speak-{hour:02d}.wav")
            if os.path.exists(speech_file):
                sounds.append(speech_file)

        if sounds:
            self._play_sequence(sounds)
        else:
            print("音频文件不存在")

    # winsound适配win更好（surface），但是存在线程阻塞，可能崩溃，稳定性不如pygame
    '''
    # 放在最前
    import winsound
    
    # 播放音乐函数（基于winsound）
    def _play_sequence(self, sound_files):
        # 播放第一个音频文件
        winsound.PlaySound(sound_files[0], winsound.SND_FILENAME)
    
        # 播放剩下的音频文件
        for sound_file in sound_files[1:]:
            # 等待前一个音频播放完成
            try:
                winsound.PlaySound(sound_file, winsound.SND_FILENAME)
            except Exception as e:
                print(f"播放失败：{sound_file} -> {e}")
    '''

    # 在surface上运行无声音，可能是pygame的依赖问题
    def _play_sequence(self, sound_files):
        pygame.mixer.music.stop()
        pygame.mixer.music.load(sound_files[0])
        pygame.mixer.music.play()

        for sound_file in sound_files[1:]:
            while pygame.mixer.music.get_busy():
                pygame.time.wait(100)
            pygame.mixer.music.load(sound_file)
            pygame.mixer.music.play()


class TrayApp:

    # 此处需要获取打包后的临时缓存文件夹_MEIxxxxx目录，不是exe文件夹
    @staticmethod
    def get_base_path():
        if hasattr(sys, '_MEIPASS'):
            return sys._MEIPASS
        return os.path.abspath(os.path.dirname(__file__))

    def __init__(self):
        self.app = QApplication(sys.argv)

        # 单实例检查
        self.shared_memory = QSharedMemory("TrayAppInstanceLock")
        if not self.check_single_instance():
            print("程序已经在运行！")
            sys.exit()

        self.base_path = self.get_base_path()

        # 托盘图标和菜单图标
        self.icon_path = self.resource_path('src-icon/tray_icon.ico')
        self.icon_test = self.resource_path("src-icon/menu_test_play.ico")
        self.icon_exit = self.resource_path("src-icon/menu_exit.ico")
        self.icon_autostart_on = self.resource_path("src-icon/menu_auto_start_on.ico")  # 开机自启开启时图标
        self.icon_autostart_off = self.resource_path("src-icon/menu_auto_start_off.ico")  # 开机自启关闭时图标
        self.icon_chime = self.resource_path("src-icon/menu_chime.ico")

        # 声音文件
        self.chime_file_west = self.resource_path('src-sound/chime-westminster.wav')
        self.chime_file_norm = self.resource_path('src-sound/chime-normal.wav')
        self.speech_dir = self.resource_path('src-sound')

        # 配置管理
        self.config = ConfigManager()

        # 初始化当前钟声路径
        self.update_chime_file()

        # 整点报时管理
        self.chimer = ChimePlayer(self.chime_file, self.speech_dir)
        self.last_chimed_hour = -1

        # 初始化托盘
        self.init_tray()
        self.start_timer()

    # 构造资源路径
    def resource_path(self, relative_path):
        return os.path.join(self.base_path, relative_path)

    # 单实例检查
    def check_single_instance(self):
        if self.shared_memory.attach():
            return False  # 已有实例在运行
        else:
            return self.shared_memory.create(1)  # 没有实例运行，创建新的共享内存块

    # 初始化托盘菜单
    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(QIcon(self.icon_path), self.app)
        self.tray_icon.setToolTip("整点报时")

        menu = QMenu()

        # tray菜单: 开机自启
        self.action_autostart = QAction(self.app)
        self.update_autostart_icon()
        self.action_autostart.triggered.connect(self.toggle_autostart)
        menu.addAction(self.action_autostart)
        menu.addSeparator()

        # tray菜单: 钟声选择
        self.action_chime_west = QAction(QIcon(self.icon_chime), "西敏寺钟声", self.app)
        self.action_chime_norm = QAction(QIcon(self.icon_chime), "普通钟声", self.app)

        self.action_chime_west.setCheckable(True)
        self.action_chime_norm.setCheckable(True)

        self.action_chime_west.triggered.connect(lambda: self.set_chime_type("westminster"))
        self.action_chime_norm.triggered.connect(lambda: self.set_chime_type("normal"))

        menu.addAction(self.action_chime_west)
        menu.addAction(self.action_chime_norm)

        self.update_chime_menu()

        menu.addSeparator()

        # tray菜单: 测试报时
        self.action_test = QAction(QIcon(self.icon_test), "测试报时", self.app)
        self.action_test.triggered.connect(lambda: self.chimer.play(hour=time.localtime().tm_hour))
        menu.addAction(self.action_test)
        menu.addSeparator()

        # tray菜单: 退出
        self.action_exit = QAction(QIcon(self.icon_exit), "退出", self.app)
        self.action_exit.triggered.connect(self.app.quit)
        menu.addAction(self.action_exit)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    # 钟声选择功能-写配置文件，切换相应的钟声
    def set_chime_type(self, chime_type):
        self.config.set_value("chimetype", chime_type)
        self.update_chime_file()
        self.chimer.chime_file = self.chime_file
        self.update_chime_menu()

    # 钟声选择功能-读配置文件，切换相应的钟声
    def update_chime_file(self):
        chime_type = self.config.get_str('chimetype', 'westminster')
        if chime_type == 'normal':
            self.chime_file = self.chime_file_norm
        else:
            self.chime_file = self.chime_file_west

    # 钟声选择功能-读配置文件，设置菜单状态
    def update_chime_menu(self):
        current_type = self.config.get_str('chimetype', 'westminster')
        if current_type == 'westminster':
            # 设置菜单check框
            self.action_chime_west.setChecked(True)
            self.action_chime_norm.setChecked(False)
            # 设置菜单文字
            self.action_chime_west.setText("西敏寺钟声（使用中）")
            self.action_chime_norm.setText("普通钟声")
        else:
            # 设置菜单check框
            self.action_chime_west.setChecked(False)
            self.action_chime_norm.setChecked(True)
            # 设置菜单文字
            self.action_chime_west.setText("西敏寺钟声")
            self.action_chime_norm.setText("普通钟声（使用中）")

    # 开启自启功能-写注册表实现开机自启
    def toggle_autostart(self):
        current_enabled = AutoStartManager.is_enabled()
        new_enabled = not current_enabled
        AutoStartManager.set_enabled(new_enabled)
        self.config.set_value("AutoStart", str(new_enabled).lower())
        self.update_autostart_icon()

    # 开启自启功能-UI处理，开机自己菜单项图标切换
    def update_autostart_icon(self):
        enabled = AutoStartManager.is_enabled()
        icon = QIcon(self.icon_autostart_on if enabled else self.icon_autostart_off)
        text = "开机自启（已开启）" if enabled else "开机自启（未开启）"
        self.action_autostart.setIcon(icon)
        self.action_autostart.setText(text)

    # 报时核心功能-定时器启动，每30秒检查一次
    def start_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_hourly_chime)
        self.timer.start(1000 * 30)

    # 报时核心功能-判断当前小时是否已整点，并在指定时间段内播放
    def check_hourly_chime(self):
        now = time.localtime()
        hour = now.tm_hour
        minute = now.tm_min
        if minute == 0 and hour != self.last_chimed_hour:
            start_hour = self.config.get_int("StartHour", 7)
            end_hour = self.config.get_int("EndHour", 22)
            if start_hour <= hour <= end_hour:
                self.chimer.play(hour=hour)
                self.last_chimed_hour = hour

    # 按照1秒检测，功耗略高一些
    '''
    def start_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_hourly_chime)
        self.timer.start(1000 * 30)
    
    def check_hourly_chime(self):
        now = time.localtime()
        if start_hour <= now.tm_hour <= end_hour and now.tm_min == 0 and now.tm_sec == 0:
        if now.tm_hour != self.last_chimed_hour:
            self.last_chimed_hour = now.tm_hour
            self.chimer.play(hour=now.tm_hour)
    '''

    def run(self):
        sys.exit(self.app.exec_())


if __name__ == '__main__':
    TrayApp().run()
