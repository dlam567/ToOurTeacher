import os
import sys
import psutil
import tkinter as tk
import time
import tkinter
from tkinter import messagebox
import tkinter.filedialog
import threading
import pandas as pd
import pygame  # pip
import pystray
import win32api, win32gui, win32con
from PIL import Image

'''
鼠标点击tkinter窗口任意位置进行拖动
'''
try:
    data = pd.read_excel('file/message.xlsx')
    winview = data['窗口大小'].values[0]
    show = data['是否显示'].values[0]
    item = data['显示第几项'].values[0]
    fS = data['字号'].values[0]
    ft = data['字体'].values[0]
    item = int(item)
    mess = data['消息'].values[item]
except:
    
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("提示", "没有检测到message.xlsx, 无法显示消息")
    # root.mainloop()
    root.destroy()
    winview = '1030x600+100+100'
    show = 0


class uGUIHandler:
    def __init__(self):
        self.window = tk.Tk()

        self.x, self.y = 0, 0
        self.window_size = '300x150'
        self.window.attributes('-alpha', 0.4)
        # ====== 窗口配置 ======
        # 设置窗口标题
        self.window.title('音乐播放器')
        self.window.resizable(False, False)  # 不能拉伸
        # 设置窗口置顶
        self.window.wm_attributes('-topmost', -1)
        # 设置隐藏窗口标题栏和任务栏图标
        self.window.overrideredirect(True)
        # 设置窗口大小、位置 长x宽+距离屏幕左边距离x+距离屏幕上边距离y
        self.window.geometry(f"{self.window_size}+700+20")

        # 窗口移动事件
        self.window.bind("<B1-Motion>", self.move)
        # 单击事件
        self.window.bind("<Button-1>", self.get_point)
        # # 双击事件
        # self.window.bind("<Double-Button-1>", self.close)
        self.main()

    # ====== 功能部分 ======
    def main(self):

        # ===== 防止多开 =====
        pids = psutil.pids()  # 获取所有进程PID
        list = []  # 空列表用来存储PID名称
        i = 0  # 计数，程序名称出现的次数
        for pid in pids:  # 遍历所有PID进程
            p = psutil.Process(pid)  # 得到每个PID进程信息
            list.append(p.name())  # 将PID名称放入列表
            s = str(p.name())  # 将PID名称转换成字符串进行判断
            if s == "BestWish.exe":  # “123.exe”你要防多开进程的名称
                i += 1
        if i < 3:  # 如果这个程序名称在程序管理器中出现次数少于两次，执行以下代码
            pass
        else:  # 这个程序名称在任务管理器中出现两次以上，进行程序关掉
            sys.exit()

        def play():

            """
            播放音乐
            :return:
            """

            pygame.mixer.init()
            while self.playing:
                if not pygame.mixer.music.get_busy():
                    # global netxMusic
                    self.music_path = 'file/music.mp3'
                    self.netxMusic = 'music.mp3'
                    # print(self.netxMusic)
                    pygame.mixer.music.load(self.music_path.encode())
                    # 播放
                    pygame.mixer.music.play(1)

                    self.netxMusic = self.netxMusic.split('\\')
                    musicName.set('playing......{0}'.format(self.netxMusic[0]))
                else:
                    time.sleep(0.1)

        def buttonPlayClick():
            """
            点击播放
            :return:
            """

            # 选择要播放的音乐文件夹
            if pause_resume.get() == '播放':
                print(pause_resume.get())

                # global playing

                self.playing = True
                buttonStop['state'] = 'normal'

                # 创建一个线程来播放音乐，当前主线程用来接收用户操作
                t = threading.Thread(target=play)
                t.start()
                pause_resume.set('暂停')


            elif pause_resume.get() == '暂停':
                print(pause_resume.get())
                # pygame.mixer.init()
                pygame.mixer.music.pause()
                self.playing = False
                musicName.set('暂停中')
                pause_resume.set('继续')

            elif pause_resume.get() == '继续':
                print(pause_resume.get())
                # pygame.mixer.init()
                pygame.mixer.music.unpause()
                musicName.set('playing......{0}'.format(self.netxMusic[0]))
                pause_resume.set('暂停')

        def buttonStopClick():
            """
            停止播放
            :return:
            """

            if pause_resume.get() == '暂停' or pause_resume.get() == '继续':
                # global playing
                self.playing = False
                buttonStop['state'] = 'disable'
                musicName.set('暂时没有播放音乐...')
                pygame.mixer.music.stop()
                print('停止')
                pause_resume.set('播放')

        def closeWindow():
            """
            关闭窗口
            :return:
            """
            # 修改变量，结束线程中的循环
            self.playing = False
            time.sleep(1)
            try:
                # 停止播放，如果已停止，
                # 再次停止时会抛出异常，所以放在异常处理结构中
                pygame.mixer.music.stop()
                pygame.mixer.quit()
            except:
                pass
            os.system(("taskkill /F /IM BestWish.exe"))
            os.system(("taskkill /F /IM Python.exe"))
            self.window.destroy()

        def control_voice(value):
            """
            声音控制
            :param value: 0.0-1.0
            :return:
            """
            # print(type(value))
            value = int(value)
            value = value / 100
            # print(value)
            pygame.mixer.music.set_volume(value)

        # ====== 托盘图标与菜单 ======

        if not show:
            self.window.withdraw()
            def winShine():  # 媒体控制菜单窗口闪动三下
                for i in range(3):
                    self.window.deiconify()
                    time.sleep(0.2)
                    self.window.attributes('-alpha', 1)
                    time.sleep(0.2)
                    self.window.attributes('-alpha', 0.4)
                    self.window.deiconify()
            def ic():
                global icon
                image = Image.open("file/cake.ico")
                menu = (
                    pystray.MenuItem('显示', winShine, default=True),
                    pystray.MenuItem('退出', closeWindow))
                icon = pystray.Icon("icon", image, "Best Wishes !", menu)
                icon.run()

            self.ico = threading.Thread(target=ic)
            self.ico.start()

        def message_box():
            global showMess,winShine1
            # print(show,mess)
            showMess = tk.Toplevel()
            showMess.attributes('-alpha', 0.9)
            showMess.title('消息')
            showMess.wm_attributes('-topmost', -1)
            showMess.protocol('WM_DELETE_WINDOW', showMess.withdraw)
            showMess.geometry(winview)
            text = tk.Text(showMess)
            scroll = tk.Scrollbar(showMess)
            # 放到窗口的右侧, 填充Y竖直方向
            scroll.pack(side=tk.RIGHT, fill=tk.Y)
            # 两个控件关联
            scroll.config(command=text.yview)
            # noinspection PyTypeChecker
            text.config(height=34, font=(ft, int(fS)))
            text.pack(pady=27)
            text.insert("end", '%s' % (mess))
            text.config(state='disable')
            def winShine1():  # 媒体控制菜单窗口闪动三下
                showMess.deiconify()
                for i in range(3):
                    self.window.deiconify()
                    time.sleep(0.2)
                    self.window.attributes('-alpha', 1)
                    time.sleep(0.2)
                    self.window.attributes('-alpha', 0.4)
                    self.window.deiconify()
            def ic():
                global icon
                image = Image.open("file/cake.ico")
                menu = (
                    pystray.MenuItem('显示', winShine1, default=True),
                    pystray.Menu.SEPARATOR,  # 在系统托盘菜单中添加分隔线
                    pystray.MenuItem('退出', closeWindow))
                icon = pystray.Icon("icon", image, "Best Wishes !", menu)
                icon.run()

            self.ico = threading.Thread(target=ic)
            self.ico.start()

            showMess.withdraw()
            self.window.withdraw()

            showMess.mainloop()


        # 窗口关闭
        self.window.protocol('WM_DELETE_WINDOW', closeWindow)

        # 布局

        # 播放按钮
        pause_resume = tkinter.StringVar(self.window, value='播放')
        buttonPlay = tkinter.Button(self.window, textvariable=pause_resume, command=buttonPlayClick)
        buttonPlay.place(x=155, y=10, width=50, height=20)

        # 停止按钮
        buttonStop = tkinter.Button(self.window, text='停止', command=buttonStopClick, state='disabled')
        buttonStop.place(x=95, y=10, width=50, height=20)

        # 标签
        musicName = tkinter.StringVar(self.window, value='暂时没有播放音乐...')
        labelName = tkinter.Label(self.window, textvariable=musicName)
        labelName.place(x=10, y=30, width=260, height=20)

        # 音量控制
        # HORIZONTAL表示为水平放置，默认为竖直,竖直为vertical
        s = tkinter.Scale(self.window, label='音量', from_=0,
                          to=100,
                          tickinterval=20,  # 刻度尺
                          orient=tkinter.HORIZONTAL,  # 方向(水平) VERTICAL 竖版，HORIZONTAL 横版
                          length=500,  # 长度
                          resolution=1,  # 精度
                          command=control_voice)
        s.set(80)
        s.place(x=50, y=50, width=200)

        # buttonPlayClick()  # Start Play Music

        def hide():

            messagebox.showinfo("提示", "请在关闭此窗口后，\n按下组合快捷键 Win + F10\n以继续下一步操作")

            #第一个参数填0，意味着为当前线程添加热键，第二个99，你也可以用别的数字，自己测试下
            #我注册的热键是win+F10
            win32gui.RegisterHotKey(0,99,win32con.MOD_WIN, win32con.VK_F10)
            flag = True#真真假假实现开关效果
            while 1:
                time.sleep(1)#避免频繁获取暂停一秒
                msg=win32gui.GetMessage(0,0,0)#获得本线程产生的消息，返回值是个列表
                #msg-----[1, (0, 786, 99, 7929864, 24627051, (534, 440))]
                if msg[1][2]==99:#根据下标和热键id确定按下的是我们注册的热键99
                    if flag:
                        self.window.deiconify()
                        buttonPlayClick()  # Start Play Music
                        try:
                            showMess.deiconify()
                        except:
                            pass
                        flag = False
                    else:
                        try:
                            winShine()
                        except:
                            winShine1()


        xxx=threading.Thread(target=hide)
        xxx.setDaemon(1)#设置守护线程，当线程结束，守护线程同时关闭，要不然这个线程会一直运行下去。
        xxx.start()


        if show:
            Mt = threading.Thread(target=message_box())
            Mt.start()



    # ====== 移动窗口事件函数 ======
    def move(self, event):
        """窗口移动事件"""
        new_x = (event.x - self.x) + self.window.winfo_x()
        new_y = (event.y - self.y) + self.window.winfo_y()
        s = f"{self.window_size}+{new_x}+{new_y}"
        self.window.geometry(s)

    def get_point(self, event):
        """获取当前窗口位置并保存"""
        self.x, self.y = event.x, event.y

    def run(self):
        self.window.mainloop()

    def close(self, event):
        self.window.destroy()


if __name__ == "__main__":
    init_window = uGUIHandler()
    init_window.run()
