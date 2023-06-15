import tkinter as tk
import time
import tkinter
import tkinter.filedialog
import threading
import pandas as pd
import pygame   # pip
'''
鼠标点击tkinter窗口任意位置进行拖动
'''
data = pd.read_excel('file/massage.xlsx')

show = data['是否显示'].values[0]
mess = data['消息'].values[0]
class uGUIHandler():
    def __init__(self):
        self.window = tk.Tk()

        self.x, self.y = 0, 0
        self.window_size = '300x150'
        self.window.attributes('-alpha',0.4)
        # ====== 窗口配置 ======
        # 设置窗口标题
        self.window.title('音乐播放器')
        self.window.resizable(False,False)  # 不能拉伸
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
                    print(self.netxMusic)
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
        
            if pause_resume.get()=='暂停' or pause_resume.get()=='继续':
                # global playing
                self.playing=False
                buttonStop['state'] = 'disable'
                musicName.set('暂时没有播放音乐...')
                pygame.mixer.music.stop()
                pause_resume.set('播放')
        
        
        
        
        
        
        
        
        def closeWindow():
            """
            关闭窗口
            :return:
            """
            # 修改变量，结束线程中的循环
        
            global playing
        
            playing = False
        
            time.sleep(0.3)
        
            try:
        
                # 停止播放，如果已停止，
        
                # 再次停止时会抛出异常，所以放在异常处理结构中
        
                pygame.mixer.music.stop()
        
                pygame.mixer.quit()
        
            except:
        
                pass
        
            self.window.destroy()
        
        
        def control_voice(value):
            """
            声音控制
            :param value: 0.0-1.0
            :return:
            """
            # print(type(value))
            value = int(value)
            value=value/100
            # print(value)
            pygame.mixer.music.set_volume(value)
        
        
        
        # 窗口关闭
        self.window.protocol('WM_DELETE_WINDOW', closeWindow)
        
        # 布局
        
        
        # 播放按钮
        pause_resume = tkinter.StringVar(self.window,value='播放')
        buttonPlay = tkinter.Button(self.window,textvariable=pause_resume,command=buttonPlayClick)
        buttonPlay.place(x=155,y=10,width=50,height=20)
        
        
        
        
        # 停止按钮
        buttonStop = tkinter.Button(self.window, text='停止',command=buttonStopClick,state='disabled')
        buttonStop.place(x=95, y=10, width=50, height=20)
        
        
        # 标签
        musicName = tkinter.StringVar(self.window, value='暂时没有播放音乐...')
        labelName = tkinter.Label(self.window, textvariable=musicName)
        labelName.place(x=10, y=30, width=260, height=20)
        
        # 音量控制
        # HORIZONTAL表示为水平放置，默认为竖直,竖直为vertical
        s = tkinter.Scale(self.window, label='音量', from_=0,
                          to=100,
                          tickinterval=20, # 刻度尺
                          orient = tkinter.HORIZONTAL, # 方向(水平) VERTICAL 竖版，HORIZONTAL 横版
                          length=500, # 长度
                          resolution=1,# 精度
                          command=control_voice)
        s.set(80)
        s.place(x=50, y=50, width=200)
        buttonPlayClick()
        if show:
            def message_box():
                print(show,mess)
                showMess = tk.Toplevel()
                showMess.attributes('-alpha',0.9)
                showMess.title('消息')
                showMess.wm_attributes('-topmost', -1)
                showMess.geometry("550x600+100+20")
                text = tk.Text(showMess)
                scroll = tk.Scrollbar(showMess)
                # 放到窗口的右侧, 填充Y竖直方向
                scroll.pack(side=tk.RIGHT,fill=tk.Y)
                # 两个控件关联
                scroll.config(command=text.yview)
                # noinspection PyTypeChecker
                text.config(height=34,font=("宋体",13))
                text.pack(pady=5)
                text.insert("end", '%s'%(mess))
                showMess.mainloop()
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

