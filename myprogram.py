import time
import threading

import tkinter
from tkinter import ttk


class GUI(object):

    def __init__(self):
        self.root = tkinter.Tk()

        self.progbar = ttk.Progressbar(self.root)
        self.progbar.config(maximum=10, mode='determinate')
        self.progbar.pack()
        self.i = 0
        self.b_start = ttk.Button(self.root, text='Start')
        self.b_start['command'] = self.start_thread
        self.b_start.pack()

    def start_thread(self):
        self.b_start['state'] = 'disable'
        self.work_thread = threading.Thread(target=work)
        self.work_thread.start()
        self.root.after(50, self.check_thread)
        self.root.after(50, self.update)

    def check_thread(self):
        if self.work_thread.is_alive():
            self.root.after(50, self.check_thread)
        else:
            self.root.destroy()        


    def update(self):
        #Updates the progressbar
        self.progbar["value"] = self.i
        if self.work_thread.is_alive():
            self.root.after(50, self.update)#method is called all 50ms

gui = GUI()

def work():
    #Do your work :D
    print("Hello" * 5000)
    for i in range(11):
        gui.i = i
        time.sleep(0.1)