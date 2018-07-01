import tkinter as tk
from tkinter import ttk


class SampleApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self)

        def saveConfig():
            
            config.set('DEFAULT', 'Param1', e1.get())
            config.set('DEFAULT', 'Blah7', e2.get())
            config.set('DEFAULT', 'Param3', e3.get())
            config.write(open('config.ini','w'))
        

        # create the Entry textboxes
        e1 = tk.Entry(self); e1.grid(row=1,column=1,sticky=tk.W)
        e2 = tk.Entry(self); e2.grid(row=2,column=1,sticky=tk.W)
        e3 = tk.Entry(self); e3.grid(row=3,column=1,sticky=tk.W)
        button = tk.Button(self, text="QUIT", command=saveConfig); button.grid(row=4,column=1,sticky=tk.W)

        

        entries = []
        options = []

        entries.append(e1); entries[-1].grid(row=1,column=1,sticky=tk.W)
        options.append("Param1")

        entries.append(e2); entries[-1].grid(row=2,column=1,sticky=tk.W)
        options.append("Blah7")

        entries.append(e3); entries[-1].grid(row=3,column=1,sticky=tk.W)
        options.append("Param3")

        # load

        import configparser
        global config 
        config = configparser.ConfigParser()
        config.read('config.ini')

        for index, e in enumerate(entries):
            e.insert(0, config.get("DEFAULT", options[index]) )

        # save

        for index, e in enumerate(entries):
            config.set("DEFAULT", options[index], e.get())

        config.write(open('config.ini','w'))
            
        
        
       

app = SampleApp()
app.mainloop()

