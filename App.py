from tkinter import *
from tkinter import ttk, VERTICAL, HORIZONTAL, N, S, E, W
from tkinter import messagebox
from tkinter import filedialog
from tkinter import simpledialog
from PIL import ImageTk,Image
import numpy as np
import uuid
import win32gui
import win32con
import win32api
import random
import collections
from tables import *
import os
import matplotlib
import queue
import logging
import signal
import time
import threading
from multiprocessing import Lock, Process
from tkinter.scrolledtext import ScrolledText
import csv
from datetime import datetime
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)
matplotlib.use('TkAgg')

HEIGHT = 720
WIDTH = 1280
font = ("Calibri", 11)

class App(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.temporal_data()

        ## LOGGER CONFIG

        self.logger = logging.getLogger(__name__)  
        self.logger.setLevel(logging.INFO)  

        # Window Config
        
        self.title("Micro-grid SCADA")
        self.geometry(str(WIDTH) + "x" + str(HEIGHT))
        self.resizable(False, False)
        self.center()

        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure('Treeview',
            background='#D3D3D3',
            foreground='white',
            rowheight=25,
            fieldbackground='#D3D3D3')
        self.style.map('Treeview',
            background=[('selected', '#347083')])

        self.database_object = Database(self)
        self.monitoreo_object = Monitoreo(self)

        self.protocol('WM_DELETE_WINDOW', lambda: self.on_closing(''))

        ## Menu Items ##

        self.my_menu = Menu(self)
        self.config(menu=self.my_menu)

        file_menu = Menu(self.my_menu, tearoff=False)
        self.my_menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New File", command=self.database_object.form)
        file_menu.add_command(label="Open File...", command=self.database_object.import_)
        file_menu.add_command(label="Save", command=lambda:self.database_object.export(self.database_object.getFile()))
        file_menu.add_command(label="Save As...", command=self.database_object.export)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=lambda: self.on_closing(''))

        alarm_menu = Menu(self.my_menu, tearoff=False)
        self.my_menu.add_cascade(label='Alarm Manager', menu=alarm_menu)
        alarm_menu.add_command(label='Open', command=lambda: AlarmManager(self))

        # Top Frame

        self.TopFrame = LabelFrame(self, bd=3)
        self.TopFrame.place(relx=0.5, rely=0.0, relwidth=1.0, relheight=0.06, anchor="n")

        # Middle Frame

        self.MidFrame = LabelFrame(self, bd=4, bg="white")
        self.MidFrame.place(relx=0.5, rely=0.06, relwidth=1.0, relheight= 0.69, anchor="n")

        space = 100
        self.grid_X = np.linspace(0, WIDTH * 0.965, int(space * WIDTH/HEIGHT))
        self.grid_Y = np.linspace(0, HEIGHT * 0.63, space)

        ## Dummy Middle Frame

        self.dMidFrame = LabelFrame(self, bd=4, bg="white")
        self.dMidFrame.place(relx=0.5, rely=0.06, relwidth=1.0, relheight= 0.69, anchor="n")

        # Bottom Frame

        self.BotFrame = LabelFrame(self, bd=4, text='Consola', relief="raised", bg="white")
        self.BotFrame.place(relx=0.5, rely=0.75, relwidth=1.0, relheight= 0.25, anchor="n")
        self.Console = ConsoleUi(self.BotFrame, self)

        # Top Elements

        self.FVImage_icon = ImageTk.PhotoImage(Image.open("img/fvG.png").resize((30,30)))
        self.barra_icon = ImageTk.PhotoImage(Image.open("img/barra.png").resize((30,30)))
        self.line_icon = ImageTk.PhotoImage(Image.open("img/line.png").resize((30,30)))
        self.trafo_icon = ImageTk.PhotoImage(Image.open("img/trafo.png").resize((30,30)))
        self.load_icon = ImageTk.PhotoImage(Image.open("img/load.png").resize((30,30)))
        self.play_icon = ImageTk.PhotoImage(Image.open("img/play.png").resize((30,30)))
        self.stop_icon = ImageTk.PhotoImage(Image.open("img/stop.png").resize((30,30)))

        FVImage_r = ImageTk.PhotoImage(Image.open("img/fvG_r.png").resize((50,50)))
        FVImage_u = ImageTk.PhotoImage(Image.open("img/fvG_u.png").resize((50,50)))
        FVImage_l = ImageTk.PhotoImage(Image.open("img/fvG_l.png").resize((50,50)))
        FVImage_d = ImageTk.PhotoImage(Image.open("img/fvG_d.png").resize((50,50)))
        self.FVImage = [FVImage_r, FVImage_u, FVImage_l, FVImage_d]

        barra_v = ImageTk.PhotoImage(Image.open("img/barra_v.png").resize((15,50)))
        barra_h = ImageTk.PhotoImage(Image.open("img/barra_h.png").resize((50,15)))
        self.barra = [barra_v, barra_h, barra_v, barra_h]

        trafo_v = ImageTk.PhotoImage(Image.open("img/trafo_v.png").resize((80,50)))
        trafo_h = ImageTk.PhotoImage(Image.open("img/trafo_h.png").resize((50,80)))
        self.trafo = [trafo_v, trafo_h, trafo_v, trafo_h]

        load_r = ImageTk.PhotoImage(Image.open("img/load.png").resize((40,40)))
        load_u = ImageTk.PhotoImage(Image.open("img/load.png").resize((40,40)))
        load_l = ImageTk.PhotoImage(Image.open("img/load.png").resize((40,40)))
        load_d = ImageTk.PhotoImage(Image.open("img/load.png").resize((40,40)))
        self.load = [load_r, load_u, load_l, load_d]

        Button(self.TopFrame, image=self.FVImage_icon, command=lambda: FormPhotovoltaic(self, None)).pack(side = LEFT)
        Button(self.TopFrame, image=self.barra_icon, command=lambda: FormTerminal(self, None)).pack(side = LEFT)
        Button(self.TopFrame, image=self.line_icon, command=lambda: FormLineSegment(self, None)).pack(side = LEFT)
        Button(self.TopFrame, image=self.trafo_icon, command=lambda: FormTrafo(self, None)).pack(side = LEFT)
        Button(self.TopFrame, image=self.load_icon, command=lambda: FormLoad(self, None)).pack(side = LEFT)
        Button(self.TopFrame, image=self.play_icon, command=self.monitoreo_object.form).pack(side = LEFT)
        Button(self.TopFrame, image=self.stop_icon, command=self.monitoreo_object.stop_loop).pack(side = LEFT)

        self.after(100, self.database_object.form)

    def temporal_data(self):

        self.changes = 0

        self.dataset_fv = {}
        self.dataset_fv2 = {}
        self.dataset_fv3 = {}

        self.all_elements_object = {}
        self.all_elements_name = {}
        self.all_elements_widget = {}

        self.trafos_object = {}
        self.trafos_name = {}
        self.trafos_widget = {}

        self.terminals_object = {}
        self.terminals_name = {}
        self.terminals_widget = {}

        self.loads_object = {}
        self.loads_name = {}
        self.loads_widget = {}

        self.segments_name = {}  
        self.segments_object = {}
        self.segments_widget = {}

        self.alarms_object = {}
        self.alarms_status = {}

        self.measurements_object = {}

        self.database_object = None
        self.monitoreo_object = None
  
    def center(self):
        """
        centers a tkinter window
        :param win: the main window or Toplevel window to center
        """
        self.update_idletasks()
        width = self.winfo_width()
        frm_width = self.winfo_rootx() - self.winfo_x()
        win_width = width + 2 * frm_width
        height = self.winfo_height()
        titlebar_height = self.winfo_rooty() - self.winfo_y()
        win_height = height + titlebar_height + frm_width
        x = self.winfo_screenwidth() // 2 - win_width // 2
        y = self.winfo_screenheight() // 2 - win_height // 2
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        self.deiconify()

    def center_top(self, toplevel):
        toplevel.update_idletasks()
        self.update_idletasks()
        x = self.winfo_x()
        y = self.winfo_y()
        w = toplevel.winfo_width()
        h = toplevel.winfo_height()  
        toplevel.geometry("%dx%d+%d+%d" % (w, h, x + self.winfo_width()/2 - w/2, y + self.winfo_height()/2 - h/2))

    def on_closing(self, event):
        if self.changes == 1:
            result = messagebox.askyesnocancel('Salir', '¿Quieres guardar antes de salir?')
            if result == True:
                self.database_object.export(self.database_object.getFile())
            elif result == False:
                pass
            else:
                return
        self.dMidFrame.place(relx=0.5, rely=0.06, relwidth=1.0, relheight= 0.69, anchor="n")
        self.destroy()

## Consola

class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    It can be used from different threads
    The ConsoleUi class polls this queue to display records in a ScrolledText widget
    """
    # Example from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06
    # (https://stackoverflow.com/questions/13318742/python-logging-to-tkinter-text-widget) is not thread safe!
    # See https://stackoverflow.com/questions/43909849/tkinter-python-crashes-on-new-thread-trying-to-log-on-main-thread

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)

class ConsoleUi:
    """Poll messages from a logging queue and display them in a scrolled text widget"""

    def __init__(self, frame, app):
        self.frame = frame
        # Create a ScrolledText wdiget
        self.scrolled_text = ScrolledText(frame, state='disabled')
        self.scrolled_text.pack(expand='yes', fill='x')
        self.scrolled_text.configure(font='TkFixedFont')
        self.scrolled_text.tag_config('INFO', foreground='black')
        self.scrolled_text.tag_config('DEBUG', foreground='gray')
        self.scrolled_text.tag_config('WARNING', foreground='orange')
        self.scrolled_text.tag_config('ERROR', foreground='red')
        self.scrolled_text.tag_config('CRITICAL', foreground='red', underline=1)
        # Create a logging handler using a queue
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        formatter = logging.Formatter('%(asctime)s: %(message)s')
        self.queue_handler.setFormatter(formatter)
        app.logger.addHandler(self.queue_handler)
        # Start polling messages from the queue
        self.frame.after(100, self.poll_log_queue)

    def display(self, record):
        msg = self.queue_handler.format(record)
        self.scrolled_text.configure(state='normal')
        self.scrolled_text.insert(END, msg + '\n', record.levelname)
        self.scrolled_text.configure(state='disabled')
        # Autoscroll to the bottom
        self.scrolled_text.yview(END)

    def poll_log_queue(self):
        # Check every 100ms if there is a new message in the queue to display
        while True:
            try:
                record = self.log_queue.get(block=False)
            except queue.Empty:
                break
            else:
                self.display(record)
        self.frame.after(100, self.poll_log_queue)

## Clases de los elementos graficos y sus metodos

class FrameObject():
    def __init__(self, window_, name, image, parent=None, x=0, y=0, orientation=0):
        self.frame = Frame(window_, width=100, height=100, bg="yellow", highlightthickness=0)
        if x!=0 or y!=0:
            self.frame.place(x=x, y=y)
        else:
            self.frame.pack()
        hwnd = self.frame.winfo_id()
        colorkey = win32api.RGB(255,255,0) #full black in COLORREF structure
        wnd_exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        new_exstyle = wnd_exstyle | win32con.WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd,win32con.GWL_EXSTYLE,new_exstyle)
        win32gui.SetLayeredWindowAttributes(hwnd,colorkey,255,win32con.LWA_COLORKEY)
        self.space = Label(self.frame, text="", bg="yellow")
        self.space.pack()
        self.orientation = orientation
        self.image = image
        self.figure = Label(self.frame, image=self.image[self.orientation], cursor="hand1", bg="yellow")
        self.figure.pack()
        self.name = Label(self.frame, text=name, bg="white")
        self.name.pack()
        self.label1 = Label(self.frame, text="P: 10 [W]", width=10, bg="white")
        self.label1.pack()
        self.label2 = Label(self.frame, text="Q: 7.5 [W]", width=10, bg="white")
        self.label2.pack()
        self.x = image[self.orientation].width()
        self.y = image[self.orientation].height()

    def rotate(self):
        self.orientation = self.orientation + 1
        if self.orientation == 4: self.orientation = 0
        self.figure.configure(image=self.image[self.orientation])
        self.draw_lines()

    def draw_lines(self):
        if self.parent.lines: 
            for line in self.parent.lines:
                terminal = line.get_other_terminal(self.parent)
                if line.terminal1.oid == self.parent.oid:
                    line.widget.move(self.getX(terminal, line), self.getY(terminal, line), line.terminal2.widget.getX(self.parent, line), line.terminal2.widget.getY(self.parent, line))
                else:
                    line.widget.move(line.terminal1.widget.getX(self.parent, line), line.terminal1.widget.getY(self.parent, line), self.getX(terminal, line), self.getY(terminal, line))

    def getX(self, other_terminal, line):
        if self.parent.type == "T":
            if self.orientation % 2 == 0:
                izq, der = self.parent.pos_lines_x()
                if other_terminal.oid in [item[0] for item in izq]:
                    pos = izq.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()-6-2
                else:
                    pos = der.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()+6-2
            else:
                top, bot  = self.parent.pos_lines_y()
                if other_terminal.oid in [item[0] for item in top]:
                    pos = top.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x*(pos+1)/(len(top)+1) + self.figure.winfo_x()-4
                else:
                    pos = bot.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x*(pos+1)/(len(bot)+1) + self.figure.winfo_x()-4
        elif self.parent.type =="TR":
            if self.orientation % 2 == 0:
                izq, der = self.parent.pos_lines_x()
                if other_terminal.oid in [item[0] for item in izq]:
                    pos = izq.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()-34
                else:
                    pos = der.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()+34
            else:
                top, bot  = self.parent.pos_lines_y()
                if other_terminal.oid in [item[0] for item in top]:
                    pos = top.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x*(pos+1)/(len(top)+1) + self.figure.winfo_x()-4
                else:
                    pos = bot.index((other_terminal.oid, line.oid))
                    self.x_draw = self.frame.winfo_x() + self.x*(pos+1)/(len(bot)+1) + self.figure.winfo_x()-4
        else:
            if self.orientation == 0:
                self.x_draw = self.frame.winfo_x() + self.x + self.figure.winfo_x()-4
            elif self.orientation == 1:
                self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()-4
            elif self.orientation == 2:
                self.x_draw = self.frame.winfo_x() + self.figure.winfo_x()
            else:
                self.x_draw = self.frame.winfo_x() + self.x/2 + self.figure.winfo_x()-4
        return self.x_draw

    def getY(self, other_terminal, line):
        if self.parent.type == "T":
            izq, der = self.parent.pos_lines_x()
            if other_terminal.oid in [item[0] for item in izq]:
                pos = izq.index((other_terminal.oid, line.oid))
                self.y_draw = self.frame.winfo_y() + self.y*(pos+1)/(len(izq)+1) + self.figure.winfo_y()-1.9
            else:
                pos = der.index((other_terminal.oid, line.oid))
                self.y_draw = self.frame.winfo_y() + self.y*(pos+1)/(len(der)+1) + self.figure.winfo_y()-1.9
        elif self.parent.type == "TR":
            izq, der = self.parent.pos_lines_x()
            if other_terminal.oid in [item[0] for item in izq]:
                pos = izq.index((other_terminal.oid, line.oid))
                self.y_draw = self.frame.winfo_y() + self.y*(pos+1)/(len(izq)+1) + self.figure.winfo_y()-1.9
            else:
                pos = der.index((other_terminal.oid, line.oid))
                self.y_draw = self.frame.winfo_y() + self.y*(pos+1)/(len(der)+1) + self.figure.winfo_y()-1.9
        else:
            if self.orientation == 0:
                self.y_draw = self.frame.winfo_y() + self.y/2-1.9 + self.figure.winfo_y()
            elif self.orientation == 1:
                self.y_draw = self.frame.winfo_y() + self.figure.winfo_y()
            elif self.orientation == 2:
                self.y_draw = self.frame.winfo_y() + self.y/2-1.9 + self.figure.winfo_y()
            else:
                self.y_draw = self.frame.winfo_y() + self.y + self.figure.winfo_y()
        return self.y_draw

    def clear(self):
        for widgets in self.frame.winfo_children():
            widgets.destroy()

    def update_labels(self):
        self.label1.configure(text="P: " + str(self.parent.measurement.rvalues[16][1]) + " [kW]")
        self.label2.configure(text="Q: " + str(self.parent.measurement.rvalues[20][1]) + " [kW]")

class LineObject():
    def __init__(self, window_, x1, y1, x2, y2, parent=None): 
        self.offsetX = 20
        self.offsetY = 20
        self.parent = parent
        self.canvas = Canvas(window_, width=abs(x2-x1)+self.offsetX*2, height=abs(y2-y1)+self.offsetY*2, bg="white", bd=0, highlightthickness=0)
        self.transparent(self.canvas)
        self.x1 = self.offsetX
        self.y1 = self.offsetY
        self.x2 = abs(x2-x1) + self.offsetX
        self.y2 = abs(y2-y1) + self.offsetY
        self.x = abs(x2-x1)/2 + self.offsetX/2
        self.y = abs(y2-y1)/2 + self.offsetY/2
        self.text = self.canvas.create_text((self.x, self.y-5), text=self.parent.name, font=font)

        self.canvas.tag_bind(self.text, '<Enter>', lambda event: self.check_hand_enter())
        self.canvas.tag_bind(self.text, '<Leave>', lambda event: self.check_hand_leave())

        if x1 < x2:
            if y1 < y2:
                self.canvas.place(x=x1-self.offsetX, y=y1-self.offsetY)
                self.orient = 1
            else:
                self.canvas.place(x=x1-self.offsetX, y=y2-self.offsetY)
                self.orient = 0
        else:
            if y1 < y2:
                self.canvas.place(x=x2-self.offsetX, y=y1-self.offsetY)
                self.orient = 0
            else:
                self.canvas.place(x=x2-self.offsetX, y=y2-self.offsetY)
                self.orient = 1

        if self.orient == 0:
            self.line1 = self.canvas.create_line(self.x1, self.y2, self.x, self.y2, fill="black", width=2)
            self.line2 = self.canvas.create_line(self.x, self.y2, self.x, self.y1, fill="black", width=2)
            self.line3 = self.canvas.create_line(self.x, self.y1, self.x2, self.y1, fill="black", width=2)
        else:
            self.line1 = self.canvas.create_line(self.x1, self.y1, self.x, self.y1, fill="black", width=2)
            self.line2 = self.canvas.create_line(self.x, self.y1, self.x, self.y2, fill="black", width=2)
            self.line3 = self.canvas.create_line(self.x, self.y2, self.x2, self.y2, fill="black", width=2)

    def transparent(self, widget):
        hwnd = widget.winfo_id()
        colorkey = win32api.RGB(255,255,255) #full white in COLORREF structure
        wnd_exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        new_exstyle = wnd_exstyle | win32con.WS_EX_LAYERED
        win32gui.SetWindowLong(hwnd,win32con.GWL_EXSTYLE,new_exstyle)
        win32gui.SetLayeredWindowAttributes(hwnd,colorkey,255,win32con.LWA_COLORKEY)

    def replace(self, x1, y1, x2, y2):
        self.canvas.config(width=abs(x2-x1)+self.offsetX*2, height=abs(y2-y1)+self.offsetY*2)
        if x1 < x2:
            if y1 < y2:
                self.canvas.place(x=x1-self.offsetX, y=y1-self.offsetY)
                self.orient = 1
            else:
                self.canvas.place(x=x1-self.offsetX, y=y2-self.offsetY)
                self.orient = 0
        else:
            if y1 < y2:
                self.canvas.place(x=x2-self.offsetX, y=y1-self.offsetY)
                self.orient = 0
            else:
                self.canvas.place(x=x2-self.offsetX, y=y2-self.offsetY)
                self.orient = 1

    def move(self, x1, y1, x2, y2):
        self.replace(x1, y1, x2, y2)
        self.x = abs(x2-x1)/2 + self.offsetX/2
        self.y = abs(y2-y1)/2 + self.offsetY/2
        self.x1 = self.offsetX
        self.y1 = self.offsetY
        self.x2 = abs(x2-x1) + self.offsetX
        self.y2 = abs(y2-y1) + self.offsetY
        if self.orient == 0:
            self.canvas.coords(self.line1, self.x1, self.y2, self.x, self.y2)
            self.canvas.coords(self.line2, self.x, self.y2, self.x, self.y1)
            self.canvas.coords(self.line3, self.x, self.y1, self.x2, self.y1)
        else:
            self.canvas.coords(self.line1, self.x1, self.y1, self.x, self.y1)
            self.canvas.coords(self.line2, self.x, self.y1, self.x, self.y2)
            self.canvas.coords(self.line3, self.x, self.y2, self.x2, self.y2)
        self.canvas.coords(self.text, self.x, self.y-5)

    def clear(self):
        self.canvas.delete('all')

    def check_hand_enter(self):
        self.canvas.config(cursor="hand1")

    def check_hand_leave(self):
        self.canvas.config(cursor="")

    def update_labels(self):
        pass

class Measurement():
    def __init__(self, name, elemento, app, ip='', slaveunit=0, registers=[], oid='', decode=False):
        self.name = name
        self.ip = ip
        self.slaveunit = slaveunit
        self.elemento = elemento
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid

        self.variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
            'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
            'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
            'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
            'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']

        self.unidades = ['V', 'V', 'V', 'V', 'V', 'V', 'V', 'V', 'A', 'A', 'A', 'A', 'A', 'W', 'W', 'W', 'W', 'VAr', 'VAr', 'VAr', 'VAr', 'Hz', '']

        if decode == True: registers = self.decode_regs(registers)
        self.update_registers(registers)

    def update_registers(self, registers):
        i = 0
        self.registers = list()
        for variable in self.variables:
            if len(registers[i]) > 0:
                self.registers.append((variable, registers[i]))
            else:
                self.registers.append((variable, ''))
            i += 1

    def delete(self):
        self.app.measurements_object.pop(self.oid)
        del self

    def connect(self):
        from pymodbus.client import ModbusTcpClient
        self.connected = False
        if '.' in self.ip:
            self.client = ModbusTcpClient(self.ip, 502)
            if self.client.connect():
                self.connected = True
                self.app.logger.log(logging.INFO, 'Conexión exitosa de ' + self.elemento.name)
            else:
                self.app.logger.log(logging.INFO, 'Conexión fallida de ' + self.elemento.name)
    
    def code_regs(self):
        aux = ''
        for variable, register in self.registers:
            aux = aux + '/' + register
        return aux

    def decode_regs(self, registers):
        register = registers.split('/')[1:]
        return register

    def read_registers(self):
        self.rvalues = list()
        if self.connected == True:
            for variable, register in self.registers:
                value = 0
                try:
                    st = datetime.now()
                    if register != '': value = self.client.read_holding_registers(int(register)-1, 1, int(self.slaveunit)).registers[0]
                    et = datetime.now()
                    self.rvalues.append((variable, value, st + (et-st)/2))
                except:
                    self.client.close()
                    self.app.logger.log(logging.INFO, 'Conexión fallida de ' + self.elemento.name)
                    self.connected = False
    
    def write_to_file(self, _lock, file):
        with _lock:
            writer = csv.writer(file)
            for variable, value, timestamp in self.rvalues:
                index = self.rvalues.index((variable, value, timestamp))
                variable, register = self.registers[index]
                if register != '': writer.writerow((self.elemento.oid, self.elemento.name, timestamp, variable, value))

    def disconnect(self):
        self.client.close()
        self.connected = False
        self.app.logger.log(logging.INFO, 'Conexión cerrada con ' + self.elemento.name)

class Photovoltaic():
    def __init__(self, name, nominalPower, nominalVoltage, app, terminal=None, widget=None, lines=[], oid='', measurement=None):
        self.name = name
        self.type = "FV"
        self.nominalPower = nominalPower
        self.nominalVoltage = nominalVoltage
        self.widget = widget
        self.terminal = terminal
        self.lines = lines
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid
        self.measurement = measurement

    def delete(self):
        global changes
        changes = 1

        while self.lines:
            self.lines[0].delete(self)
        unbindAll(self.widget.figure)
        self.widget.clear()
        self.widget.frame.destroy()
        self.terminal = None
        self.lines = []
        self.delete_alarms()
        self.measurement.delete()

        self.app.dataset_fv.pop(self.oid)
        self.app.dataset_fv2.pop(self.oid)
        self.app.dataset_fv3.pop(self.widget.frame)
        self.app.all_elements_object.pop(self.oid)
        self.app.all_elements_name.pop(self.oid)
        self.app.all_elements_widget.pop(self.widget.frame)

        del self.widget
        del self

    def actualizar(self):
        self.widget.name.configure(text=self.name)
        self.widget.draw_lines()

        self.app.dataset_fv2[self.oid] = self.name
        self.app.all_elements_name[self.oid] = self.name

    def delete_alarms(self):
        alarms = []
        for value in self.app.alarms_object.values():
            if value.elemento == self:
                alarms.append(value)
        for alarm in alarms:
            self.app.alarms_object.pop(alarm.oid)

class LineSegment():
    def __init__(self, name, nominalVoltage, app, terminal1=None, terminal2=None, widget=None, oid='', measurement=None):
        self.name = name
        self.type = "LT"
        self.nominalVoltage = nominalVoltage
        self.terminal1 = terminal1
        self.terminal2 = terminal2
        self.widget = widget
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid
        self.measurement = measurement

    def delete(self, *element):
        self.app.changes = 1

        terminal1 = self.terminal1
        terminal2 = self.terminal2
        self.terminal1.lines.remove(self)
        self.terminal2.lines.remove(self)
        self.delete_alarms()
        self.measurement.delete()

        self.app.segments_object.pop(self.oid)
        self.app.segments_name.pop(self.oid)
        self.app.segments_widget.pop(self.widget.canvas) 

        terminal1.widget.draw_lines()
        terminal2.widget.draw_lines()
        self.widget.clear()
        del self

    def reset(self):
        terminal1 = self.terminal1
        terminal2 = self.terminal2
        self.terminal1.lines.remove(self)
        self.terminal2.lines.remove(self)
        terminal1.widget.draw_lines()
        terminal2.widget.draw_lines()

    def get_other_terminal(self, terminal):
        if terminal == self.terminal1: return self.terminal2
        return self.terminal1

    def actualizar(self):
        self.widget.canvas.itemconfig(self.widget.text, text=self.name)

    def delete_alarms(self):
        alarms = []
        for value in self.app.alarms_object.values():
            if value.elemento == self:
                alarms.append(value)
        for alarm in alarms:
            self.app.alarms_object.pop(alarm.oid)    

class Terminal():
    def __init__(self, name, nominalVoltage, widget, app, lines=[], oid='', measurement=None):
        self.name = name
        self.type = "T"
        self.nominalVoltage = nominalVoltage
        self.lines = lines
        self.widget = widget
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid
        self.measurement = measurement

    def delete(self):
        self.app.changes = 1
        
        while self.lines:
            self.lines[0].delete(self)
        unbindAll(self.widget.figure)
        self.widget.clear()
        self.widget.frame.destroy()
        self.lines = []
        self.delete_alarms()
        self.measurement.delete()

        self.app.terminals_object.pop(self.oid)
        self.app.terminals_name.pop(self.oid)
        self.app.terminals_widget.pop(self.widget.frame)
        self.app.all_elements_object.pop(self.oid)
        self.app.all_elements_name.pop(self.oid)
        self.app.all_elements_widget.pop(self.widget.frame)

        del self.widget
        del self

    def pos_lines_x(self):
        izq = {}
        der = {}
        for line in self.lines:
            terminal = line.get_other_terminal(self)
            if terminal.widget.frame.winfo_x() < self.widget.frame.winfo_x():
                izq[terminal.oid, line.oid] = terminal.widget.frame.winfo_y()
            else:
                der[terminal.oid, line.oid] = terminal.widget.frame.winfo_y()
        sort_izq = collections.OrderedDict(sorted(izq.items(), key=lambda kv: kv[1]))
        sort_der = collections.OrderedDict(sorted(der.items(), key=lambda kv: kv[1]))
        r_izq = list(sort_izq.keys())
        r_der = list(sort_der.keys())
        return r_izq, r_der

    def pos_lines_y(self):
        top = {}
        bot = {}
        for line in self.lines:
            terminal = line.get_other_terminal(self)
            if terminal.widget.frame.winfo_y() < self.widget.frame.winfo_y():
                top[terminal.oid] = terminal.widget.frame.winfo_x()
            else:
                bot[terminal.oid] = terminal.widget.frame.winfo_x()
        sort_top = collections.OrderedDict(sorted(top.items(), key=lambda kv: kv[1]))
        sort_bot = collections.OrderedDict(sorted(bot.items(), key=lambda kv: kv[1]))
        return list(sort_top.keys()), list(sort_bot.keys())

    def actualizar(self):
        self.widget.name.configure(text=self.name)
        self.widget.draw_lines()

        self.app.terminals_name[self.oid] = self.name
        self.app.all_elements_name[self.oid] = self.name

    def delete_alarms(self):
        alarms = []
        for value in self.app.alarms_object.values():
            if value.elemento == self:
                alarms.append(value)
        for alarm in alarms:
            self.app.alarms_object.pop(alarm.oid)

class Trafo():
    def __init__(self, name, nominalVoltage_h, nominalVoltage_l, app, lines=[], widget=None, oid='', measurement=None):
        self.name = name
        self.type = "TR"
        self.nominalVoltage_h = nominalVoltage_h
        self.nominalVoltage_l = nominalVoltage_l
        self.lines = lines
        self.widget = widget
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid
        self.measurement = measurement

    def delete(self):
        self.app.changes = 1
        
        while self.lines:
            self.lines[0].delete(self)
        unbindAll(self.widget.figure)
        self.widget.clear()
        self.widget.frame.destroy()
        self.lines = []
        self.delete_alarms()
        self.measurement.delete()

        self.app.trafos_object.pop(self.oid)
        self.app.trafos_name.pop(self.oid)
        self.app.trafos_widget.pop(self.widget.frame)
        self.app.all_elements_object.pop(self.oid)
        self.app.all_elements_name.pop(self.oid)
        self.app.all_elements_widget.pop(self.widget.frame)

        del self.widget
        del self

    def pos_lines_x(self):
        izq = {}
        der = {}
        for line in self.lines:
            terminal = line.get_other_terminal(self)
            if terminal.widget.frame.winfo_x() < self.widget.frame.winfo_x():
                izq[terminal.oid, line.oid] = terminal.widget.frame.winfo_y()
            else:
                der[terminal.oid, line.oid] = terminal.widget.frame.winfo_y()
        sort_izq = collections.OrderedDict(sorted(izq.items(), key=lambda kv: kv[1]))
        sort_der = collections.OrderedDict(sorted(der.items(), key=lambda kv: kv[1]))
        r_izq = list(sort_izq.keys())
        r_der = list(sort_der.keys())
        return r_izq, r_der

    def pos_lines_y(self):
        top = {}
        bot = {}
        for line in self.lines:
            terminal = line.get_other_terminal(self)
            if terminal.widget.frame.winfo_y() < self.widget.frame.winfo_y():
                top[terminal.oid] = terminal.widget.frame.winfo_x()
            else:
                bot[terminal.oid] = terminal.widget.frame.winfo_x()
        sort_top = collections.OrderedDict(sorted(top.items(), key=lambda kv: kv[1]))
        sort_bot = collections.OrderedDict(sorted(bot.items(), key=lambda kv: kv[1]))
        return list(sort_top.keys()), list(sort_bot.keys()) 

    def actualizar(self):
        self.widget.name.configure(text=self.name)
        self.widget.draw_lines()

        self.app.trafos_name[self.oid] = self.name
        self.app.all_elements_name[self.oid] = self.name 

    def delete_alarms(self):
        alarms = []
        for value in self.app.alarms_object.values():
            if value.elemento == self:
                alarms.append(value)
        for alarm in alarms:
            self.app.alarms_object.pop(alarm.oid)

class Load():
    def __init__(self, name, nominalVoltage, app, lines=[], widget=None, oid='', measurement=None):
        self.name = name
        self.type = "L"
        self.nominalVoltage = nominalVoltage
        self.lines = lines
        self.widget = widget
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid
        self.measurement = measurement

    def delete(self):
        self.app.changes = 1
        
        while self.lines:
            self.lines[0].delete(self)
        unbindAll(self.widget.figure)
        self.widget.clear()
        self.widget.frame.destroy()
        self.lines = []
        self.delete_alarms()
        self.measurement.delete()

        self.app.loads_object.pop(self.oid)
        self.app.loads_name.pop(self.oid)
        self.app.loads_widget.pop(self.widget.frame)
        self.app.all_elements_object.pop(self.oid)
        self.app.all_elements_name.pop(self.oid)
        self.app.all_elements_widget.pop(self.widget.frame)

        del self.widget
        del self

    def actualizar(self):
        self.widget.name.configure(text=self.name)
        self.widget.draw_lines()

        self.app.loads_name[self.oid] = self.name
        self.app.all_elements_name[self.oid] = self.name

    def delete_alarms(self):
        alarms = []
        for value in self.app.alarms_object.values():
            if value.elemento == self:
                alarms.append(value)
        for alarm in alarms:
            self.app.alarms_object.pop(alarm.oid)

class Alarm():
    def __init__(self, name, elemento, variable, condicion, umbral, prioridad, app, oid=''):
        self.name = name
        self.elemento = elemento
        self.variable = variable
        self.condicion = condicion
        self.umbral = umbral
        self.prioridad = prioridad
        self.app = app
        self.oid = str(uuid.uuid4())
        if oid != '': self.oid = oid

    def delete(self):
        self.app.changes = 1
        
        self.app.alarms_object.pop(self.oid)
        del self

## Formularios y varios para gestionar los elementos graficos

def unbindAll(widget):
    widget.unbind("<Button-1>")
    widget.unbind("<B1-Motion>")
    widget.unbind("<Button-3>")

def config(class_, app):
    if class_.type == "T":
        FormTerminal(app, class_)
    elif class_.type == 'FV':
        FormPhotovoltaic(app, class_)
    elif class_.type == "LT":
        FormLineSegment(app, class_)
    elif class_.type == "TR":
        FormTrafo(app, class_)
    else:
        FormLoad(app, class_)

class FormAlarm():
    def __init__(self, app, class_=None):
        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title('Nueva Alarma')
        self.app = app
        self.class_ = class_

        if self.class_ != None: self.top.title("Nueva alarma para " + self.class_.name)
        self.top.resizable(False, False)

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        ## Datos para Combobox

        self.variables = [
                'Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']

        self.condiciones = ['<', '>']

        self.prioridades = ['Alto', 'Normal', 'Bajo']

        ## Frame que contiene formulario

        self.lFrame = LabelFrame(self.top, text='Configuracion')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        self.name = StringVar(value="Alarma 1")
        self.umbral = StringVar(value='220')

        elements = list(self.app.all_elements_name.values())

        if class_ == None:
            Label(self.lFrame, text="Elemento", font=font, width=30, anchor=W).grid(row=0, column=0, padx=10)
            if len(elements) > 0:
                self.cbElemento = ttk.Combobox(self.lFrame, values=elements, font=font, state='readonly', width=24)
                self.cbElemento.grid(row=0, column=1, columnspan=2, sticky=W, padx=10, pady=3)
            else:
                self.cbElemento = ttk.Combobox(self.lFrame, font=font, state='readonly', width=24)
                self.cbElemento.grid(row=0, column=1, columnspan=2, sticky=W, padx=10, pady=3)
                self.cbElemento.set("Null")

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=1, column=0, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=1, column=1, columnspan=2, padx=10)

        Label(self.lFrame, text="Variable eléctrica:", font=font, width=30, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        self.cbVariables = ttk.Combobox(self.lFrame, values=self.variables, font=font, state='readonly', width=24)
        self.cbVariables.grid(row=2, column=1, columnspan=2, sticky=W, padx=10, pady=3)

        Label(self.lFrame, text="Condición:", font=font, width=30, anchor=W).grid(row=3, column=0, sticky=W, padx=10)
        self.cbCondiciones = ttk.Combobox(self.lFrame, values=self.condiciones, font=font, state='readonly', width=24)
        self.cbCondiciones.grid(row=3, column=1, columnspan=2, sticky=W, padx=10, pady=3)

        Label(self.lFrame, text="Valor umbral:", font=font, width=30, anchor=W).grid(row=4, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.umbral, width=30).grid(row=4, column=1, columnspan=2, padx=10)

        Label(self.lFrame, text="Prioridad:", font=font, width=30, anchor=W).grid(row=5, column=0, sticky=W, padx=10)
        self.cbPrioridades = ttk.Combobox(self.lFrame, values=self.prioridades, font=font, state='readonly', width=24)
        self.cbPrioridades.grid(row=5, column=1, columnspan=2, sticky=W, padx=10, pady=3)

        Label(self.lFrame, text='').grid(row=6,column=1) 

        ## Botones fuera de Frame

        Button(self.top, text="Aceptar", command=lambda: self.crear_alarma(class_), font=font).grid(row=5, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=5, column=9, padx=5, pady=5)

        self.app.center_top(self.top)    

    def crear_alarma(self, class_):
        name_val = self.name.get()
        umbral_val = self.umbral.get()
        variable_val = self.cbVariables.get()
        condicion_val = self.cbCondiciones.get() 
        prioridad_val = self.cbPrioridades.get()
        if self.class_ == None and self.cbElemento.get() != '':
            class_ = self.app.all_elements_object[list(self.app.all_elements_object.keys())[self.cbElemento.current()]]

        if name_val == '' or umbral_val == '' or variable_val == '' or condicion_val == '' or prioridad_val == '' or class_ == None:
            messagebox.showerror(title='Error', message='Faltan datos', parent=self.top)
            return

        self.top.destroy()

        new_alarm = Alarm(name=name_val, elemento=class_, variable=variable_val, condicion=condicion_val, umbral=umbral_val, prioridad=prioridad_val, app=self.app)

        self.app.alarms_object[new_alarm.oid] = new_alarm
        self.app.alarms_status[new_alarm.oid] = 0

        if self.class_ == None:
            self.app.alarm_manager.update()

        self.app.logger.log(logging.INFO, 'Alarma (' + name_val + ') creada satisfactoriamente para el elemento ' + class_.name)

        self.app.changes = 1

class FormTerminal():
    def __init__(self, app, class_=None):

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title("Nueva Barra")
        if class_ != None: self.top.title('Modificando ' + class_.name)
        self.top.resizable(False, False)
        self.app = app

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        self.class_= class_

        self.configFrame()
        self.ipFrame()

        if class_ == None: Button(self.top, text="Aceptar", command=self.accept, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        if class_ != None: Button(self.top, text="Aceptar", command=self.update, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=4, column=9, padx=5, pady=5)

        app.center_top(self.top)

    def configFrame(self):
        self.lFrame = LabelFrame(self.top, text='Configuración general')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name = StringVar(value="Barra 1")
            self.nominalVoltage = StringVar(value="23000")
        else:
            self.name = StringVar(value=self.class_.name)
            self.nominalVoltage = StringVar(value=self.class_.nominalVoltage)

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal [V]:", font=font, width=20, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltage, width=30).grid(row=2, column=1, columnspan=2, padx=10)

    def ipFrame(self):

        variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']    

        ## Info para Config. de IP
        self.ipFrame = LabelFrame(self.top, text='Configuración IP')
        self.ipFrame.grid(row=1, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name_M = StringVar(value='')
            self.ip = StringVar(value="")
            self.slaveunit = IntVar()
        else:
            self.name_M = StringVar(value=self.class_.measurement.name)
            self.ip = StringVar(value=self.class_.measurement.ip)
            self.slaveunit = StringVar(value=self.class_.measurement.slaveunit)    

        Label(self.ipFrame, text="Nombre medidor:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.name_M, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Dirección IP:", font=font, width=25, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.ip, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Slave Unit:", font=font, width=25, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.slaveunit, width=30).grid(row=2, column=1, columnspan=2, padx=10)

        self.registers = []
        count = 0
        for variable in variables:
            if self.class_ == None:
                self.value = StringVar(value='')
            else:
                index = self.class_.measurement.variables.index(variable)
                self.value = StringVar(value=self.class_.measurement.registers[index][1])
            self.registers.append(self.value)
            Label(self.ipFrame, text='Reg. de ' + variable, font=font, width=25, anchor=W).grid(row=3+count, column=0, sticky=W, padx=10)
            Entry(self.ipFrame, textvariable=self.value, width=30).grid(row=3+count, column=1, columnspan=2, padx=10)
            count += 1

    def accept(self):
        name_data = self.name.get()
        nominalVoltage_data = self.nominalVoltage.get()

        self.top.destroy()

        frame = FrameObject(self.app.MidFrame, name_data, self.app.barra)

        bind_p(frame.figure, self.app)

        new = Terminal(name=name_data, nominalVoltage=nominalVoltage_data, widget=frame, lines=[], app=self.app)

        frame.parent = new
        new.widget.draw_lines()

        self.app.terminals_object[new.oid] = new
        self.app.terminals_name[new.oid] = new.name
        self.app.terminals_widget[new.widget.frame] = new.widget

        self.app.all_elements_object[new.oid] = new
        self.app.all_elements_name[new.oid] = new.name
        self.app.all_elements_widget[new.widget.frame] = new.widget

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        new_measure = Measurement(name=name_M_data, elemento=new, app=self.app, ip=ip_data, slaveunit=slaveunit_data, registers=registers_data)
        new.measurement = new_measure    
        self.app.measurements_object[new_measure.oid] = new_measure

        self.app.logger.log(logging.INFO, 'Barra (' + name_data + ') creada satisfactoriamente')

        self.app.changes = 1

    def update(self):
        self.class_.name = self.name.get()
        self.class_.nominalVoltage = self.nominalVoltage.get()
 
        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        self.class_.measurement.name = name_M_data
        self.class_.measurement.ip = ip_data
        self.class_.measurement.slaveunit = slaveunit_data
        self.class_.measurement.update_registers(registers_data)

        self.class_.actualizar()
        self.top.destroy()

        self.app.logger.log(logging.INFO, 'Barra (' + self.name.get() + ') modificado satisfactoriamente')
        self.app.changes = 1

class FormLoad():
    def __init__(self, app, class_=None):

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title("Nueva Carga")
        if class_ != None: self.top.title('Modificando ' + class_.name)
        self.top.resizable(False, False)
        self.app = app

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        self.class_= class_

        self.configFrame()
        self.ipFrame()

        if class_ == None: Button(self.top, text="Aceptar", command=self.accept, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        if class_ != None: Button(self.top, text="Aceptar", command=self.update, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=4, column=9, padx=5, pady=5)

        app.center_top(self.top)

    def configFrame(self):
        self.lFrame = LabelFrame(self.top, text='Configuración general')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name = StringVar(value="Load 1")
            self.nominalVoltage = StringVar(value="23000")
        else:
            self.name = StringVar(value=self.class_.name)
            self.nominalVoltage = StringVar(value=self.class_.nominalVoltage)

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal [V]:", font=font, width=20, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltage, width=30).grid(row=2, column=1, columnspan=2, padx=10)

    def ipFrame(self):

        variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']    

        ## Info para Config. de IP
        self.ipFrame = LabelFrame(self.top, text='Configuración IP')
        self.ipFrame.grid(row=1, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name_M = StringVar(value='')
            self.ip = StringVar(value="")
            self.slaveunit = IntVar()
        else:
            self.name_M = StringVar(value=self.class_.measurement.name)
            self.ip = StringVar(value=self.class_.measurement.ip)
            self.slaveunit = StringVar(value=self.class_.measurement.slaveunit)    

        Label(self.ipFrame, text="Nombre medidor:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.name_M, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Dirección IP:", font=font, width=25, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.ip, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Slave Unit:", font=font, width=25, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.slaveunit, width=30).grid(row=2, column=1, columnspan=2, padx=10)

        self.registers = []
        count = 0
        for variable in variables:
            if self.class_ == None:
                self.value = StringVar(value='')
            else:
                index = self.class_.measurement.variables.index(variable)
                self.value = StringVar(value=self.class_.measurement.registers[index][1])
            self.registers.append(self.value)
            Label(self.ipFrame, text='Reg. de ' + variable, font=font, width=25, anchor=W).grid(row=3+count, column=0, sticky=W, padx=10)
            Entry(self.ipFrame, textvariable=self.value, width=30).grid(row=3+count, column=1, columnspan=2, padx=10)
            count += 1

    def accept(self):
        name_data = self.name.get()
        nominalVoltage_data = self.nominalVoltage.get()

        self.top.destroy()

        frame = FrameObject(self.app.MidFrame, name_data, self.app.load)

        bind_p(frame.figure, self.app)

        new = Load(name=name_data, nominalVoltage=nominalVoltage_data, widget=frame, lines=[], app=self.app)

        frame.parent = new
        new.widget.draw_lines()

        self.app.loads_object[new.oid] = new
        self.app.loads_name[new.oid] = new.name
        self.app.loads_widget[new.widget.frame] = new.widget

        self.app.all_elements_object[new.oid] = new
        self.app.all_elements_name[new.oid] = new.name
        self.app.all_elements_widget[new.widget.frame] = new.widget

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        new_measure = Measurement(name=name_M_data, elemento=new, app=self.app, ip=ip_data, slaveunit=slaveunit_data, registers=registers_data)
        new.measurement = new_measure    
        self.app.measurements_object[new_measure.oid] = new_measure

        self.app.logger.log(logging.INFO, 'Carga (' + name_data + ') creada satisfactoriamente')
        self.app.changes = 1

    def update(self):
        self.class_.name = self.name.get()
        self.class_.nominalVoltage = self.nominalVoltage.get()

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        self.class_.measurement.name = name_M_data
        self.class_.measurement.ip = ip_data
        self.class_.measurement.slaveunit = slaveunit_data
        self.class_.measurement.update_registers(registers_data)

        self.class_.actualizar()
        self.top.destroy()

        self.app.logger.log(logging.INFO, 'Transformador (' + self.name.get() + ') modificado satisfactoriamente')
        self.app.changes = 1

class FormTrafo():
    def __init__(self, app, class_=None):

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title('Nuevo Transformador')
        if class_ != None: self.top.title('Modificando ' + class_.name)
        self.top.resizable(False, False)

        self.app = app

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        self.class_= class_

        self.configFrame()
        self.ipFrame()

        if class_ == None: Button(self.top, text="Aceptar", command=self.accept, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        if class_ != None: Button(self.top, text="Aceptar", command=self.update, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=4, column=9, padx=5, pady=5)

        app.center_top(self.top)

    def configFrame(self):
        self.lFrame = LabelFrame(self.top, text='Configuración general')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name = StringVar(value="Trafo 1")
            self.nominalVoltageH = StringVar(value="23000")
            self.nominalVoltageL = StringVar(value='230')
        else:
            self.name = StringVar(value=self.class_.name)
            self.nominalVoltageH = StringVar(value=self.class_.nominalVoltage_h)
            self.nominalVoltageL = StringVar(value=self.class_.nominalVoltage_l)

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal AT [V]:", font=font, width=20, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltageH, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal BT [V]:", font=font, width=20, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltageL, width=30).grid(row=2, column=1, columnspan=2, padx=10)

    def ipFrame(self):

        variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']    

        ## Info para Config. de IP
        self.ipFrame = LabelFrame(self.top, text='Configuración IP')
        self.ipFrame.grid(row=1, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name_M = StringVar(value='')
            self.ip = StringVar(value="")
            self.slaveunit = IntVar()
        else:
            self.name_M = StringVar(value=self.class_.measurement.name)
            self.ip = StringVar(value=self.class_.measurement.ip)
            self.slaveunit = StringVar(value=self.class_.measurement.slaveunit)    

        Label(self.ipFrame, text="Nombre medidor:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.name_M, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Dirección IP:", font=font, width=25, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.ip, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Slave Unit:", font=font, width=25, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.slaveunit, width=30).grid(row=2, column=1, columnspan=2, padx=10)

        self.registers = []
        count = 0
        for variable in variables:
            if self.class_ == None:
                self.value = StringVar(value='')
            else:
                index = self.class_.measurement.variables.index(variable)
                self.value = StringVar(value=self.class_.measurement.registers[index][1])
            self.registers.append(self.value)
            Label(self.ipFrame, text='Reg. de ' + variable, font=font, width=25, anchor=W).grid(row=3+count, column=0, sticky=W, padx=10)
            Entry(self.ipFrame, textvariable=self.value, width=30).grid(row=3+count, column=1, columnspan=2, padx=10)
            count += 1

    def accept(self):
        name_data = self.name.get()
        nominalVoltageH_data = self.nominalVoltageH.get()
        nominalVoltageL_data = self.nominalVoltageL.get()

        self.top.destroy()

        frame = FrameObject(self.app.MidFrame, name_data, self.app.trafo)

        bind_p(frame.figure, self.app)

        new = Trafo(name=name_data, nominalVoltage_h=nominalVoltageH_data, nominalVoltage_l=nominalVoltageL_data, widget=frame, lines=[], app=self.app)
        print(new.nominalVoltage_h)
        print(new.nominalVoltage_l)
        frame.parent = new
        new.widget.draw_lines()

        self.app.trafos_object[new.oid] = new
        self.app.trafos_name[new.oid] = new.name
        self.app.trafos_widget[new.widget.frame] = new.widget

        self.app.all_elements_object[new.oid] = new
        self.app.all_elements_name[new.oid] = new.name
        self.app.all_elements_widget[new.widget.frame] = new.widget

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        new_measure = Measurement(name=name_M_data, elemento=new, app=self.app, ip=ip_data, slaveunit=slaveunit_data, registers=registers_data)
        new.measurement = new_measure    
        self.app.measurements_object[new_measure.oid] = new_measure

        self.app.logger.log(logging.INFO, 'Transformador (' + name_data + ') creado satisfactoriamente')
        self.app.changes = 1

    def update(self):
        self.class_.name = self.name.get()
        self.class_.nominalVoltage_h = self.nominalVoltageH.get()
        self.class_.nominalVoltage_l = self.nominalVoltageL.get()

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        self.class_.measurement.name = name_M_data
        self.class_.measurement.ip = ip_data
        self.class_.measurement.slaveunit = slaveunit_data
        self.class_.measurement.update_registers(registers_data)

        self.class_.actualizar()
        self.top.destroy()

        self.app.logger.log(logging.INFO, 'Carga (' + self.name.get() + ') modificada satisfactoriamente')
        self.app.changes = 1

class FormPhotovoltaic():
    def __init__(self, app, class_=None):

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title("Nuevo Generador FV")
        if class_ != None: self.top.title('Modificando ' + class_.name)
        self.top.resizable(False, False)
        self.app = app

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        self.class_= class_

        self.configFrame()
        self.ipFrame()

        if class_ == None: Button(self.top, text="Aceptar", command=self.accept, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        if class_ != None: Button(self.top, text="Aceptar", command=self.update, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=4, column=9, padx=5, pady=5)

        app.center_top(self.top)

    def configFrame(self):
        self.lFrame = LabelFrame(self.top, text='Configuración general')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name = StringVar(value="Angamos")
            self.nominalPower = StringVar(value="100")
            self.nominalVoltage = StringVar(value="23000")
        else:
            self.name = StringVar(value=self.class_.name)
            self.nominalPower = StringVar(value=self.class_.nominalPower)
            self.nominalVoltage = StringVar(value=self.class_.nominalVoltage)

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Potencia Nominal [W]:", font=font, width=30, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalPower, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal [V]:", font=font, width=30, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltage, width=30).grid(row=2, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Terminal:", font=font, width=30, anchor=W).grid(row=3, column=0, sticky=W, padx=10)

        if len(self.app.terminals_name) > 0:
            terminals = list(self.app.terminals_name.values())
            self.cbTerminal = ttk.Combobox(self.lFrame, values=terminals, font=font, width=24)
            if self.class_ != None:
                if self.class_.terminal == None: 
                    self.cbTerminal.set('Null')
                else:
                    self.cbTerminal.current(terminals.index(self.class_.terminal.name))
            self.cbTerminal.grid(row=3, column=1, columnspan=2, sticky=W, padx=10, pady=3)
        else:
            self.cbTerminal = ttk.Combobox(self.lFrame, font=font, width=24)
            self.cbTerminal.grid(row=3, column=1, columnspan=2, sticky=W, padx=10, pady=3)
            self.cbTerminal.set("Null")

    def ipFrame(self):

        variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']    

        ## Info para Config. de IP
        self.ipFrame = LabelFrame(self.top, text='Configuración IP')
        self.ipFrame.grid(row=1, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name_M = StringVar(value='')
            self.ip = StringVar(value="")
            self.slaveunit = IntVar()
        else:
            self.name_M = StringVar(value=self.class_.measurement.name)
            self.ip = StringVar(value=self.class_.measurement.ip)
            self.slaveunit = StringVar(value=self.class_.measurement.slaveunit)    

        Label(self.ipFrame, text="Nombre medidor:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.name_M, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Dirección IP:", font=font, width=25, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.ip, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Slave Unit:", font=font, width=25, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.slaveunit, width=30).grid(row=2, column=1, columnspan=2, padx=10)

        self.registers = []
        count = 0
        for variable in variables:
            if self.class_ == None:
                self.value = StringVar(value='')
            else:
                index = self.class_.measurement.variables.index(variable)
                self.value = StringVar(value=self.class_.measurement.registers[index][1])
            self.registers.append(self.value)
            Label(self.ipFrame, text='Reg. de ' + variable, font=font, width=30, anchor=W).grid(row=3+count, column=0, sticky=W, padx=10)
            Entry(self.ipFrame, textvariable=self.value, width=30).grid(row=3+count, column=1, columnspan=2, padx=10)
            count += 1

    def accept(self):
        name_data = self.name.get()
        nominalPower_data = self.nominalPower.get()
        nominalVoltage_data = self.nominalVoltage.get()
        if len(self.app.terminals_name) > 0:
            cmb_object = self.app.terminals_object[list(self.app.terminals_name.keys())[self.cbTerminal.current()]]
            if self.cbTerminal.get() == '': cmb_object = None
        else:
            cmb_object = None

        self.top.destroy()

        frame = FrameObject(self.app.MidFrame, name_data, self.app.FVImage)
        bind_p(frame.figure, self.app)
        new = Photovoltaic(name=name_data, nominalPower=nominalPower_data, nominalVoltage=nominalVoltage_data, widget=frame, lines=[], terminal=cmb_object, app=self.app)

        frame.parent = new
        new.widget.draw_lines()

        self.app.dataset_fv[new.oid] = new
        self.app.dataset_fv2[new.oid] = new.name
        self.app.dataset_fv3[new.widget.frame] = new.widget

        self.app.all_elements_object[new.oid] = new
        self.app.all_elements_name[new.oid] = new.name
        self.app.all_elements_widget[new.widget.frame] = new.widget

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        new_measure = Measurement(name=name_M_data, elemento=new, app=self.app, ip=ip_data, slaveunit=slaveunit_data, registers=registers_data)
        new.measurement = new_measure    
        self.app.measurements_object[new_measure.oid] = new_measure

        self.app.logger.log(logging.INFO, 'Paneles fotovoltaicos (' + name_data + ') creado satisfactoriamente')
        self.app.changes = 1

    def update(self):
        self.class_.name = self.name.get()
        self.class_.nominalVoltage = self.nominalVoltage.get()
        if len(self.app.terminals_name) > 0:
            cmb_object = self.app.terminals_object[list(self.app.terminals_name.keys())[self.cbTerminal.current()]]
            self.class_.terminal = cmb_object
        else:
            cmb_object = None

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        self.class_.measurement.name = name_M_data
        self.class_.measurement.ip = ip_data
        self.class_.measurement.slaveunit = slaveunit_data
        self.class_.measurement.update_registers(registers_data)

        self.class_.actualizar()
        self.top.destroy()

        self.app.logger.log(logging.INFO, 'Paneles fotovoltaicos (' + self.name.get() + ') modificado satisfactoriamente')
        self.app.changes = 1

class FormLineSegment():
    def __init__(self, app, class_=None):

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title('Nueva Linea')
        if class_ != None: self.top.title('Modificando ' + class_.name)
        self.top.resizable(False, False)

        self.app = app

        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()

        self.class_= class_
        self.data_object1 = list()
        self.data_object2 = list()

        self.configFrame()
        self.ipFrame()

        if class_ == None: Button(self.top, text="Aceptar", command=self.accept, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        if class_ != None: Button(self.top, text="Aceptar", command=self.update, font=font).grid(row=4, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=4, column=9, padx=5, pady=5)

        app.center_top(self.top)

    def configFrame(self):
        self.lFrame = LabelFrame(self.top, text='Configuración general')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name = StringVar(value="Default: Elm.1 - Elm.2")
            self.nominalVoltage = StringVar(value="23000")
        else:
            self.name = StringVar(value=self.class_.name)
            self.nominalVoltage = StringVar(value=self.class_.nominalVoltage)

        Label(self.lFrame, text="Nombre:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="Tensión Nominal [kV]:", font=font, width=20, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.lFrame, textvariable=self.nominalVoltage, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.lFrame, text="From:", font=font).grid(row=2, column=0, sticky=W, padx=10)

        elements = list(self.app.all_elements_name.values())

        if len(elements) > 0:
            self.cbElement1 = ttk.Combobox(self.lFrame, values=elements, font=font, width=24)
            self.cbElement1.grid(row=2, column=1, columnspan=2, sticky=W, padx=10, pady=3)
        else:
            self.cbElement1 = ttk.Combobox(self.lFrame, font=font, width=24)
            self.cbElement1.grid(row=2, column=1, columnspan=2, sticky=W, padx=10, pady=3)
            self.cbElement1.set("Null")

        Label(self.lFrame, text="To:", font=font).grid(row=3, column=0, sticky=W, padx=10)

        if len(elements) > 0:
            self.cbElement2 = ttk.Combobox(self.lFrame, values=elements, font=font, width=24)
            self.cbElement2.grid(row=3, column=1, columnspan=2, sticky=W, padx=10, pady=3)
        else:
            self.cbElement2 = ttk.Combobox(self.lFrame, font=font, width=24)
            self.cbElement2.grid(row=3, column=1, columnspan=2, sticky=W, padx=10, pady=3)
            self.cbElement2.set("Null")

    def ipFrame(self):

        variables = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
                'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
                'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
                'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
                'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']    

        ## Info para Config. de IP
        self.ipFrame = LabelFrame(self.top, text='Configuración IP')
        self.ipFrame.grid(row=1, column=0, sticky=W, columnspan=10, padx=10, pady=5)

        if self.class_ == None:
            self.name_M = StringVar(value='')
            self.ip = StringVar(value="")
            self.slaveunit = IntVar()
        else:
            self.name_M = StringVar(value=self.class_.measurement.name)
            self.ip = StringVar(value=self.class_.measurement.ip)
            self.slaveunit = StringVar(value=self.class_.measurement.slaveunit)    

        Label(self.ipFrame, text="Nombre medidor:", font=font, width=30, anchor=W).grid(row=0, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.name_M, width=30).grid(row=0, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Dirección IP:", font=font, width=25, anchor=W).grid(row=1, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.ip, width=30).grid(row=1, column=1, columnspan=2, padx=10)
        Label(self.ipFrame, text="Slave Unit:", font=font, width=25, anchor=W).grid(row=2, column=0, sticky=W, padx=10)
        Entry(self.ipFrame, textvariable=self.slaveunit, width=30).grid(row=2, column=1, columnspan=2, padx=10)

        self.registers = []
        count = 0
        for variable in variables:
            if self.class_ == None:
                self.value = StringVar(value='')
            else:
                index = self.class_.measurement.variables.index(variable)
                self.value = StringVar(value=self.class_.measurement.registers[index][1])
            self.registers.append(self.value)
            Label(self.ipFrame, text='Reg. de ' + variable, font=font, width=25, anchor=W).grid(row=3+count, column=0, sticky=W, padx=10)
            Entry(self.ipFrame, textvariable=self.value, width=30).grid(row=3+count, column=1, columnspan=2, padx=10)
            count += 1

    def accept(self):
        name_data = self.name.get()
        nominalVoltage_data = self.nominalVoltage.get()
        if self.cbElement1.current() != -1 and self.cbElement2.current() != -1:
            cmb_object1 = self.app.all_elements_object[list(self.app.all_elements_object.keys())[self.cbElement1.current()]]
            cmb_object2 = self.app.all_elements_object[list(self.app.all_elements_object.keys())[self.cbElement2.current()]]
        else:
            messagebox.showerror(title='Error', message='Terminales incorrectos', parent=self.top)
            return

        if cmb_object1 == cmb_object2:
            messagebox.showerror(title='Error', message='Terminales iguales', parent=self.top)
            return

        self.top.destroy()

        if name_data == "Default: Elm.1 - Elm.2": name_data = str(cmb_object1.name + " - " + cmb_object2.name)

        new_line = LineSegment(name=name_data, nominalVoltage=nominalVoltage_data, app=self.app, terminal1=cmb_object1, terminal2=cmb_object2) 
        cmb_object1.lines.append(new_line)
        cmb_object2.lines.append(new_line)
        new_widget = LineObject(self.app.MidFrame, cmb_object1.widget.getX(cmb_object2, new_line),
                                            cmb_object1.widget.getY(cmb_object2, new_line),
                                            cmb_object2.widget.getX(cmb_object1, new_line),
                                            cmb_object2.widget.getY(cmb_object1, new_line),
                                            new_line)
        new_line.widget = new_widget

        new_line.terminal1.widget.draw_lines()
        new_line.terminal2.widget.draw_lines()

        new_widget.canvas.tag_bind(new_widget.text, "<Button-3>", lambda e: menu_items(e, self.app))

        self.app.segments_object[new_line.oid] = new_line
        self.app.segments_name[new_line.oid] = new_line.name
        self.app.segments_widget[new_widget.canvas] = new_widget

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        new_measure = Measurement(name=name_M_data, elemento=new_line, app=self.app, ip=ip_data, slaveunit=slaveunit_data, registers=registers_data)
        new_line.measurement = new_measure    
        self.app.measurements_object[new_measure.oid] = new_measure

        self.app.logger.log(logging.INFO, 'Linea (' + name_data + ') creada satisfactoriamente')
        self.app.changes = 1

    def update(self):
        self.class_.name = self.name.get()
        self.class_.nominalVoltage = self.nominalVoltage.get()

        cmb_object1 = self.class_.terminal1
        cmb_object2 = self.class_.terminal2

        if self.cbElement1.get() != '': cmb_object1 = self.app.all_elements_object[list(self.app.all_elements_object.keys())[self.cbElement1.current()]]
        if self.cbElement2.get() != '': cmb_object2 = self.app.all_elements_object[list(self.app.all_elements_object.keys())[self.cbElement2.current()]]

        self.class_.reset()
        self.class_.terminal1 = cmb_object1
        self.class_.terminal2 = cmb_object2
        cmb_object1.lines.append(self.class_)
        cmb_object2.lines.append(self.class_)
        cmb_object1.widget.draw_lines()
        cmb_object2.widget.draw_lines()

        name_M_data = self.name_M.get()
        ip_data = self.ip.get()
        slaveunit_data = self.slaveunit.get()
        
        registers_data = list()
        for register in self.registers:
            registers_data.append(register.get())

        self.class_.measurement.name = name_M_data
        self.class_.measurement.ip = ip_data
        self.class_.measurement.slaveunit = slaveunit_data
        self.class_.measurement.update_registers(registers_data)

        self.class_.actualizar()
        self.top.destroy()

        self.app.logger.log(logging.INFO, 'Linea (' + self.name.get() + ') modificada satisfactoriamente')
        self.app.changes = 1

## Drag and Drop capabilities

def bind_p(widget, app):
    widget.bind("<Button-1>", drag_start_p)
    widget.bind("<B1-Motion>", lambda e: drag_motion_p(e, app))
    widget.bind("<Button-3>", lambda e: menu_items(e, app))

def drag_start_p(event):
    widget = event.widget.master
    widget.startX = event.x
    widget.startY = event.y

def drag_motion_p(event, app):
    app.changes = 1
    widget = event.widget.master
    x = widget.winfo_x() - widget.startX + event.x
    y = widget.winfo_y() - widget.startY + event.y
    x, y = placeGrid(x, y, app)
    widget.place(x=x,y=y)
    class_ = app.all_elements_widget[widget]
    class_.draw_lines()
    for line in class_.parent.lines:
        terminal = line.get_other_terminal(class_.parent)
        terminal.widget.draw_lines()

def find_nearest(array, value):
    array = np.asarray(array)
    idx = (np.abs(array - value)).argmin()
    return array[idx]

def placeGrid(x, y, app):
    x0 = find_nearest(app.grid_X, x)
    y0 = find_nearest(app.grid_Y, y)
    return x0, y0 

## Gestion de archivos

class Database():
    def __init__(self, app):  

        ## Variable para que la función 'importar_' sepa si debe cerrar este form
        self.ini = 1
        self.app = app
        self.h5file = None

    def getFile(self):
        return self.h5file

    def form(self):

        if self.h5file:
            result = messagebox.askyesno('Nuevo archivo', '¿Quieres guardar el proyecto anterior?')
            if result == True: self.export(self.h5file)

        self.ini = 1

        self.top = Toplevel(self.app)
        self.top.mainloop
        
        self.top.resizable(False, False)
        self.top.grab_set()
        if self.h5file:
            self.top.protocol('WM_DELETE_WINDOW', lambda e:self.top.destroy())
            self.top.bind('<Escape>', lambda e: self.top.destroy())
            self.top.title('Escoger nuevo archivo.')
        else:
            self.top.bind('<Escape>', self.app.on_closing)
            self.top.protocol('WM_DELETE_WINDOW', lambda:self.app.on_closing(''))
            self.top.title("Inicio")
        self.top.focus_force()

        if self.h5file: 
            t1 = 'Escoger nombre del proyecto'
        else:
            t1 = '...o importar proyecto'

        self.lFrame = LabelFrame(self.top, text='Nuevo proyecto')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)
        Label(self.lFrame, text=t1, font=font, width=25, pady=5, padx=5, anchor=W).grid(row=0, column=0, columnspan=3, sticky=W)

        self.name = StringVar()

        Label(self.lFrame, text="Nombre:", font=font, width=25, anchor=W, pady=5).grid(row=2, column=0, padx=10, pady=5)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=2, column=1, columnspan=2, padx=10, pady=10)

        Button(self.top, text="Aceptar", command=self.newFile, font=font).grid(row=2, column=8, sticky=E, padx=0, pady=5)

        if self.h5file: 
            Button(self.top, text="Cancelar", command=lambda:self.top.destroy(), font=font).grid(row=2, column=9, padx=5, pady=5)
        else:
            Button(self.lFrame, text='+', command=self.import_).grid(row=2, column=3, padx=10, pady=10)
            Button(self.top, text="Cancelar", command=lambda:self.app.on_closing(''), font=font).grid(row=2, column=9, padx=5, pady=5)
            

        self.app.center_top(self.top)

    def newFile(self):
        filename = self.name.get() + '.h5'

        if filename == '.h5':
            messagebox.showerror(title='Error', message='Escoger nombre de proyecto', parent=self.top)
            return

        self.app.dMidFrame.place(relx=0.5, rely=0.06, relwidth=1.0, relheight= 0.69, anchor="n")

        if self.h5file:
            uuid_list = list(self.app.all_elements_object.keys())
            for uuid in uuid_list:
                self.app.all_elements_object[uuid].delete()

    

        self.h5file = open_file(filename, "w")

        group = self.h5file.create_group('/', 'data', title='Data')
        self.h5file.create_table(group, 'static_elements', Elements, title='Static Elements')
        self.h5file.create_table(group, 'segment_elements', Segments, title='Segment Elements')
        self.h5file.create_table(group, 'alarms', Alarms, title='Alarms')
        self.h5file.create_table(group, 'measurement_elements', Measurements, title='Measurement Elements')

        self.app.logger.log(logging.INFO, 'Nuevo archivo creado.')

        self.top.destroy()
        self.h5file.close()

        self.app.dMidFrame.place_forget()

    def import_(self):
        cwd = os.getcwd()
        self.app.logger.log(logging.INFO, 'Importando...')
        filename = filedialog.askopenfilename(initialdir = cwd,
                                            title = "Select a File",
                                            filetypes = (("HDF5 files",
                                                            "*.h5*"),
                                                        ("all files",
                                                            "*.*")))
        if filename == '': return
        if self.ini == 1:
            self.ini = 0
            self.top.destroy()

        self.h5file = open_file(os.path.basename(filename), mode='r+')
        table_elm = self.h5file.root.data.static_elements
        table_sgm = self.h5file.root.data.segment_elements
        table_alm = self.h5file.root.data.alarms
        table_msm = self.h5file.root.data.measurement_elements

        uuid = []
        name = []
        nominalPower = []
        nominalVoltage1 = []
        nominalVoltage2 = []
        element_type = []
        grid_x = []
        grid_y = []
        orientation = []

        for element in table_elm.iterrows():
            uuid.append(element['uuid'].decode())
            name.append(element['name'].decode())
            nominalPower.append(element['nominalPower'])
            nominalVoltage1.append(element['nominalVoltage1'])
            nominalVoltage2.append(element['nominalVoltage2'])
            element_type.append(element['element_type'].decode())
            grid_x.append(element['grid_x'])
            grid_y.append(element['grid_y'])
            orientation.append(element['orientation'])

        for i in range(0,len(uuid)):
            if element_type[i] == 'FV':
                frame = FrameObject(window_= self.app.MidFrame, name=name[i], image=self.app.FVImage, x=grid_x[i], y=grid_y[i], orientation=orientation[i])

                bind_p(frame.figure, self.app)

                new = Photovoltaic(name=name[i], nominalPower=nominalPower[i], nominalVoltage=nominalVoltage1[i], widget=frame, lines=[], oid=uuid[i], app=self.app)

                frame.parent = new

                self.app.dataset_fv[new.oid] = new
                self.app.dataset_fv2[new.oid] = new.name
                self.app.dataset_fv3[new.widget.frame] = new.widget

                self.app.all_elements_object[new.oid] = new
                self.app.all_elements_name[new.oid] = new.name
                self.app.all_elements_widget[new.widget.frame] = new.widget

            elif element_type[i] == 'T':
                frame = FrameObject(window_=self.app.MidFrame, name=name[i], image=self.app.barra, x=grid_x[i], y=grid_y[i], orientation=orientation[i])

                bind_p(frame.figure, self.app)

                new = Terminal(name=name[i], nominalVoltage=nominalVoltage1[i], widget=frame, lines=[], oid=uuid[i], app=self.app)

                frame.parent = new

                self.app.terminals_object[new.oid] = new
                self.app.terminals_name[new.oid] = new.name
                self.app.terminals_widget[new.widget.frame] = new.widget

                self.app.all_elements_object[new.oid] = new
                self.app.all_elements_name[new.oid] = new.name
                self.app.all_elements_widget[new.widget.frame] = new.widget

            elif element_type[i] == 'L':
                frame = FrameObject(window_=self.app.MidFrame, name=name[i], image=self.app.load, x=grid_x[i], y=grid_y[i], orientation=orientation[i])

                bind_p(frame.figure, self.app)

                new = Load(name=name[i], nominalVoltage=nominalVoltage1 [i], widget=frame, lines=[], oid=uuid[i], app=self.app)

                frame.parent = new

                self.app.loads_object[new.oid] = new
                self.app.loads_name[new.oid] = new.name
                self.app.loads_widget[new.widget.frame] = new.widget

                self.app.all_elements_object[new.oid] = new
                self.app.all_elements_name[new.oid] = new.name
                self.app.all_elements_widget[new.widget.frame] = new.widget

            elif element_type[i] == 'TR':
                frame = FrameObject(window_=self.app.MidFrame, name=name[i], image=self.app.trafo, x=grid_x[i], y=grid_y[i], orientation=orientation[i])

                bind_p(frame.figure, self.app)

                new = Trafo(name=name[i], nominalVoltage_h=nominalVoltage1[i], nominalVoltage_l=nominalVoltage2[i], widget=frame, lines=[], oid=uuid[i], app=self.app)

                frame.parent = new

                self.app.trafos_object[new.oid] = new
                self.app.trafos_name[new.oid] = new.name
                self.app.trafos_widget[new.widget.frame] = new.widget

                self.app.all_elements_object[new.oid] = new
                self.app.all_elements_name[new.oid] = new.name
                self.app.all_elements_widget[new.widget.frame] = new.widget

        uuid = []
        name = []
        nominalVoltage = []
        element_type = []
        grid_x = []
        grid_y = []
        terminal1 = []
        terminal2 = []

        for segment in table_sgm.iterrows():
            uuid.append(segment['uuid'].decode())
            name.append(segment['name'].decode())
            nominalVoltage.append(segment['nominalVoltage'])
            element_type.append(segment['element_type'].decode())
            grid_x.append(segment['grid_x'])
            grid_y.append(segment['grid_y'])
            terminal1.append(segment['terminal1'].decode())
            terminal2.append(segment['terminal2'].decode())

        for i in range(0,len(uuid)):
            new_line = LineSegment(name=name[i], nominalVoltage=220, terminal1=self.app.all_elements_object[terminal1[i]], 
                                    terminal2=self.app.all_elements_object[terminal2[i]], widget=None, oid=uuid[i], app=self.app) 
            self.app.all_elements_object[terminal1[i]].lines.append(new_line)
            self.app.all_elements_object[terminal2[i]].lines.append(new_line)
            new_widget = LineObject(window_=self.app.MidFrame, x1=self.app.all_elements_object[terminal1[i]].widget.getX(self.app.all_elements_object[terminal2[i]], new_line), 
                                                        y1=self.app.all_elements_object[terminal1[i]].widget.getY(self.app.all_elements_object[terminal2[i]], new_line),
                                                        x2=self.app.all_elements_object[terminal2[i]].widget.getX(self.app.all_elements_object[terminal1[i]], new_line), 
                                                        y2=self.app.all_elements_object[terminal2[i]].widget.getY(self.app.all_elements_object[terminal1[i]], new_line),
                                                        parent=new_line)
            new_line.widget = new_widget
            new_widget.canvas.update_idletasks()

            new_widget.canvas.tag_bind(new_widget.text, "<Button-3>", lambda e: menu_items(e, self.app))

            self.app.segments_object[new_line.oid] = new_line
            self.app.segments_name[new_line.oid] = new_line.name
            self.app.segments_widget[new_widget.canvas] = new_widget
            
        uuid = []
        name = []
        uuid_element = []
        ip = []
        slaveunit = []
        registers = []

        for element in table_msm.iterrows():
            uuid.append(element['uuid'].decode())
            name.append(element['name'].decode())
            uuid_element.append(element['uuid_element'].decode())
            ip.append(element['ip'].decode())  
            slaveunit.append(element['slaveunit'])
            registers.append(element['registers'].decode())  

        for i in range(0,len(uuid)):
            elemento = {**self.app.all_elements_object, **self.app.segments_object}[uuid_element[i]]
            new_measure = Measurement(name=name[i], oid=uuid[i],  elemento=elemento, app=self.app,
                                    ip=ip[i], slaveunit=slaveunit[i], registers=registers[i], decode=True)
            elemento.measurement = new_measure
            self.app.measurements_object[new_measure.oid] = new_measure

        uuid = []
        name = []
        uuid_element = []
        prioridad = []
        umbral = []
        variable = []
        condicion = []

        for alarm in table_alm.iterrows():
            uuid.append(alarm['uuid'].decode())
            name.append(alarm['name'].decode())
            variable.append(alarm['variable'].decode())
            condicion.append(alarm['condicion'].decode())
            umbral.append(alarm['umbral'])
            prioridad.append(alarm['prioridad'].decode())
            uuid_element.append(alarm['elemento'].decode())

        for i in range(0,len(uuid)):
            elemento = {**self.app.all_elements_object, **self.app.segments_object}[uuid_element[i]]
            new_alarm = Alarm(name=name[i], elemento=elemento, variable=variable[i], condicion=condicion[i], umbral=umbral[i], prioridad=prioridad[i], app=self.app)
            self.app.alarms_object[new_alarm.oid] = new_alarm
            self.app.alarms_status[new_alarm.oid] = 0

        for sgm in self.app.all_elements_object.values():
            sgm.widget.draw_lines()

        self.app.changes = 0

        self.app.logger.log(logging.INFO, 'Archivo ' + os.path.basename(filename) + ' importado exitosamente.')
        self.h5file.close()

        self.app.dMidFrame.place_forget()

    def export(self, file=None):
        save = False

        if file == None:
            cwd = os.getcwd()

            filename = filedialog.asksaveasfilename(initialdir = cwd,
                                                title = "Save File",
                                                filetypes = (("HDF5 files",
                                                                "*.h5*"),
                                                            ("all files",
                                                                "*.*")))
            if filename == '': 
                return
        else:
            filename = file.filename
            save = True

        if '.h5' not in filename: filename = filename + '.h5'
        self.h5file = open_file(os.path.basename(filename), mode='w')
        group = self.h5file.create_group('/', 'data', title='Data')
        table_elm = self.h5file.create_table(group, 'static_elements', Elements, title='Static Elements')
        table_sgm = self.h5file.create_table(group, 'segment_elements', Segments, title='Segment Elements')
        table_alm = self.h5file.create_table(group, 'alarms', Alarms, title='Alarms')
        table_msm = self.h5file.create_table(group, 'measurement_elements', Measurements, title='Measurement Elements')

        elemento = table_elm.row
        segmento = table_sgm.row
        alarma = table_alm.row
        measure = table_msm.row

        for element in self.app.all_elements_object.values():
            elemento['uuid'] = element.oid
            elemento['name'] = element.name
            try:
                elemento['nominalPower'] = element.nominalPower
            except:
                elemento['nominalPower'] = 0
            try:
                elemento['nominalVoltage1'] = element.nominalVoltage_h
                elemento['nominalVoltage2'] = str(element.nominalVoltage_l)
            except:
                elemento['nominalVoltage1'] = element.nominalVoltage
                elemento['nominalVoltage2'] = element.nominalVoltage
            elemento['element_type'] = element.type
            elemento['grid_x'] = element.widget.frame.winfo_x()
            elemento['grid_y'] = element.widget.frame.winfo_y()
            elemento['orientation'] = element.widget.orientation

            elemento.append()

        table_elm.flush()

        for segment in self.app.segments_object.values():
            segmento['uuid'] = segment.oid
            segmento['name'] = segment.name
            segmento['nominalVoltage'] = segment.nominalVoltage
            segmento['element_type'] = segment.type
            segmento['grid_x'] = segment.widget.canvas.winfo_x()
            segmento['grid_y'] = segment.widget.canvas.winfo_y()
            segmento['terminal1'] = segment.terminal1.oid
            segmento['terminal2'] = segment.terminal2.oid

            segmento.append()

        table_sgm.flush()

        for alarm in self.app.alarms_object.values():
            alarma['name'] = alarm.name
            alarma['elemento'] = alarm.elemento.oid
            alarma['variable'] = alarm.variable
            alarma['condicion'] = alarm.condicion
            alarma['umbral'] = alarm.umbral
            alarma['prioridad'] = alarm.prioridad
            alarma['uuid'] = alarm.oid

            alarma.append()

        table_alm.flush()

        for measure_ in self.app.measurements_object.values():
            measure['uuid'] = measure_.oid
            measure['name'] = measure_.name
            measure['uuid_element'] = measure_.elemento.oid
            measure['ip'] = measure_.ip
            measure['slaveunit'] = int(measure_.slaveunit)
            measure['registers'] = measure_.code_regs()

            measure.append()

        table_msm.flush()

        self.app.changes = 0

        if save == True:
            self.app.logger.log(logging.INFO, 'Archivo guardado satisfactoriamente.')
        else:
            self.app.logger.log(logging.INFO, 'Exportación finalizada satisfactoriamente.')

        self.h5file.close()

    def on_closing(self, event):
        if self.app.changes == 1:
            result = messagebox.askyesnocancel('Salir', '¿Quieres guardar antes de salir?')
            if result == True:
                self.export(self.h5file)
            elif result == False:
                pass
            else:
                return
        self.app.destroy()

## Historizador    

class HistoricalData():
    def __init__(self, app, class_):

        self.top = Toplevel(app)
        self.top.geometry('1200x700')
        self.top.mainloop
        self.top.title("Data historica de "+class_.name)
        self.top.resizable(False, False)
        self.top.focus_force()
        self.data_replay = False
        self.top.bind('<Escape>', lambda e: self.on_closing())
        self.lock = None
        self.file = None
        self.app = app

        if app.monitoreo_object.lock != None:
            self.lock = app.monitoreo_object.lock
            self.file = app.monitoreo_object.getFilename()

        self.top.protocol('WM_DELETE_WINDOW', self.on_closing)

        self.class_ = class_

        ### Frame para escoger las variables eléctricas a graficar ###

        self.lFrame_hd = LabelFrame(self.top, text='Opciones')
        self.lFrame_hd.grid(row=0, column=0, padx=10, pady=10, sticky='ewn')
        self.top_label = Label(self.lFrame_hd, text='Escoger variable electrica a graficar:', width=60, anchor='w', padx=10)
        self.top_label.pack()

        self.var = StringVar()

        self.VARIABLES = ['Tensión AN', 'Tensión BN', 'Tensión CN', 'Tensión AB', 'Tensión BC',
            'Tensión CA', 'Tensión FF', 'Tensión FN', 'Corriente Media', 'Corriente A', 
            'Corriente B', 'Corriente C', 'Corriente Neutro', 'Potencia Activa A', 'Potencia Activa B',
            'Potencia Activa C', 'Potencia Activa 3P', 'Potencia Reactiva A', 'Potencia Reactiva B',
            'Potencia Reactiva C', 'Potencia Reactiva 3P', 'Frecuencia', 'Factor de Potencia']

        Label(self.lFrame_hd, text='').pack() 

        for text in self.VARIABLES:
                Radiobutton(self.lFrame_hd, text=text, variable=self.var, value=text, anchor='w', width=50, padx=20).pack()

        ### Frame para graficar los datos ###

        self.rFrame_hd = Frame(self.top)
        self.rFrame_hd.grid(row=0, column=1, sticky='ew')

        self.buttonAccept = Button(self.top, text='Graficar', command=self.plot_data)
        self.buttonAccept.place(x=150, y=660)
        self.buttonClear = Button(self.top, text='Limpiar', command=self.clear_frame)
        self.buttonClear.place(x=250, y=660)

        app.center_top(self.top)

    def plot_data(self):
        try: 
            with self.lock:
                df = pd.read_csv(self.file, encoding='latin-1')
                name_value = self.class_.name
                df = df.query("name == @name_value")
                variable_value = self.var.get()
                df = df.query("variable == @variable_value")
                y_values = df['value'].values.tolist()
                x_values = np.linspace(0, len(y_values), num=len(y_values))
        except:
            messagebox.showerror(title='Error', message='Sistema inactivo', parent=self.top)
            return

        self.figure = Figure(figsize=(7.4, 6), dpi=100)
        self.figure_canvas = FigureCanvasTkAgg(self.figure, self.rFrame_hd)
        NavigationToolbar2Tk(self.figure_canvas, self.rFrame_hd)
        self.axes = self.figure.add_subplot()

        self.axes.plot(x_values, y_values, 'r-')
        self.axes.set_title(variable_value + ' de ' + self.class_.name)
        self.axes.set_ylabel('Watts')

        self.figure_canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)

        self.app.after(1000, self.update_data)

    def update_data(self):

        self.axes.clear()

        try: 
            with self.lock:
                df = pd.read_csv(self.file, encoding='latin-1')
                name_value = self.class_.name
                df = df.query("name == @name_value")
                variable_value = self.var.get()
                df = df.query("variable == @variable_value")
                y_values = df['value'].values.tolist()
                x_values = np.linspace(0, len(y_values), num=len(y_values))
        except NameError:
            messagebox.showerror(title='Error', message='Sistema inactivo', parent=self.top)
            return

        self.axes.plot(x_values, y_values, 'r-')
        self.axes.set_title(variable_value + ' de ' + self.class_.name)
        self.axes.set_ylabel('Watts')
        self.figure_canvas.draw_idle()
        self.data_replay = self.app.after(1000, self.update_data)

    def clear_frame(self):
        for child in self.rFrame_hd.winfo_children():
            child.destroy()
        if self.data_replay: self.app.after_cancel(self.data_replay)
    
    def on_closing(self):
        if self.data_replay: self.app.after_cancel(self.data_replay)
        self.top.destroy()

## Funcion para gestionar el menu que se abre con click derecho a los objetos

def menu_items(event, app):  
    frame = event.widget.master
    if frame == app.MidFrame:
        object_ = app.segments_widget[event.widget]
    else:
        object_ = app.all_elements_widget[frame]
    m = Menu(app, tearoff = 0)
    m.add_command(label ='Parameters', command=lambda:config(object_.parent, app))
    m.add_command(label ='Historical Data', command=lambda:HistoricalData(app, object_.parent))
    m.add_command(label ='Set Alarm', command=lambda:FormAlarm(app, object_.parent))
    m.add_separator()
    if object_.parent.type != 'LT': 
        m.add_command(label ="Rotate", command=object_.rotate)
        m.add_separator()
    m.add_command(label ="Delete", command=lambda:object_.parent.delete())

    try:
        m.tk_popup(event.x_root, event.y_root)
    finally:
        m.grab_release()

## Frame para gestionar alarmas

class AlarmManager():  
    def __init__(self, app):

        app.alarm_manager = self

        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title('Alarm Manager')
        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()
        self.top.resizable(False, False)
        self.app = app

        ## Frame que contiene al Treeview
        self.tree_frame = Frame(self.top)
        self.tree_frame.pack(pady=10, padx=20)

        ## Scrollbar de Treeview
        self.tree_scroll = Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side=RIGHT, fill=Y)

        ## Treeview
        self.alarms = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll.set, selectmode='extended')
        self.alarms.pack()

        ## Scrollbar config
        self.tree_scroll.config(command=self.alarms.yview)

        ## Definir columnas de alarmas
        self.alarms['columns'] = ('Alarma', 'Elemento', 'Variable', 'Condicion', 'Prioridad', 'UUID')

        ## Formato columnas
        self.alarms.column('#0', width=0, stretch=NO)
        self.alarms.column('Alarma', anchor=CENTER, width=100)
        self.alarms.column('Elemento', anchor=CENTER, width=100)
        self.alarms.column('Variable', anchor=CENTER, width=100)
        self.alarms.column('Condicion', anchor=CENTER, width=100)
        self.alarms.column('Prioridad', anchor=CENTER, width=100)
        self.alarms.column('UUID', width=0, stretch=NO)

        ## Headings
        self.alarms.heading('#0', text='', anchor=CENTER)
        self.alarms.heading('Alarma', text='Alarma', anchor=CENTER)
        self.alarms.heading('Elemento', text='Elemento', anchor=CENTER)
        self.alarms.heading('Variable', text='Variable', anchor=CENTER)
        self.alarms.heading('Condicion', text='Condicion', anchor=CENTER)
        self.alarms.heading('Prioridad', text='Prioridad', anchor=CENTER)
        self.alarms.heading('UUID', text='UUID', anchor=CENTER)

        ## Stripped Row Tags
        self.alarms.tag_configure('oddrow', background='white')
        self.alarms.tag_configure('evenrow', background='lightblue')

        ## Add data
        count=0
        for alarma in self.app.alarms_object.values():
            if count % 2 == 0: 
                self.alarms.insert(parent='', index='end', iid=count, text='', 
                                    values=(alarma.name, alarma.elemento.name, alarma.variable, str(alarma.condicion + str(alarma.umbral)), alarma.prioridad, alarma.oid), tags=('evenrow'))
            else:
                self.alarms.insert(parent='', index='end', iid=count, text='', 
                                    values=(alarma.name, alarma.elemento.name, alarma.variable, str(alarma.condicion + str(alarma.umbral)), alarma.prioridad, alarma.oid), tags=('oddrow'))
            count += 1

        ## Buttons

        self.commandFrame = LabelFrame(self.top, text='Comandos')
        self.commandFrame.pack(fill='x', expand='yes', padx=20, pady=10)

        self.create_alarm = Button(self.commandFrame, text='Crear Alarma', command=lambda:FormAlarm(self.app))
        self.create_alarm.grid(row=0, column=0, pady=10, padx=10)
        self.delete_alarm = Button(self.commandFrame, text='Borrar Alarma', command=self.borrar_alarma)
        self.delete_alarm.grid(row=0, column=1, pady=10, padx=10)

        self.app.center_top(self.top)

    def borrar_alarma(self):
        try:
            x = self.alarms.selection()[0]
            selected = self.alarms.focus()
            values = self.alarms.item(selected, 'values')
            self.alarms.delete(x)
            self.app.alarms_object[values[5]].delete()
        except:
            messagebox.showerror(title='Error', message='No se selecciono alarma', parent=self.top)
            self.top.focus_force()
            return

    def update(self):
        for item in self.alarms.get_children():
            self.alarms.delete(item)

        count=0
        for alarma in self.app.alarms_object.values():
            if count % 2 == 0: 
                self.alarms.insert(parent='', index='end', iid=count, text='', 
                                    values=(alarma.name, alarma.elemento.name, alarma.variable, str(alarma.condicion + str(alarma.umbral)), alarma.prioridad, alarma.oid), tags=('evenrow'))
            else:
                self.alarms.insert(parent='', index='end', iid=count, text='', 
                                    values=(alarma.name, alarma.elemento.name, alarma.variable, str(alarma.condicion + str(alarma.umbral)), alarma.prioridad, alarma.oid), tags=('oddrow'))
            count += 1

## Monitoreo

class Monitoreo():
    def __init__(self, app):
        self.app = app
        self.lock = None
        self.file = None

    def form(self):
        self.top = Toplevel(app)
        self.top.mainloop
        self.top.title("Iniciar monitoreo")
        self.top.resizable(False, False)
        self.top.grab_set()
        self.top.bind('<Escape>', self.top.destroy)
        self.top.focus_force()

        self.lFrame = LabelFrame(self.top, text='Datafile')
        self.lFrame.grid(row=0, column=0, sticky=W, columnspan=10, padx=10, pady=5)
        Label(self.lFrame, text='Importar o crear nuevo archivo .csv', font=font, pady=5, padx=5).grid(row=0, column=0, columnspan=10, sticky=W)

        self.name = StringVar()

        Label(self.lFrame, text="Nombre:", font=font, width=20, anchor=W, pady=5).grid(row=2, column=0, padx=10, pady=5)
        Entry(self.lFrame, textvariable=self.name, width=30).grid(row=2, column=1, columnspan=2, padx=10, pady=10)
        Button(self.lFrame, text='+', command=self.fileSelect).grid(row=2, column=3, padx=10, pady=10)

        Button(self.top, text="Aceptar", command=self.play, font=font).grid(row=2, column=8, sticky=E, padx=0, pady=5)
        Button(self.top, text="Cancelar", command=self.top.destroy, font=font).grid(row=2, column=9, padx=5, pady=5)

        self.app.center_top(self.top)

    def alarm_loop(self):
        for alarm in app.alarms_object.values():
            if alarm.elemento.measurement.connected:
                index = alarm.elemento.measurement.variables.index(alarm.variable)
                if float(alarm.elemento.measurement.rvalues[index][1]) > float(alarm.umbral) and self.app.alarms_status[alarm.oid] == 0:
                    self.app.alarms_status[alarm.oid] = 1
                    self.app.logger.log(logging.WARNING, alarm.variable + ' de ' + alarm.elemento.name + ' en niveles preocupantes: ' + str(alarm.elemento.measurement.rvalues[index][1]) + ' ' + alarm.elemento.measurement.unidades[index])
                elif float(alarm.elemento.measurement.rvalues[index][1]) < float(alarm.umbral) and self.app.alarms_status[alarm.oid] == 1:
                    self.app.alarms_status[alarm.oid] = 0

    def update_loop(self, ms):

        self.name = self.getFilename()

        st = time.time()
        if self.name[-4:] != '.csv':
            self.name = self.name + '.csv'

        for objeto in {**self.app.all_elements_object, **self.app.segments_object}.values():
            if objeto.measurement.connected:
                objeto.measurement.read_registers()
                with open(self.name, 'a', newline='') as file:
                    if self.first_loop:
                        self.first_loop = False
                        writer = csv.writer(file)
                        writer.writerow(('uuid', 'name', 'datetime', 'variable', 'value'))
                    objeto.measurement.write_to_file(self.lock, file)
                objeto.widget.update_labels()

        thread = threading.Thread(target=self.alarm_loop)
        thread.start()
        thread.join()
        et = time.time()
        self.update_replay = app.after(np.max(int(ms - (et - st)*1000), 0), self.update_loop, ms)

    def play(self):

        self.first_loop = True

        self.lock = Lock()

        self.app.logger.log(logging.INFO, 'Inicializando monitoreo.')

        for measure in self.app.measurements_object.values():
            measure.connect()

        self.ms = 1000
        self.update_loop(self.ms)

    def stop_loop(self):
        try:
            self.app.after_cancel(self.update_replay)
            for measure in self.app.measurements_object.values():
                if measure.connected == True: measure.disconnect()
            self.lock = None
        except:
            self.app.logger.log(logging.INFO, 'No se esta monitoreando.')
            return

        self.app.logger.log(logging.INFO, 'Deteniendo monitoreo.')

    def getFilename(self):
        self.top.destroy()
        try: 
            return self.name.get()
        except:
            return self.name

    def fileSelect(self):
        cwd = os.getcwd()

        self.filename = filedialog.askopenfilename(initialdir = cwd,
                                            title = "Select a File",
                                            filetypes = (("CSV files",
                                                            "*.csv*"),
                                                        ("all files",
                                                            "*.*")))

        if self.filename == '': return
        self.top.destroy()

        self.play()

# Clases para definir tablas del archivo HDF5

class Elements(IsDescription):
    uuid = StringCol(64)
    name = StringCol(32)
    nominalPower = Int32Col()
    nominalVoltage1 = Int32Col()
    nominalVoltage2 = Int32Col()
    element_type = StringCol(16)
    grid_x = Int32Col()
    grid_y = Int32Col()
    orientation = Int32Col()

class Segments(IsDescription):
    uuid = StringCol(64)
    name = StringCol(64)
    nominalVoltage = Int32Col()
    element_type = StringCol(16)
    terminal1 = StringCol(64)
    terminal2 = StringCol(64)
    grid_x = Int32Col()
    grid_y = Int32Col()

class Alarms(IsDescription):
    name = StringCol(64)
    elemento = StringCol(64)
    variable = StringCol(64)
    condicion = StringCol(32)
    umbral = Int32Col()
    prioridad = StringCol(32)
    uuid = StringCol(64)

class Measurements(IsDescription):
    uuid = StringCol(64)
    name = StringCol(64)
    uuid_element = StringCol(64)
    ip = StringCol(64)
    slaveunit = Int32Col()
    registers = StringCol(64)

if __name__ == '__main__':
    app = App()
    app.mainloop()