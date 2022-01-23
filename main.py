#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jan 12 00:33:58 2022

@author: ozgur oney
"""
from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from tkinter import *
from tkinter import messagebox

window = Tk()
window.title('Workpaper Creator')
# window.geometry('400x300')
window.config(bg='#456')
f = ('sans-serif', 13)
btn_font = ('sans-serif', 10)
bgcolor = '#BF5517'

genvar = StringVar()
genopt = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
genvar.set('1')
# den_sayisi=['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']

def clear_inputs():
    baslangictarihi.delete(0, 'end')
    bitistarihi.delete(0, 'end')
    denetim_elemani.delete(0, 'end')
    
def generate():
     return None
# frames
frame = Frame(window, padx=20, pady=20, bg=bgcolor)
frame.pack(expand=True, fill=BOTH)

#labels
Label(frame, text = "Baş. tarihi giriniz: ", font=f, 
      bg=bgcolor).grid(column=0, row=0, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Bit. tarihi giriniz: ", font=f, 
      bg=bgcolor).grid(column=0, row=1, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Den. Elemanı İsim: ", font=f, 
      bg=bgcolor).grid(column=0, row=2, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Kaç denetim adımı?: ", font=f, 
      bg=bgcolor).grid(column=0, row=3, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))

#input widgets
baslangictarihi = Entry(frame, width=20, font=f)
baslangictarihi.grid(row=0, column=1)
bitistarihi = Entry(frame, width=20, font=f)
bitistarihi.grid(row=1, column=1)
denetim_elemani = Entry(frame, width=20, font=f)
denetim_elemani.grid(row=2, column=1)
sayi = OptionMenu(
    frame, 
    genvar,
    *genopt
)
sayi.grid(row=3, column=1, pady=(5,0))
sayi.config(width=16, font=f)

#defaults
baslangictarihi.insert(0,'1.1.2000')
bitistarihi.insert(0,'1.1.2000')
denetim_elemani.insert(0,'John Doe')

submit_btn = Button(
    btn_frame,
    text='Generate Word',
    command=generate,
    font=btn_font,
    padx=10, 
    pady=5
)
submit_btn.pack(side=LEFT, expand=True, padx=(15, 0))

clear_btn = Button(
    btn_frame,
    text='Clear',
    command=clear_inputs,
    font=btn_font,
    padx=10, 
    pady=5,
    width=7
)
clear_btn.pack(side=LEFT, expand=True, padx=15)

exit_btn = Button(
    btn_frame,
    text='Exit',
    command=lambda:window.destroy(),
    font=btn_font,
    padx=10, 
    pady=5
)
exit_btn.pack(side=LEFT, expand=True)

# mainloop
window.mainloop()