#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jan 12 00:33:58 2022

@author: ozgur oney
"""
import os
from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

window = Tk()
window.title('Workpaper Creator')
# set minimum window size value
window.minsize(600, 400)
 
# set maximum window size value
window.maxsize(600, 400)
 
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
    denetim_elemani_adi.delete(0, 'end')
    pathLabel.delete(0,'end')
    
def select_file():
    window.withdraw()
    file_path = askopenfilename(title="Open file", 
                                       filetypes=[("Word",".docx"),("TXT",".txt"),
                                                  ("All files",".*")])
    if file_path != "":
        print ("you chose file with path:", file_path)

    else:
        print ("you didn't open anything!")
    pathLabel.delete(0, END)
    pathLabel.insert(0, file_path)
    file_path = os.path.dirname(file_path)
    window.deiconify()
    return file_path

def generate():
    # return None
    file_path = pathLabel.get()
    print(file_path)
    with open (file_path, encoding='utf8') as f:
        # read lines for auto-numbering
        lines = f.readlines()
        lines = [line.rstrip() for line in lines]
        #for each line in the doc
        for line in lines:
            #create document
            document = Document()
            
            # Create a character level style Object ("CommentsStyle") then defines its parameters
            obj_styles = document.styles
            obj_charstyle = obj_styles.add_style('CommentStyle', WD_STYLE_TYPE.CHARACTER)
            obj_font = obj_charstyle.font
            obj_font.size = Pt(12)
            obj_font.name = 'Times New Roman'
            
            baslik = document.add_paragraph()
            baslik.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            baslik_run = baslik.add_run('S-100 ÇALIŞMA KAĞIDI', style='CommentStyle').bold=True
            
            tarih = document.add_paragraph()
            tarih.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # u_tarih = input("Tarih giriniz: ")
            tarih_run = tarih.add_run('Tarih: ', style='CommentStyle').bold=True
            u_tarih = tarih.add_run(baslangictarihi.get(), style='CommentStyle')
            
            ck_num= document.add_paragraph()
            ck_num.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            ck_num.add_run('Çalışma Kâğıdı Numarası: ', style='CommentStyle').bold=True
            index = [x for x in range(len(lines)) if line in lines[x]]
            ck_num.add_run(str(index[0]+1).strip("[]"), style='CommentStyle')
            
            denetim_adimi = document.add_paragraph()
            denetim_adimi.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            denetim_adimi.add_run('İlgili Denetim Adımı: ', style='CommentStyle').bold=True
            denetim_adimi.add_run(line, style='CommentStyle')
            
            denetim_elemani = document.add_paragraph()
            denetim_elemani.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            denetim_elemani.add_run('Testi Gerçekleştiren Denetim Elemanı:  ', style='CommentStyle').bold=True
            
            isim = str(denetim_elemani_adi.get()) +  ", Denetim Genel Müdürlüğü, Yetkili BT Denetçi Yardımcısı" 
            denetim_elemani.add_run(isim, style='CommentStyle')
            # print(gui.getAuditorName())
            # denetim_elemani.add_run(gui.getAuditorName(), style='CommentStyle')
            
            orneklem = document.add_paragraph()
            orneklem.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            orneklem.add_run('Örneklem Yöntemi:  ', style='CommentStyle').bold=True
            
            incelenen = document.add_paragraph()
            incelenen.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            incelenen.add_run('İncelenen Dokümanlar:  ', style='CommentStyle').bold=True
        
            bulgu_num = document.add_paragraph()
            bulgu_num.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            bulgu_num.add_run('Bulgu Numarası:  ', style='CommentStyle').bold=True
        
            test_prod = document.add_paragraph()
            test_prod.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            test_prod.add_run('Ayrıntılı Test Prosedürü:  ', style='CommentStyle').bold=True
        
            filename = line.rstrip() + '.docx'
            document.save(filename)     
    messagebox.showinfo("Sonuc", "Dosya(lar) başarıyla oluşturuldu.")
    return None

# frames
frame = Frame(window, padx=20, pady=20, bg=bgcolor)
frame.pack(expand=True, fill=BOTH)

#labels
Label(frame, text= "Den. Programı seçiniz: ", font=f, 
      bg=bgcolor).grid(column=0, row=0, padx=15, pady=15)
pathLabel = Entry(frame,textvariable="")
pathLabel.grid(column=1, row=0, padx=15, pady=15)
Label(frame, text = "Baş. tarihi giriniz: ", font=f, 
      bg=bgcolor).grid(column=0, row=1, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Bit. tarihi giriniz: ", font=f, 
      bg=bgcolor).grid(column=0, row=2, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Den. Elemanı İsim: ", font=f, 
      bg=bgcolor).grid(column=0, row=3, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))
Label(frame, text = "Kaç denetim adımı?: ", font=f, 
      bg=bgcolor).grid(column=0, row=4, padx=15, pady=15)
btn_frame = Frame(frame, bg=bgcolor)
btn_frame.grid(columnspan=2, pady=(50, 0))

#input widgets   
fileSelect = Button(
    frame,
    text='Dosya Seç',
    command=select_file,
    font=btn_font,
    padx=10, 
    pady=5,
    width=3
)
fileSelect.grid(row=0, column=2)
baslangictarihi = Entry(frame, width=20, font=f)
baslangictarihi.grid(row=1, column=1)
bitistarihi = Entry(frame, width=20, font=f)
bitistarihi.grid(row=2, column=1)
denetim_elemani_adi = Entry(frame, width=20, font=f)
denetim_elemani_adi.grid(row=3, column=1)
sayi = OptionMenu(
    frame, 
    genvar,
    *genopt
)
sayi.grid(row=4, column=1, pady=(5,0))
sayi.config(width=16, font=f)

#defaults
baslangictarihi.insert(0,'1.1.2000')
bitistarihi.insert(0,'1.1.2000')
denetim_elemani_adi.insert(0,'John Doe')

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