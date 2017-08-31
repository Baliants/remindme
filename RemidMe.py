# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 12:12:08 2017

@author: Suren Baliants

To do:


1) переделать механизм получения дат. 
    Получать даты в новый список, а потом его сортировать и выводить. 
    Добавит сложность log(N) зато решит проблему если список дат в Екселе не отсортирован.

2) добавить кнопку "Настройка" с функциями
- выбора пути к файлу данных
- изменения языка
- выбора цвета заголовков и вкл/выключение функций разных цветов заголовка.    
"""

import datetime               #Импорт библиотеки времени

from tkinter import *         #Импорт библиотеки графического интерфейса

import win32com.client        #импорт библиотеки для работы с Екселем

import os.path                #импорт библиотеки для задания относительных путей


# решения вопроса бага абсолюдных после компиляции
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

#
def MisingFile():
    MisingFileWindow = Tk()
    MisingFileWindow.wm_title("ОШИБКА")
    MisingFileWindow.iconbitmap(resource_path("remind.ico"))
    txt = Label(MisingFileWindow, 
                 text="\nФайл данных не найден!\n\n  Разместите файл data.xlsx в папке с программой  \n", 
                 font = "Arial 10") 
    txt.pack()
    MisingFileWindow.mainloop() 

# запуск приложения екселя
def RunData(event):
    '''
    Запуск файла данных
    '''
    try:
        shell.Run('data.xlsx')
    except :
        MisingFile()   
    
# запуск кнопки "О программе"
def About(event):
    '''
    Вызов окна "О программе"
    '''
    AboutWindow = Tk()
    AboutWindow.wm_title("О программе")
    AboutWindow.iconbitmap(resource_path("remind.ico"))
    txt1 = Label(AboutWindow, 
                 text="\n   Автор программы: Бальянц Сурэн. bsv181@gmail.com \n   Версия программы – 1.0, 30.08.2017 \n------------------------", 
                 font = "Arial 12 bold") 
    txt1.pack(anchor= NW)
    
    txtAbout = (
"   Данная программа была написана как выпускной проект на языке Python."
"\n   Реализация на другом языке была бы более оптимальна, но требовалось использовать именно Python. "
"\n   Цель программы –используя базу в формате Excel напоминать о событиях за разное количество дней. "
"\n   Программа будет особенно полезна в отделах кадров. Она позволяет легко проследить разные сроки, "
"\n   связанные с большим количеством разных людей и событий. При этом сохраняются все возможности  "
"\n   гибкой работы с базой данных в формате Excel. "
"\n\n   Для корректной работы программы необходимо следующее:"
"\n   - файл Excel должен называться data.xlsx и находится в папке с программой; "
"\n   - ячейка Z1 должна содержать количество дней, за сколько нужно напомнить до события; "
"\n   - данные о дате события должны начинается с ячейки A2 и ниже; "
"\n   - один лист = один срок напоминания. Т.е. если у даны однотипные события, которые  "
"\n     должны иметь разные сроки напоминания, для корректной работы программы, их нужно  "
"\n     разделить на разные листы и выставить соответствующий срок напоминания "
"\n   - имена листам лучше всего давать информативные, т.к. они будут обображатся в программе. \n"
)

         
    txt2 = Label(AboutWindow, text=txtAbout,
            font = "Arial 10",
            justify = "left")
    txt2.pack(anchor= NW)
    txt2.pack(fill = "both")
   
    AboutWindow.mainloop()  

# ====== основная функция =====   
def RemindMe(book):
 #=================== открытие окна и создание кнопок     
    window = Tk()
    window.wm_title("Напоминалка")
    window.iconbitmap(resource_path("remind.ico"))

    absPath1 = os.path.abspath(resource_path("open.gif"))
    absPath2 = os.path.abspath(resource_path("about.gif"))
      
    image1 = PhotoImage(file=absPath1) # импорт иконки открытия файла
    image2 = PhotoImage(file=absPath2) # импорт иконки открытия файла
    
    opnData = Button(window,                  #родительское окно
                 text=" Открыть файл данных                       ",       #надпись на кнопке
                 #width=30,height=1,     #ширина и высота
                 bg="white",fg="black",
                 font = "Arial 8",
                 image = image1,
                 compound="left") #цвет фона и надписи
    opnData.bind("<Button-1>", RunData)       #при нажатии ЛКМ на кнопку вызывается функция Hello
    opnData.pack(side = "top", anchor= NW)
    #opnData.grid(row = 1, column = 1, fill = 'both')
    #opnData.place(x = 20, y = 40)
    
    about = Button(window,                  #родительское окно
                 text=" О программе                                                              ",       #надпись на кнопке
                 #width=300,height=15,     #ширина и высота
                 bg="white",fg="black",
                 image = image2, 
                 compound="left",
                 font = "Arial 8") #цвет фона и надписи
    about.bind("<Button-1>", About)       #при нажатии ЛКМ на кнопку вызывается функция Hello
    about.place(x = 175, y = 0)
              
    #=============end====== открытие окна и создание кнопок 
   
    blankLine = Label(window, text="", font = "Arial 2")        
    blankLine.pack()  
    
    now = datetime.datetime.now() #Сегодняшнаяя дата
    nowInISO = now.isocalendar()
    daysFromThisNY = ((nowInISO[1]-1)*7)+nowInISO[2]+1         #получение количесва дней с начала года
    colourList = ["midnight blue"]         #цвет заголовка события. Можно задать несколько цвет (каждый заголовок будет отдельным цветом)
    colourNum = 0                          # Для этого отключить colourNum и включить код  ниже для изменения цвета 
    rowInTk = 2
    for listName in book.Worksheets:               #получение имен
        sheet = book.Sheets(listName.Name)          #определение активного листа активном
        daysToRimind = int(sheet.Cells(1,26).value)           #Получение количества дней за сколько нужно напоминать
        heading = Label(window, text=" -----   "+listName.Name+"   ----- ", fg=colourList[colourNum], font = "Arial 10 bold")
        heading.pack()
    #   код позволяющий менять цвет заголовков события    См.  colourList   colourNum
    #    if colourNum < len(colourList)-1:
    #        colourNum+=1
    #    else:
    #        colourNum=0
    #    heading2 = Label(window, text="(напоминать за "+str(daysToRimind)+" дней)", fg="red", font = "Arial 12")
    #    heading2.pack() 
        rowNum = 2
        while sheet.Cells(rowNum,1).value != None:
            date = sheet.Cells(rowNum,1).value
            dateInISO = date.isocalendar()                                          # конвертация даты в кортеж ISO
            if nowInISO[0]-dateInISO[0] == 1:
                daysFromThisNY += 365
            
            daysFromNY = ((dateInISO[1]-1)*7)+dateInISO[2]+1
            
            difDate = daysFromNY-daysFromThisNY
            who = sheet.Cells(rowNum,2).value
            
            if who == None:                     #обходит ошибку пустых ячеек
                who = "Нет данных"
                
            if difDate <= daysToRimind and difDate > 1:         
                w = Label(window, text="  Через "+str(difDate)+" д. "+str(who)+" ("+str(datetime.datetime.date(date))+")", font = "Arial 8")        
                w.pack(anchor= NW)
            elif difDate == 0:
                w = Label(window, text="  Cегодня "+who+" ("+str(datetime.datetime.date(date))+")", fg="blue", font = "Arial 8")
                w.pack(anchor= NW)
            elif difDate == -1:
                w = Label(window, text="  Вчера "+who+" ("+str(datetime.datetime.date(date))+")", font = "Arial 8")
                w.pack(anchor= NW)
            elif difDate == 1:
                w = Label(window, text="  Завтра "+who+" ("+str(datetime.datetime.date(date))+")", font = "Arial 8")
                w.pack(anchor= NW)
            rowNum +=1
        blankLine = Label(window, text="", font = "Arial 2")        
        blankLine.pack()  
          
    #закрываем книгу
    book.Close()
    
    #закрываем COM объект
    Excel.Quit()
    
    window.mainloop()        
    
    
# ================    
# Собственно программа    
# ================   
# подготовка к запуску екселя
shell = win32com.client.Dispatch("WScript.Shell")
Excel = win32com.client.Dispatch("Excel.Application")
xlsPath = "data.xlsx"
absPath = os.path.abspath(xlsPath)

# проверка, существует ли такой файл
try:
    book = Excel.Workbooks.Open(absPath)
    RemindMe(book)                           # вызов основной функции программы           
except :
    MisingFile()
