import wikipedia as w
from googletrans import *
import tkinter
from tkinter import *
from tkinter import messagebox
import win32com.client as robo
import speech_recognition as sr
import pyperclip as pc

speaker = robo.Dispatch('SAPI.spvoice')
trans = Translator()
root = tkinter.Tk()
root.resizable(0, 0)
root.title('Mini Wikipedia')
text = StringVar()
optlist = ['Hindi', 'Japanese', 'German', 'English']
optvar = StringVar()
res = StringVar()
optvar.set('Change Language')
r = sr.Recognizer()
m = sr.Microphone()
label_text = StringVar()
label_text.set('Search Results will be shown below')


def copy_text():
    try:
        pc.copy(resultbar.get(1.0,END))
        messagebox.showinfo('COPIED SUCCESSFULLY', 'Text has been copied to clipboard successfully')
    except EXCEPTION:
        messagebox.showerror('ERROR', 'NO Data found')


def exit_gui():
    exit('force closed')


def interior():
    icon = PhotoImage(file=r'C:\\Users\\rahul\\Pictures\\src.png')
    speech_icon = PhotoImage(file=r'C:\\Users\\rahul\\Pictures\\mike.png')
    speak_icon = PhotoImage(file=r'C:\\Users\\rahul\\Pictures\\voice.png')
    frame1 = Frame(root, bg='cyan', borderwidth=9, relief=SUNKEN)
    frame1.pack(side='top', fill='x')
    e1 = Entry(frame1, borderwidth=9, relief=SUNKEN, width=50, textvariable=text)
    e1.insert(END,'Search Here......')
    e1.grid(row=0, column=0, ipady=10)
    bsrc = Button(frame1, image=icon, borderwidth=9, bg='yellow', command=result)
    bsrc.image = icon
    bsrc.grid(row=0, column=1)
    bvoice = Button(frame1, borderwidth=9, bg='orange', command=voice, image=speech_icon)
    bvoice.image = speech_icon
    bvoice.grid(row=0, column=2)
    bspk = Button(frame1, borderwidth=9, bg='black', image=speak_icon, command=listening)
    bspk.image = speak_icon
    bspk.grid(row=0, column=3)
    frame4 = Frame(root, bg='blue', borderwidth=9, relief=SUNKEN)
    frame4.pack(side=TOP, fill='x')
    frame2 = Frame(root, bg='red', borderwidth=9, relief=SUNKEN)
    frame2.pack(side='top', fill='x')

    frame3 = Frame(root, bg='blue', borderwidth=9, relief=SUNKEN)
    frame3.pack(side=TOP, fill='x')

    global resultbar, srclabel
    srclabel = Label(frame4, borderwidth=9, relief=SUNKEN, text='Search Results', bg='green', textvariable=label_text,
                     font='Arial 10 bold')
    srclabel.pack(side=TOP, fill='x')
    resultbar = Text(frame2, width=60, borderwidth=9, relief=SUNKEN)
    resultbar.StringVar = res
    resultbar.grid(row=0, column=0, ipady=20)
    optionbtn = OptionMenu(frame3, optvar, *optlist, command=lang)
    optionbtn.config(borderwidth=9)
    optionbtn.grid(row=0, column=0)

    btncopy = Button(frame3, borderwidth=9, relief=SUNKEN, text='COPY TEXT', command=copy_text)
    btncopy.grid(row=0, column=1)

    btexit = Button(frame3, borderwidth=9, relief=SUNKEN, text='Exit', command=exit_gui)
    btexit.grid(row=0, column=2)


def result():

    try:
        resultbar.delete(1.0, END)
        global s
        s = w.summary(text.get())
        resultbar.insert(END, s)
        label_text.set('Search Results for : {}'.format(text.get()))
    except :
        messagebox.showwarning('Connection Error','No internet connection')

def lang(event):
    global translated
    global inc
    inc = 0
    if event == 'Hindi':

        inc = 1
        resultbar.delete(1.0, END)
        translated = trans.translate(s, dest='hi')
        resultbar.insert(END, translated)
    elif event == 'German':
        inc = 2
        resultbar.delete(1.0, END)
        translated = trans.translate(s, dest='de')
        resultbar.insert(END, translated)
    elif event == 'English':
        inc = 3
        resultbar.delete(1.0, END)
        translated = trans.translate(s, dest='en')
        resultbar.insert(END, translated)
    elif event == 'Japanese':
        resultbar.delete(1.0, END)
        translated = trans.translate(s, dest='ja')
        resultbar.insert(END, translated)


def voice():
    speaker.speak(resultbar.get(1.0, 'end'))
    label_text.set('Search Results for : {}'.format(text.get()))


def listening():

    with m as souce:
        audio_text = r.listen(souce)
        
        try:
            label_text.set('Listening........')
            t = r.recognize_google(audio_text)
            text.set(t)
            print(t)
            label_text.set('Search Results for : {}'.format(text.get()))

            resultbar.delete(1.0, END)
            resultbar.insert(END, w.summary(t))
        except :
            messagebox.showerror("Speech Error", "Error in listening your voice")


interior()
root.mainloop()
