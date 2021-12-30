# Import the required libraries
import os
from tkinter import *
from tkinter import ttk
import tkinter.filedialog as fd

# Create an instance of tkinter frame or window
root = Tk()
list = []
# Set the geometry of tkinter frame
root.geometry("700x350")

# file = fd.askopenfilenames(parent=win, title='Choose a File')
# #    print(win.splitlist(file))
# list = win.splitlist(file)

# for item in list:
#     head_tail = os.path.split(item)
#     print(head_tail[0] + '/_' + head_tail[1][:-5] + '.' + head_tail[1][-5:])
#     print(head_tail[1][-5:])
def show_text():
    label_text.set('Heh, heh, heh, you said, "' + entry_text.get() + '"')
    print(int(f'{entry_text.get()}') + 10)

entry_text = StringVar()
entry = Entry(root, width=10, textvariable=entry_text)
entry.pack()

button = Button(root, text="Click Me", command=show_text)
button.pack()

label_text = StringVar()
label = Label(root, textvariable=label_text)
label.pack()

root.mainloop()