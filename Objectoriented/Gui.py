import tkinter as tk

window = tk.Tk()
window.title('LOGEFI SERVICES')
window.geometry('500x800')
paris02b = tk.Button(window, text='Paris-02-法语授权书')
paris02b.pack()
paris16b = tk.Button(window, text='Paris-16-法语授权书')
paris16b.pack()
cscb = tk.Button(window, text='Châtillon-sur-Cluses法语授权书')
cscb.pack()
changeparis02b = tk.Button(window, text='税务代表变更授权书 Paris 02')
changeparis02b.pack()

window.mainloop()
