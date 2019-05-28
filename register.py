from tkinter import *
from tkinter import messagebox
import pandas as pd
import re
import datetime


now = datetime.datetime.now()
df = pd.read_excel(r"C:\Users\n189031\Downloads\TCOUG_spring_wksp_attend.xlsx")

first_last_name_arr = df['FirstLastName'].values.tolist()
first_name_arr = df['FirstName'].values.tolist()
last_name_arr = df['LastName'].values.tolist()
company_arr = df['Company'].values.tolist()
email_arr = df['Email'].values.tolist()
renewal_arr = df['RenewalDate'].tolist()
role_arr = df['Role'].tolist()


class AutocompleteEntry(Entry):
    def __init__(self, first_last_name_arr, *args, **kwargs):

        Entry.__init__(self, *args, **kwargs)
        self.first_last_name_arr = first_last_name_arr
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.lb_up = False

    def changed(self, name, index, mode):

        if self.var.get() == '':
            self.lb.destroy()
            self.lb_up = False
        else:
            words = self.comparison()
            if words:
                if not self.lb_up:
                    self.lb = Listbox(width=40)
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
                    self.lb_up = True
                self.lb.delete(0, END)
                for w in words:
                    self.lb.insert(END, w)
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False

    def selection(self, event):
        if self.lb_up:
            self.var.set(self.lb.get(ACTIVE))
            global fln
            fln = (self.lb.get(ACTIVE))
            self.lb.destroy()
            self.lb_up = False
            self.icursor(END)
            #Button(text='Check in', command=check_in_button, height=5, width=20).place(x=300, y=100)

    def up(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':
                self.lb.selection_clear(first=index)
                index = str(int(index)-1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def down(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != END:
                self.lb.selection_clear(first=index)
                index = str(int(index)+1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def comparison(self):
        pattern = re.compile('.*' + self.var.get() + '.*')
        return [w for w in self.first_last_name_arr if re.match(pattern, w)]


def check_in_button():

    def exit_button():
        register_window.configure(bg='SystemMenu')
        dues_check.destroy()

    def callback():
        register_window.configure(bg='SystemMenu')
        dues_check.destroy()

    idx = first_last_name_arr.index(fln)
    #print(idx)
    #print("Hello! " + first_name_arr[idx] + " " + last_name_arr[idx])
    rough_renewal_date = (str(renewal_arr[idx]))
    #print(rough_renewal_date)
    renewal_date = (rough_renewal_date.split(' ')[0])
    #print(renewal_date)
    #print(now.strftime("%Y-%m-%d")
    if renewal_date < now.strftime("%Y-%m-%d"):
        dues_check = Toplevel(register_window)
        dues_check.title(first_name_arr[idx] + " " + last_name_arr[idx] + " is overdue!")
        dues_check.geometry("410x50+500+500")
        dues_check.configure(bg='red')
        register_window.configure(bg='red')
        display = Label(dues_check, text="Renewal Date is: " + renewal_date)
        exit = Button(dues_check, text='EXIT', command=exit_button)
        display.pack()
        exit.pack()
        #print("Hello " + first_name_arr[idx] + " " + last_name_arr[idx] + " you are over due on membership!")
    else:
        dues_check = Toplevel(register_window)
        dues_check.title(first_name_arr[idx] + " " + last_name_arr[idx] + " is current on membership!")
        dues_check.geometry("410x50+500+500")
        dues_check.configure(bg='green')
        register_window.configure(bg='green')
        display = Label(dues_check, text="Renewal Date is: " + renewal_date)
        exit = Button(dues_check, text='EXIT', command=exit_button)
        display.pack()
        exit.pack()
        #print("Hello " + first_name_arr[idx] + " " + last_name_arr[idx] + " you are current on membership!")
    dues_check.protocol("WM_DELETE_WINDOW", callback)


if __name__ == '__main__':

    def callback():
        MsgBox = messagebox.askquestion('Exit TCOUG Registration', 'Are you sure you want to exit the '
                                                                            'application', icon='warning')
        if MsgBox == 'yes':
            register_window.destroy()
        else:
            messagebox.showinfo('Return', 'You will now return to TCOUG registration')

    register_window = Tk()
    register_window.title("TCOUG Registration")
    register_window.geometry("500x215+450+450")
    tcoug = PhotoImage(file=r"C:\Users\n189031\Downloads\TCOUGLogo.png")
    tcouglabel = Label(image=tcoug)
    tcouglabel.place(x=280, y=15)

    entry = AutocompleteEntry(first_last_name_arr, register_window, width=40).place(x=10, y=15)
    Button(text='Check in', command=check_in_button, height=5, width=20).place(x=300, y=100)
    register_window.protocol("WM_DELETE_WINDOW", callback)

    register_window.mainloop()
