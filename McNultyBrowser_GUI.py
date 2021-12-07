import threading
import time
import queue
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from McNultyBrowser import McNultyBrowser

class Login_screen:

    def __init__(self,window):
        self.master = window
        self.br = McNultyBrowser()
        self.make_login()
        self.remember = False

    def make_login(self):
        self.login_frame = tk.Frame(master=self.master)
        self.username = tk.Entry(master = self.login_frame, width = 25)
        self.password = tk.Entry(master = self.login_frame, width = 25)
        username_label = tk.Label(master = self.login_frame, text = "Username:")
        password_label = tk.Label(master = self.login_frame, text = "Password:")
        self.login_button = tk.Button(master = self.login_frame, text = "LOGIN", command = self.login, width = 20)
        self.login_frame.pack(fill=tk.BOTH, expand=1)
        username_label.grid(row = 1, column = 1, padx = 10, pady = 10)
        password_label.grid(row = 1, column = 3, padx = 10, pady = 10)
        self.username.grid(row = 1, column = 2, padx = 10, pady = 10)
        self.password.grid(row = 1, column = 4, padx = 10, pady = 10)
        self.login_button.grid(row = 2, column = 2, padx = 10, pady = 10, columnspan = 4)
        for i in range(0,5):
            self.login_frame.grid_columnconfigure(i,weight=1)
        for i in range(0,3):
            self.login_frame.grid_rowconfigure(i,weight=1)

    def progress(self):
        self.prog_bar = ttk.Progressbar(
            master = self.login_frame, orient="horizontal",
            length=200, mode="indeterminate"
            )
        self.prog_bar.grid(row = 2, column = 2)

    def process_queue(self):
        try:
            msg = self.queue.get(0)
            self.prog_bar.stop()
            self.prog_bar.destroy()
            self.login_button.grid(row = 2, column = 2, padx = 10, pady = 10, columnspan = 4)
            if msg == 'done':
                self.master.destroy()
                search_window = tk.Tk()
                search_window.title("McNulty 4-Column Formatter and Poster")
                for i in range(0,4):
                    search_window.grid_columnconfigure(i,weight=1)
                for i in range(0,4):
                    search_window.grid_rowconfigure(i,weight=1)
                Search_App(search_window, self.br)
                search_window.mainloop()
            else:
                messagebox.showerror('Small Problem...','Invalid Username or Password')
                self.login_button['state'] = 'active'
        except queue.Empty:
            self.prog_bar.step(10)
            self.master.after(100, self.process_queue)

    def login(self):
        self.login_button['state'] = 'disabled'
        self.login_button.grid(row = 2, column = 3, padx = 10, pady = 10, columnspan = 4)
        self.progress()
        self.prog_bar.start()
        self.queue = queue.Queue()
        thread_task_login(self.queue,self.br,self.username.get(),self.password.get()).start()
        self.process_queue()


class Search_App:

    def __init__(self,window,br):
        self.master = window
        self.br = br
        self.index_list = []*6
        self.index = 0
        self.info = []
        self.make_search()
        self.make_results()

    def set_input(self,Company_Name,Location1,Location2,CAGE,DACIS,WRAP_Type):
        self.results_box.delete('1.0',tk.END)
        self.results_box.insert('1.0','Name: ' + Company_Name + '\n')
        self.results_box.insert('2.0','Location: ' + Location1 + '\n')
        self.results_box.insert('3.0', Location2 + '\n')
        self.results_box.insert('4.0', 'CAGE: ' + CAGE + '\n')
        self.results_box.insert('5.0','DACIS: ' + DACIS + '\n')
        self.results_box.insert('6.0','Type: ' + WRAP_Type + '\n')

    def change_data_add(self):
        self.index+=1
        if self.index > len(self.br.link_info)-1:
            self.index = len(self.br.link_info)-1
        elif self.index < 0:
            self.index = 0
        self.info = self.br.link_info[self.index]
        self.set_input(self.info[0],self.info[3],self.info[2],self.info[8],self.info[7],self.info[4])

    def change_data_sub(self):
        self.index-=1
        if self.index > len(self.br.link_info)-1:
            self.index = len(self.br.link_info)-1
        elif self.index < 0:
            self.index = 0
        self.info = self.br.link_info[self.index]
        self.set_input(self.info[0],self.info[3],self.info[2],self.info[8],self.info[7],self.info[4])

    def progress(self,master):
        self.prog_bar = ttk.Progressbar(
            master = master, orient="horizontal",
            length=200, mode="indeterminate"
            )
        self.prog_bar.pack(side=tk.TOP)

    def process_queue(self):
        try:
            msg = self.queue.get(0)
            self.prog_bar.stop()
            self.prog_bar.destroy()
            self.index = 0
            if len(self.br.link_info) == 0:
                self.results_box.delete('1.0',tk.END)
                self.results_box.insert('1.0', 'No results')
                self.search_button['state'] = 'active'
                self.generate_button['state'] = 'active'
            else:
                self.info = self.br.link_info[self.index]
                self.set_input(self.info[0],self.info[3],self.info[2],self.info[8],self.info[7],self.info[4])
                self.search_button['state'] = 'active'
                self.generate_button['state'] = 'active'
        except queue.Empty:
            self.prog_bar.step(10)
            self.master.after(100, self.process_queue)

    def search_and_display(self):
        self.search_button['state'] = 'disabled'
        self.progress(self.search_frame)
        self.prog_bar.start()
        CAGEs = self.search_bar.get().split(',')
        if CAGEs[0] == '':
            self.results_box.delete('1.0',tk.END)
            self.results_box.insert('1.0','Must enter valid CAGE code to search.')
            self.prog_bar.stop()
            self.prog_bar.destroy()
            self.search_button['state'] = 'active'
        else:
            self.queue = queue.Queue()
            thread_task_search(self.queue,self.br,CAGEs).start()
            self.process_queue()

    def make_search(self):
        self.search_frame = tk.Frame(master=self.master, relief = 'ridge')
        search_frame_label = tk.Label(master=self.search_frame,text="WRAP Calculator")
        search_frame_label.config(font = ("Helvetica",16))
        search_frame_label.pack(pady = 25,padx=10)
        search_label = tk.Label(master = self.search_frame, text = "Search Prime Mover by CAGE, DACIS, DUNS, etc.")
        self.search_bar = tk.Entry(master = self.search_frame, width = 50)
        self.search_button = tk.Button(master=self.search_frame,text='Search',command = self.search_and_display)
        self.search_button.config(height = 4,width = 5)
        self.search_frame.grid(row = 1, column = 1, pady = 5, padx =5)
        self.search_button.pack(side = tk.RIGHT,fill=tk.BOTH, expand=1)
        search_label.pack(padx=5, pady=2)
        self.search_bar.pack(padx=20, pady=5)
        for i in range(0,5):
            self.search_frame.grid_columnconfigure(i,weight=1)
        for i in range(0,3):
            self.search_frame.grid_rowconfigure(i,weight=1)

    def generate_WRAP(self):
        self.queue = queue.Queue()
        self.generate_button['state'] = 'disabled'
        self.progress(self.submission_frame)
        self.prog_bar.start()
        thread_task_generate(self.queue,self.br).start()
        self.process_queue()

    def add_to_list(self):
        self.br.final_link_list.append(self.br.link_list[self.index])
        info = self.info[0] + ' ' + self.info[8] + ' ' + self.info[4] + '\n'
        self.to_do_box.insert(tk.END,info)

    def clear(self):
        self.br.final_link_list = []
        self.to_do_box.delete('1.0',tk.END)

    def make_results(self):
        results_frame = tk.Frame(master=self.master, relief = 'ridge')
        self.submission_frame = tk.Frame(master=self.master, relief = 'ridge')
        results_label = tk.Label(master = results_frame,text = "Results")
        scroll_button_right = tk.Button(master = results_frame, text = '>>', command = self.change_data_add)
        scroll_button_left = tk.Button(master = results_frame, text = '<<', command = self.change_data_sub)
        self.generate_button = tk.Button(master = self.submission_frame, text = 'Generate WRAPs', command = self.generate_WRAP)
        add_to_list_button = tk.Button(master = results_frame, text = 'Add WRAP', command = self.add_to_list)
        clear_list_button = tk.Button(master = self.submission_frame, text = 'CLEAR', command = self.clear)
        self.results_box = tk.Text(master = results_frame, width=50, height=7)
        self.to_do_box = tk.Text(master = self.submission_frame, width=50, height=7)
        results_frame.grid(row=2,column=1,pady=5,padx = 5)
        self.submission_frame.grid(row=3,column=1,pady=5,padx=5)
        results_label.pack(padx=10, pady=2)
        self.results_box.pack(padx=10, pady=5)
        scroll_button_left.pack(padx=10, pady=5,side = tk.LEFT)
        scroll_button_right.pack(padx=10, pady=5,side = tk.RIGHT)
        add_to_list_button.pack(padx=10, pady=5, side = tk.BOTTOM)
        self.to_do_box.pack(padx=10, pady=5)
        clear_list_button.pack(padx=10, pady=5)
        self.generate_button.pack(padx=10, pady=5, side = tk.BOTTOM)
        for i in range(0,5):
            results_frame.grid_columnconfigure(i,weight=1)
        for i in range(0,3):
            results_frame.grid_rowconfigure(i,weight=1)

class thread_task_search(threading.Thread):

    def __init__(self,queue,br,CAGEs):
        super().__init__()
        self.br = br
        self.CAGEs = CAGEs
        self.queue = queue

    def run(self):
        self.br.link_list = []
        self.br.link_info = []
        for CAGE in self.CAGEs:
            self.br.search(CAGE)
        self.queue.put('done')

class thread_task_login(threading.Thread):

    def __init__(self,queue,br,username,password):
        super().__init__()
        self.br = br
        self.queue = queue
        self.username = username
        self.password = password

    def run(self):
        try:
            self.br.login(self.username.strip(),self.password.strip())
            self.queue.put('done')
        except ValueError:
            self.queue.put('invalid')

class thread_task_generate(threading.Thread):

        def __init__(self,queue,br):
            super().__init__()
            self.br = br
            self.queue = queue

        def run(self):
            for link in self.br.final_link_list:
                self.br.get_WRAP(link)
            self.queue.put('done')

if __name__ == '__main__':
    login_window = tk.Tk()
    login_window.title("PrimeMover Login")
    Login_screen(login_window)
    login_window.mainloop()
