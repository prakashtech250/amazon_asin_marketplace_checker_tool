import requests
import os
from sys import platform
import pandas as pd
from datetime import datetime
from openpyxl import Workbook

from scrapy.selector import Selector
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter import messagebox

from threading import Thread

class amazonApi:
    def __init__(self):
        global running
        running = False
        self.filename = None
        self.excel_filename = None
        self.sheet_title = []

    def init_UI(self):
        root = tk.Tk()
        root.title('Amazon Asin Market Checker Tool')
        root.geometry('600x400')
        icon = tk.PhotoImage(file='./icon.png')
        root.iconphoto(False,icon)
        domain_list = self.get_domains()

        frame1 = tk.Frame(master=root)
        frame1.pack(fill=tk.X,padx=10, pady=2)
        
        fontStyle = ("Lucida Grande", 10)
        label1 = tk.Label(frame1, text='Enter Asin: ',font=fontStyle)
        label1.pack(side=tk.LEFT)

        self.entry1 = tk.Entry(frame1, width=40) 
        self.entry1.pack(side=tk.LEFT)

        open_button = tk.Button(
            frame1,
            text='Select a File',
            command=self.select_file
        )

        open_button.pack(side=tk.LEFT)

        self.check_button = tk.Button(
            frame1,
            text='Start',
            width=10,
            bg='green',
            command=self.start
        )
        self.check_button.pack(side=tk.RIGHT)

        output_frame = tk.Frame(root)
        v_scrollbar = tk.Scrollbar(output_frame, orient=tk.VERTICAL, jump=1)
        h_scrollbar = tk.Scrollbar(output_frame, orient=tk.HORIZONTAL, jump=1)

        self.my_tree = ttk.Treeview(output_frame,yscrollcommand=v_scrollbar.set,xscrollcommand=h_scrollbar.set)

        v_scrollbar.config(command=self.my_tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar.config(command=self.my_tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        all_columns = ['count','asin','description'] + domain_list
        self.my_tree['columns'] = all_columns
        
        self.my_tree.column("#0",width=0, stretch=0)
        self.my_tree.column('count', anchor=tk.W, width=40, minwidth=40)
        self.my_tree.column('asin', anchor=tk.W, width=100, minwidth=100)
        self.my_tree.column('description',anchor=tk.W,width=120, minwidth=50)
        for domain in domain_list:
            width = len(domain)*10
            self.my_tree.column(domain,anchor=tk.W,width=width, minwidth=width)

        self.my_tree.heading('#0', text='', anchor=tk.W)
        self.my_tree.heading('count', text="ID", anchor=tk.W)
        self.my_tree.heading('asin', text='ASIN', anchor=tk.CENTER)
        self.my_tree.heading('description', text='Description',anchor=tk.W)
        for domain in domain_list:
            self.my_tree.heading(domain,text=domain,anchor=tk.W)
 

        self.my_tree.pack(fill=tk.BOTH, expand=True)
        output_frame.pack(fill=tk.BOTH, expand=True)

        my_menu = tk.Menu(root)
        root.config(menu=my_menu)
        
        about_text = 'Amazon Asin Marketplace Checker Tool \nDevloper: Prakash \nContact: https://www.fiverr.com/prakashtech250'
        my_menu.add_command(label="About", command=lambda: self.print_information(about_text))
        my_menu.add_command(label="Exit", command=lambda: self.exit_window(root))
        root.mainloop()
    
    def exit_window(self, root):
        global running
        if messagebox.askokcancel("Exit", "Do you want to Exit?"):
            root.destroy()
            if running:
                self.driver.quit()

    
    def print_information(self, message):
        messagebox.showinfo("Information",message)

    def error_information(self, message):
        messagebox.showwarning("Warning",message)

    def open_browser(self) :
        if platform == 'linux':
            driver_path = '{}/driver/geckodriver_linux'.format(os.getcwd())
        elif platform == "darwin":
            driver_path = '{}/driver/geckodriver_mac'.format(os.getcwd())
        elif platform == 'win32' or platform == 'win':
            driver_path = '{}/driver/geckodriver_win.exe'.format(os.getcwd())
        print('opening browser. Please wait.....')
        options = Options()
        options.add_argument("--incognito")
        options.headless = True
        self.driver = webdriver.Firefox(options=options,executable_path=r'{}'.format(driver_path))

    def get_page_source(self,url):
        try:
            self.driver.get(url)
            page_source = self.driver.page_source
            res = Selector(text=page_source)
            return res
            
        except Exception as e:
            print('Error: {}'.format(e))
    
    def get_domains(self):
        with open('domains.txt','r') as f:
            return f.read().split('\n')
    
    def create_filename(self):
        if not os.path.exists('output'):
            os.makedirs('output')
        now = datetime.now()
        dt_string = now.strftime("%Y%m%d_%H_%M")
        filename = 'output/{}.xlsx'.format(dt_string)
        return filename
    
    def select_file(self):
        filetypes = (
            ('text files', '*.txt'),
            ('All files', '*.*')
        )

        cwd = os.getcwd()
        filename = fd.askopenfilename(
            title='Oepn a file',
            initialdir=cwd,
            filetypes=filetypes)
        
        if filename:
            try:
                self.filename = r"{}".format(filename)
                self.setTextInput(filename)
            except ValueError:
                print('File couldn\'t be open')
        else:
            self.filename = None


    def setTextInput(self,text):
        self.entry1.delete(0,"end")
        self.entry1.insert(0, text)
    
    def insert(self, data):
        self.my_tree.insert(parent='', index='end', iid=self.iid, text='', values=data)

    
    def delete_data(self, iid):
        self.my_tree.delete(iid)

    def delete_all_data(self):
        x=self.my_tree.get_children()
        for item in x:
            self.my_tree.delete(item)

    def run(self, asin):
        domain_list = self.get_domains()
        start_data = [self.id,asin]
        description = []
        presence_asin = []
        description_index = 0
        for domain in domain_list:    
            if running:
                self.insert(start_data+description+presence_asin)
                url = 'https://www.{}/dp/{}'.format(domain,asin)
                response = self.get_page_source(url)
                title = response.css('#productTitle::text').get()
                if title and running:
                    if description_index == 0:
                        title = title.strip()
                        description.append(title)
                        description_index+=1
                    presence_asin.append('Yes')
                else:
                    presence_asin.append('No')
                self.delete_data(self.id)
                
        data = start_data + description + presence_asin
        self.insert(data)  
        return data[1:] 
            
    def get_asins(self):
        get_asin = self.get_asin()
        if self.filename:
            with open(self.filename, 'r') as f:
                asin_list = f.read().split('\n')
            return asin_list
        elif get_asin:
            return get_asin
        else:
            self.error_information('Field is empty. Select the file or enter asin code.')

    def get_asin(self):
        asin_list = []
        asin = self.entry1.get()
        if asin is not None:
            asin_list.append(asin)
        return asin_list

    def start(self):
        global running 
        if running == False:
            self.id = 0
            self.iid = 0
            self.delete_all_data()
            t1 = Thread(target=self.main)
            t1.start()
        else:
            self.error_information('Script is running. Please wait.')

    def stop(self):
        global running
        if messagebox.askokcancel("Stop", "Do you want to Stop?"):
            if running:
                running = False
                self.filename = None
                self.check_button.config(text='Start',bg='green',command=self.start)
                self.driver.quit()
    
    def finish_task(self, filename):
        global running
        running = False
        self.filename = None
        self.driver.quit()
        self.check_button.config(text='Start',bg='green',command=self.start)
        self.print_information('Task is completed. \nOutput is saved as {}'.format(filename))

    def get_sheet_title(self):
        sheet_title = ['Asin','Description']
        domain_list = self.get_domains()
        for domain in domain_list:
            sheet_title.append(domain)
        return sheet_title
  
    def main(self):
        global running
        asin_list = []
        asin_list = self.get_asins()  
        if len(asin_list[0]) > 0:
            excel_filename = self.create_filename()
            sheet_title = self.get_sheet_title()
            wb = Workbook()
            ws = wb.active
            ws.append(sheet_title)
            self.check_button.config(text='Stop',bg='Red',command=self.stop)
            running = True
            self.open_browser()
            for asin in asin_list:
                if running:
                    try:
                        data = self.run(asin)
                        ws.append(data)
                        wb.save(excel_filename)
                    except Exception as e:
                        print('Error: {}'.format(e))
                    self.iid += 1
                    self.id += 1 
            self.finish_task(excel_filename)
        else:
            self.error_information('Open a file or enter any asin!!!')


if __name__=='__main__':
    api = amazonApi()
    api.init_UI()
    

    

