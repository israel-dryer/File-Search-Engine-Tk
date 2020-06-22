"""
    File Search Engine
    Author: Israel Dryer
    Modified: 2020-06-21
"""

import csv
import pathlib
import subprocess
import datetime
from queue import Queue
from threading import Thread
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askdirectory, asksaveasfilename
import ttkthemes

class Engine(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('File Search Engine')
        self.style = ttkthemes.ThemedStyle()
        self.style.theme_use('arc')
        self.withdraw()
        self.iconbitmap('icon.ico')
        self.wm_state('zoomed')
        self.platform = self.tk.call('tk', 'windowingsystem')
        
        # application variables
        self.search_type_var = tk.StringVar()
        self.search_term_var = tk.StringVar()
        self.search_path_var = tk.StringVar()        

        # window frames
        self.frm_main = ttk.Frame(self, padding=10)
        self.frm_top = ttk.Frame(self.frm_main)
        self.frm_term = ttk.Frame(self.frm_top)
        self.frm_type = ttk.LabelFrame(self.frm_top, text='Type', padding=4)
        self.frm_path = ttk.Frame(self.frm_top)

        # search path input
        self.path_lbl = ttk.Label(self.frm_path, text='Path', width=6)
        self.path_entry = ttk.Entry(self.frm_path, textvariable=self.search_path_var, width=60)

        # search term input
        self.term_lbl = ttk.Label(self.frm_term, text='Term', width=6)
        self.term_entry = ttk.Entry(self.frm_term, textvariable=self.search_term_var, width=60)
        self.term_entry.focus()

        # search type selection
        self.type_contains = ttk.Radiobutton(self.frm_type, text='Contains', value='contains', variable=self.search_type_var)
        self.type_startswith = ttk.Radiobutton(self.frm_type, text='StartsWith', value='startswith', variable=self.search_type_var)
        self.type_endswith = ttk.Radiobutton(self.frm_type, text='EndsWith', value='endswith', variable=self.search_type_var)

        # form buttons
        self.btn_browse = ttk.Button(self.frm_top, text='Browse', command=self.on_browse)
        self.btn_search = ttk.Button(self.frm_top, text='Search', command=self.on_search)

        # search results tree
        self.tree = ttk.Treeview(self.frm_main)
        self.tree['columns'] = ('modified', 'type', 'size', 'path')
        self.tree.column('#0', width=400)
        self.tree.column('modified', width=150, stretch=False, anchor=tk.E)
        self.tree.column('type', width=50, stretch=False, anchor=tk.E)
        self.tree.column('size', width=50, stretch=False, anchor=tk.E)        
        self.tree.heading('#0', text='Name')
        self.tree.heading('modified', text='Modified date')
        self.tree.heading('type', text='Type')
        self.tree.heading('size', text='Size')
        self.tree.heading('path', text='Path')      

        # progress bar
        self.prog_bar = ttk.Progressbar(self.frm_main, orient=tk.HORIZONTAL, mode='indeterminate')

        # right-click menu for treeview
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label='Reveal in file manager', command=self.on_doubleclick_tree)
        self.menu.add_command(label='Export results to csv', command=self.export_to_csv)
        
        # add widgets to window
        self.frm_top.pack(side=tk.TOP, fill=tk.X)
        self.frm_path.grid(row=0, column=0, padx=4, pady=2, sticky=tk.NSEW)
        self.path_lbl.pack(side=tk.LEFT)
        self.path_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        self.btn_browse.grid(row=0, column=1, padx=4, pady=2, sticky=tk.NSEW)
        self.frm_term.grid(row=1, column=0, padx=4, pady=2, sticky=tk.NSEW)
        self.term_lbl.pack(side=tk.LEFT, )
        self.term_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        self.btn_search.grid(row=1, column=1, padx=4, pady=2, sticky=tk.NSEW)        
        self.frm_type.grid(row=0, column=2, rowspan=2, padx=4, pady=2, sticky=tk.NSEW)
        self.type_contains.grid(row=0, column=0, sticky=tk.W)
        self.type_startswith.grid(row=2, column=0, sticky=tk.W)
        self.type_endswith.grid(row=1, column=0, sticky=tk.W)
        self.tree.pack(side=tk.TOP, padx=4, pady=4, fill=tk.BOTH, expand=tk.YES)
        self.prog_bar.pack(side=tk.TOP, padx=5, pady=2, fill=tk.X)
        self.frm_main.pack(fill=tk.BOTH, expand=tk.YES)

        # event binding
        self.tree.bind('<Double-1>', self.on_doubleclick_tree)

        if self.platform == 'aqua':
            self.tree.bind('<2>', self.right_click_tree)
            self.tree.bind('<Control-1>', self.right_click_tree)
        else:
            self.tree.bind('<3>', self.right_click_tree)

        # intial settings
        self.search_type_var.set('contains')
        self.search_path_var.set(pathlib.os.getcwd())
        self.search_count = 0
        self.deiconify()

    def on_browse(self):
        """Callback for directory browse"""
        path = askdirectory(title='Directory')
        if path:
            self.search_path_var.set(path)
            self.update_idletasks()

    def on_doubleclick_tree(self, event=None):
        """Callback for double-click tree menu"""
        try:
            id = self.tree.selection()[0]
        except IndexError:
            return
        if id.startswith('I'):
            self.reveal_in_explorer(id)

    def right_click_tree(self, event=None):
        """Callback for right-click tree menu"""
        try:
            id = self.tree.selection()[0]
        except IndexError:
            return
        if id.startswith('I'):
            self.menu.entryconfigure('Export results to csv', state=tk.DISABLED)
            self.menu.entryconfigure('Reveal in file manager', state=tk.NORMAL)
        else:
            self.menu.entryconfigure('Export results to csv', state=tk.NORMAL)
            self.menu.entryconfigure('Reveal in file manager', state=tk.DISABLED)
        self.menu.post(event.x_root, event.y_root)

    def on_search(self):
        """Search for a term based on the search type"""
        search_term = self.search_term_var.get()
        search_path = self.search_path_var.get()
        search_type = self.search_type_var.get()
        if search_term == '':
            return
        Thread(target=file_search, args=(search_term, search_path, search_type), daemon=True).start()
        self.prog_bar.start(10)
        self.search_count += 1
        id = self.tree.insert('', 'end', self.search_count, text=f'Search {self.search_count}')
        self.tree.item(id, open=True)
        self.check_queue(id)

    def reveal_in_explorer(self, id):
        """Callback for double-click event on tree"""
        values = self.tree.item(id, 'values')
        path = pathlib.Path(values[-1]).absolute().parent
        pathlib.os.startfile(path)

    def export_to_csv(self, event=None):
        """Export values to csv file"""
        try:
            id = self.tree.selection()[0]
        except IndexError:
            return

        filename = asksaveasfilename(initialfile='results.csv', filetypes=[('Comma-separated', '*.csv'), ('Text', '*.txt')])
        if filename:
            with open(filename, mode='w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Name', 'Modified date', 'Type', 'Size', 'Path'])
                children = self.tree.get_children(id)
                for child in children:
                    name = [self.tree.item(child, 'text')]
                    values = list(self.tree.item(child, 'values'))
                    writer.writerow(name + values)
        # open file in default program
        # pathlib.os.startfile(filename) # is this the desired behavior?

    def check_queue(self, id):
        """Check file queue and print results if not empty"""
        if searching and not file_queue.empty():
            filename = file_queue.get()
            self.insert_row(filename, id)
            self.update_idletasks()
            self.after(1, lambda: self.check_queue(id))
        elif not searching and not file_queue.empty():
            while not file_queue.empty():
                filename = file_queue.get()
                self.insert_row(filename, id)
            self.update_idletasks()
            self.prog_bar.stop()
        elif searching and file_queue.empty():
            self.after(100, lambda: self.check_queue(id))
        else:
            self.prog_bar.stop()

    def insert_row(self, file, id):
        """Insert new row in tree search results"""
        try:
            file_stats = file.stat()
            file_name = file.stem
            file_modified = datetime.datetime.fromtimestamp(file_stats.st_mtime).strftime('%m/%d/%Y %I:%M:%S%p')
            file_type = file.suffix.lower()
            file_size = convert_size(file_stats.st_size)
            file_path = file.absolute()
            self.tree.insert(id, 'end', text=file_name, values=(file_modified, file_type, file_size, file_path))
        except OSError:
            return

def file_search(term, search_path, search_type):
    """Recursively search directory for matching files"""
    set_searching(1)
    if search_type == 'contains':
        find_contains(term, search_path)
    elif search_type == 'startswith':
        find_startswith(term, search_path)
    elif search_type == 'endswith':
        find_endswith(term, search_path)            

def find_contains(term, search_path):
    """Find all files that contain the search term"""
    for path, _, files in pathlib.os.walk(search_path):
        if files:
            for file in files:
                if term in file:
                    file_queue.put(pathlib.Path(path) / file)
    set_searching(False)

def find_startswith(term, search_path):
    """Find all files that start with the search term"""
    for path, _, files in pathlib.os.walk(search_path):
        if files:
            for file in files:
                if file.startswith(term):
                    file_queue.put(pathlib.Path(path) / file)
    set_searching(False)

def find_endswith(term, search_path):
    """Find all files that end with the search term"""
    for path, _, files in pathlib.os.walk(search_path):
        if files:
            for file in files:
                if file.endswith(term):
                    file_queue.put(pathlib.Path(path) / file)
    set_searching(False)

def set_searching(state=False):
    """Set searching status"""
    global searching
    searching = state

def convert_size(size):
    """Convert bytes to mb or kb depending on scale"""
    kb = size // 1000
    mb = round(kb / 1000, 1)
    if kb > 1000:
        return f'{mb:,.1f} MB' 
    else:
        return f'{kb:,d} KB'  

           
if __name__ == '__main__':
    
    file_queue = Queue()
    searching = False
    app = Engine()
    app.mainloop()