import openpyxl
import keyboard
import pyautogui
import threading
import logging
from tkinter import *
from tkinter import ttk, messagebox, filedialog

# تنظیمات لاگ
logging.basicConfig(
    filename='macro_log.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Macro Manager")
        self.root.geometry("800x600")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.is_running = False
        self.commands = {}
        self.current_file = "Commands.xlsx"
        self.hotkeys = []

        self.create_widgets()
        self.load_commands()

    def create_widgets(self):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel", command=self.open_file)
        file_menu.add_command(label="Reload", command=self.reload_commands)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)

        toolbar = Frame(self.root, bd=1, relief=RAISED)
        toolbar.pack(side=TOP, fill=X)

        self.start_btn = Button(toolbar, text="▶️ Start", command=self.start_macro)
        self.start_btn.pack(side=LEFT, padx=2, pady=2)

        self.stop_btn = Button(toolbar, text="⏹ Stop", command=self.stop_macro, state=DISABLED)
        self.stop_btn.pack(side=LEFT, padx=2, pady=2)

        self.tree = ttk.Treeview(self.root, columns=('Key', 'Action', 'Type'), show='headings')
        self.tree.heading('Key', text='Key')
        self.tree.heading('Action', text='Action')
        self.tree.heading('Type', text='Type')
        self.tree.column('Key', width=100)
        self.tree.column('Action', width=300)
        self.tree.column('Type', width=100)
        self.tree.pack(fill=BOTH, expand=True, padx=5, pady=5)

        self.status = Label(self.root, text="Ready", bd=1, relief=SUNKEN, anchor=W)
        self.status.pack(side=BOTTOM, fill=X)

    def update_status(self, text):
        self.status.config(text=text)
        self.status.update_idletasks()

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.current_file = file_path
            self.load_commands()

    def reload_commands(self):
        self.load_commands()

    def load_commands(self):
        try:
            self.commands = {}
            wb = openpyxl.load_workbook(self.current_file)
            sheet = wb.active

            for item in self.tree.get_children():
                self.tree.delete(item)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                key = row[0]
                action = row[1]
                action_type = row[2] if len(row) > 2 else "write"
                if key:
                    self.commands[key] = (action, action_type)
                    self.tree.insert('', 'end', values=(key, action, action_type))

            self.update_status(f"Commands loaded successfully from {self.current_file}")
            logging.info(f"Commands loaded from {self.current_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")
            logging.error(f"Excel load error: {str(e)}")

    def start_macro(self):
        if not self.is_running:
            self.is_running = True
            self.start_btn.config(state=DISABLED)
            self.stop_btn.config(state=NORMAL)

            for key, (action, action_type) in self.commands.items():
                hotkey = keyboard.add_hotkey(key, lambda k=key, a=action, t=action_type: self.execute_action(k, a, t), suppress=True)
                self.hotkeys.append(hotkey)

            self.update_status("Macro running...")

    def stop_macro(self):
        self.is_running = False
        for hotkey in self.hotkeys:
            try:
                keyboard.remove_hotkey(hotkey)
            except Exception as e:
                logging.warning(f"Failed to remove hotkey: {e}")
        self.hotkeys.clear()
        self.start_btn.config(state=NORMAL)
        self.stop_btn.config(state=DISABLED)
        self.update_status("Macro stopped")

    def execute_action(self, key, action, action_type):
        try:
            if action_type == "write":
                keyboard.write(action)
            elif action_type == "click":
                pyautogui.click()
            elif action_type == "hotkey":
                keyboard.send(action)
            elif action_type == "script":
                exec(action)
            logging.info(f"Executed: {key} -> {action} ({action_type})")
        except Exception as e:
            logging.error(f"Execution error for {key}: {str(e)}")

    def on_close(self):
        self.stop_macro()
        self.root.destroy()


if __name__ == "__main__":
    root = Tk()
    app = MacroApp(root)
    root.mainloop()
