import pandas as pd
from pynput import keyboard
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import os

class KeyloggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Keylogger")
        
        self.is_logging = False
        self.keys = []
        
        self.start_button = tk.Button(root, text="Start", command=self.start_logging)
        self.start_button.pack(pady=10)
        
        self.stop_button = tk.Button(root, text="Stop", command=self.stop_logging)
        self.stop_button.pack(pady=10)
        
        self.status_label = tk.Label(root, text="Status: Not Logging")
        self.status_label.pack(pady=10)
        
        self.log_listener = None
        
    def on_press(self, key):
        try:
            self.keys.append(key.char)
        except AttributeError:
            self.keys.append(f"[{key}]")
    
    def start_logging(self):
        if not self.is_logging:
            self.is_logging = True
            self.status_label.config(text="Status: Logging")
            self.log_listener = keyboard.Listener(on_press=self.on_press)
            self.log_listener.start()
        else:
            messagebox.showinfo("Info", "Logging is already started.")
    
    def stop_logging(self):
        if self.is_logging:
            self.is_logging = False
            self.status_label.config(text="Status: Not Logging")
            self.log_listener.stop()
            
            # Save to Excel
            self.save_to_excel()
            self.keys = []
        else:
            messagebox.showinfo("Info", "Logging is not started.")
    
    def save_to_excel(self):
        now = datetime.now()
        filename = now.strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"
        df = pd.DataFrame({"Keystrokes": self.keys})
        
        if not os.path.exists("logs"):
            os.makedirs("logs")
        
        filepath = os.path.join("logs", filename)
        df.to_excel(filepath, index=False)
        messagebox.showinfo("Info", f"Log saved to {filepath}")

if __name__ == "__main__":
    root = tk.Tk()
    app = KeyloggerApp(root)
    root.mainloop()
