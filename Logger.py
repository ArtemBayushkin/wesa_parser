import logging
import tkinter as tk


class GUILogHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, f"{msg}\n")
        self.text_widget.see(tk.END)