import tkinter as tk
import customtkinter
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("400x150")
        self.title("PowerPoint Translator")

        self.label=customtkinter.CTkLabel(self,text="PowerPoint Translator")
        self.label.cget("font").configure(size=24)
        self.label.pack()

        self.button = customtkinter.CTkButton(
            self,
            text="➕\nDrag & Drop Here",
            bg_color="blue",
            fg_color="white",
            command=self.import_file
        )
        self.button.pack(expand=True, fill="both", padx=40, pady=20)

        self.button.drop_target_register(DND_FILES)
        self.button.dnd_bind("<<Drop>>", self.drop)

    def import_file(self):
        file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            # Process the selected file (you can replace this with your own logic)
            print("Selected file:", file_path)

    def drop(self, event):
        file_path=event.data
        file_path=file_path.replace("{","")
        file_path=file_path.replace("}","")
        self.button.configure(text=file_path)
        print("Selected file:", file_path)

    def button_callback(self):
        print("button clicked")


app = App()
app.mainloop()