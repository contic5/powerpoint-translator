import tkinter as tk
import customtkinter
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog

from pptx import Presentation
from deep_translator import GoogleTranslator
import time
import os

MAX_SHAPES_TO_TRANSLATE=5
TARGET_LANGUAGE="Spanish"

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

        self.results_label=customtkinter.CTkLabel(self,text="Upload a Powerpoint file to translate it.")
        self.results_label.pack()

        self.prs=None
        self.shapes_translated=0
        self.total_shapes=0

        self.slide_index=0
        self.total_slides=0

        self.file_name=""
        self.out_file=""

    def import_file(self):
        file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("PowerPoint Files", "*.pptx"),])
        if file_path:
            # Process the selected file (you can replace this with your own logic)
            print("Selected file:", file_path)
            self.translate_powerpoint(file_path)

    def drop(self, event):
        file_path=event.data
        file_path=file_path.replace("{","")
        file_path=file_path.replace("}","")
        self.button.configure(text=file_path)
        print("Selected file:", file_path)
        self.translate_powerpoint(file_path)

    def get_total_shapes(self):
        total_shapes=0
        for slide in self.prs.slides:
            for shape in slide.shapes:
                total_shapes+=1
        return total_shapes

    def update_shape(self):
        slide=self.prs.slides[self.slide_index]
        shape=slide.shapes[self.shape_index]
        shape_was_translated=False

        try:
            translated_text=GoogleTranslator(source='auto', target='es').translate(shape.text)
            shape.text=translated_text
            shape_was_translated=True
        except AttributeError:
            pass

        status_text=f'''
        Progress
        Slides: {self.slide_index+1}/{self.total_slides}
        Shapes: {self.shapes_translated}/{self.total_shapes}
        '''
        self.results_label.configure(text=status_text)

        self.shape_index+=1
        self.shapes_translated+=1
        if self.shape_index>=len(slide.shapes):
            self.slide_index+=1
            self.shape_index=0
        
        if self.slide_index<len(self.prs.slides) and self.shapes_translated<MAX_SHAPES_TO_TRANSLATE:
            if shape_was_translated:
                self.after(1000, self.update_shape)
            else:
                self.update_shape()
        else:
            self.after(1000, self.save_presentation)
    
    def save_presentation(self):
        self.prs.save(f"results/{self.out_file}")
        print(f"Total Shapes {self.total_shapes}")
        print(f"{self.file_name} translated to {TARGET_LANGUAGE}")

        self.results_label.configure(f"{self.file_name} translated to {TARGET_LANGUAGE}")
    
    def translate_powerpoint(self,file_path):
        self.prs = Presentation(file_path)
        self.file_name=file_path.split("/")[-1]        
    
        self.out_file=self.file_name.replace(".pptx",f"_{TARGET_LANGUAGE}.pptx")
        if(self.out_file.count("_")==1):
            self.out_file=self.out_file.replace("_"," ")
        print(self.file_name,self.out_file)
        print("Powerpoint Opened")

        self.shapes_translated=0
        self.shape_index=0
        self.total_shapes=self.get_total_shapes()

        self.slide_index=0
        self.total_slides=len(self.prs.slides)

        self.update_shape()

        

app = App()
app.mainloop()