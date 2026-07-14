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

        self.slide_number=1
        self.total_slides=0

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

    def get_total_shapes(self,prs):
        total_shapes=0
        for slide in prs.slides:
                
            for shape in slide.shapes:
                #Try to translate shape text
                try:
                    if shape.text:
                        total_shapes+=1
                #Skip if the shape has no text
                except AttributeError:
                    pass
        return total_shapes

    def translate_powerpoint(self,file_path):
        prs = Presentation(file_path)
        file_name=file_path.split("/")[-1]        
    
        out_file=file_name.replace(".pptx",f"_{TARGET_LANGUAGE}.pptx")
        if(out_file.count("_")==1):
            out_file=out_file.replace("_"," ")
        print(file_name,out_file)
        print("Powerpoint Opened")

        self.shapes_translated=0
        self.total_shapes=self.get_total_shapes(prs)

        self.slide_number=1
        self.total_slides=len(prs.slides)

        for slide in prs.slides:
            if shapes_translated>=MAX_SHAPES_TO_TRANSLATE:
                break
                
            for shape in slide.shapes:
                #Try to translate shape text
                try:
                    translated_text=GoogleTranslator(source='auto', target='es').translate(shape.text)
                    print(translated_text)
                    shape.text=translated_text

                    status_text=f'''
                    Progress
                    Slides: {slide_number}/{total_slides}
                    Shapes: {shapes_translated}/{total_shapes}
                    '''
                    self.results_label.configure(text=status_text)

                    time.sleep(1)
                    shapes_translated+=1

                    if shapes_translated>=MAX_SHAPES_TO_TRANSLATE:
                        break
                #Skip if the shape has no text
                except AttributeError:
                    pass
            slide_number+=1

        prs.save(f"results/{out_file}")
        print(f"Total Shapes {total_shapes}")
        print(f"{file_name} translated to {TARGET_LANGUAGE}")

        self.results_label.configure(f"{file_name} translated to {TARGET_LANGUAGE}")
        #os.open(f'results/{out_file}')

app = App()
app.mainloop()