import openpyxl 
import cv2
from tkinter import *
from tkinter import filedialog

from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

import numpy as np

from fpdf import FPDF
import os
class Generator():
    '''
    Creates a certificate generator object with methods to
    automatically generate certificates from a given excelsheet
    and a certificate template
    '''
    font_options = []
    def __init__(self, dir_excelsheet, sheet_name, header_location, dir_certificate, dir_output):
        '''
        Initialize the generator class object and load required constants.
        For version, there will be no GUI, only a script 
        '''
        self.dir_certificate    = dir_certificate
        self.dir_output         = dir_output
        self.names              = self._load_names(dir_excelsheet, sheet_name, header_location)
        self.certificate        = self._load_certificate()
    def __click_event(self, event, x, y, a, b):
        if event == cv2.EVENT_LBUTTONDOWN:
            self.coordinates = (x,y)
            self.coordinates_as_fraction = (x/self.original_width, y/self.original_height)
            print(self.coordinates)
    def _load_names(self, dir_excelsheet, sheet_name, header_location):
        '''
        Returns an array of all the names in the excelsheet
        with the location "dir_excelsheet"
        '''
        book = openpyxl.load_workbook(filename = dir_excelsheet)
        cells = book[sheet_name]["A"]
        # This will be later changed to ask the user what range to use for the name
        names = []
        for i in range(1, len(cells)):
            names.append(cells[i].value)
        return names
    def _load_certificate(self):
        '''
        Loads the image in of the certificate and coordinates are marked here
        '''
        img = cv2.imread(self.dir_certificate)
        height, width, channels = img.shape
        self.original_height    = height
        self.original_width       = width

        cv2.namedWindow('Point Coordinates')
        cv2.setMouseCallback('Point Coordinates', self.__click_event)
        while True:
            cv2.imshow('Point Coordinates',img)
            k = cv2.waitKey(1) & 0xFF
            if k == 27:
                break
        cv2.destroyAllWindows()
        return img
    def generate_preview(self, name):
        # Factor in the length of the name and the font size
        # to change how text is overlayed
        self.image = self._prepare_certificate(name, isPreview=True)
        cv2.namedWindow('Preview')
        while True:
            cv2.imshow('Preview', self.image)
            k = cv2.waitKey(1) & 0xFF
            if k == 27:
                break
            if cv2.getWindowProperty('Preview',cv2.WND_PROP_VISIBLE) < 1:        
                break
        cv2.destroyAllWindows()
        # Add logic to stop generation if certificate is not proper
    def _prepare_certificate(self, name, isPreview=False):
        img = cv2.imread(self.dir_certificate)

        # check scaling here
        if isPreview:
            height, width, channels = img.shape
            image = Image.fromarray(img)
            draw = ImageDraw.Draw(image)
            img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2RGBA)
            height, width, channels = img.shape
            # return a resized image
            font_size = 70
            font = ImageFont.truetype("temp/static/Comfortaa-Bold.ttf", font_size)
            
            # check scaling here
            draw.text(((self.coordinates_as_fraction[0]* width -len(name)/2*font_size/1.9), (self.coordinates_as_fraction[1]* height -font_size/2)), name, font = font, fill=(0,0,0))
            image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2RGBA)
            return image
        else:
            image = Image.fromarray(img)
            draw = ImageDraw.Draw(image)
            font_size = 70
            font = ImageFont.truetype("temp/static/Comfortaa-Bold.ttf", font_size)
            # check scaling here
            draw.text(((self.coordinates[0]-len(name)/2*font_size/1.9), (self.coordinates[1]-font_size/2)), name, font = font, fill=(0,0,0))
            image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            return image
    def generate_certificates(self):
        for i in self.names:
            img = self._prepare_certificate(i)
            image = Image.fromarray(img)
            image.save(self.dir_output+"/"+i+".jpg", quality=100, subsampling=0)

    def generate_pdf(self):
        '''
        Convert all the images from the output dir into a pdf to make
        printing easier
        '''
        dir_list = os.listdir("./"+self.dir_output)
        pdf = FPDF()
        for i in dir_list:
            print(i)
            pdf.add_page()
            pdf.image(os.path.join(self.dir_output,i))
            # os.remove(i)
        pdf.output(os.path.join("yourfile.pdf"), "F")
        

class UI():
    '''
    The UI class that governs how the UI is structured.
    '''
    def __init__(self):
        self.window = Tk()
        self.window.title("Automatic Certificate Generator")
        self.window.geometry("500x500")
        
        self.button_excel = Button(self.window, 
                        text = "Select excelsheet",
                        command = self._get_excelsheet)
        self.button_excel.grid(column = 2, row = 1)
        
        self.button_certificate = Button(self.window, text = "Select certificate", command = self._get_certificate)
        self.button_certificate.grid(column=2, row=2)
        
        self.label_excelsheet = Label(self.window, text="Select excelsheet", width=50, height = 4, fg="black")
        self.label_certificate = Label(self.window, text="Select certificate", width=50, height = 4, fg="black")

        self.label_output = Label(self.window, text="Select output folder", width=50, height = 4, fg="black")
        self.button_output = Button(self.window, text="Select output folder", command=self._get_outputdir)
        self.button_output.grid(column=2, row=3)


        self.label_excelsheet.grid(column = 1, row = 1)
        self.label_certificate.grid(column = 1, row = 2)
        self.label_output.grid(column=1, row = 3)
        
        self.window.mainloop()
    def _get_excelsheet(self):
        self.dir_excelsheet = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select Excelsheet",
                                          filetypes = (("Excel sheet",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*")))
        self.label_excelsheet.configure(text="File selected")
    def _get_certificate(self):
        self.dir_certificate = filedialog.askopenfilename(initialdir="/", title="Select Certificate sample", filetypes=(("PNG", "*.png"), ("JPG", "*.jpeg"),("all files", "*.*")))
        self.label_certificate.configure(text="File selected")
        if(self.dir_excelsheet and self.dir_certificate):
            self.button_pick = Button(self.window, text = "Choose location", command=self._generate)
            self.button_pick.grid(column=1, row=4)
    def _get_outputdir(self):
        self.dir_output = filedialog.askdirectory()
        self.label_output.configure(text="Folder selected: "+self.dir_output)
    def _generate(self):
        main = Generator(self.dir_excelsheet, "Sheet1", "Name", self.dir_certificate, self.dir_output)
        main.generate_preview("Tanay Rikeshbhai Shah")
        self.button_fix = Button(self.window, text="Generate certificates", command=main.generate_certificates)
        self.button_fix.grid(column=1, row=5)
        self.button_pdf = Button(self.window, text="Generate pdf", command=main.generate_pdf)
        self.button_pdf.grid(column=1, row=6)
         
def main():
    ui = UI()
    

if __name__=='__main__':
    main()