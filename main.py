import openpyxl
import cv2

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
        For version, there will be no GUI, only a script th
        '''
        self.names          = self._load_names(dir_excelsheet, sheet_name, header_location)
        self.certificate    = self._load_certificate(dir_certificate)
        self.coordinates    = self._set_coordinates()
        self.font_option    = self._set_font()
        print(self.names)
        cv2.imshow(self.certificate)
    def _load_names(self, dir_excelsheet, sheet_name, header_location):
        '''
        Returns an array of all the names in the excelsheet
        with the location "dir_excelsheet"
        '''
        book = openpyxl.load_workbook(filename = dir_excelsheet)
        cells = book[sheet_name][header_location]
        # This will be later changed to ask the user what range to use for the name
        names = []
        for i in cells:
            names.append(i.value)
        return names

    def _load_certificate(self, dir_certificate):
        '''
        Returns the image in which certificate is to be loaded
        '''
        img = cv2.imread(dir_certificate, 1)
        return img
    def _set_coordinates(self):
        '''
        Display the image and ask for user input. Returns a tuple
        that contains the center of both of the 
        '''
        return ()
    def _set_format(self):
        '''
        Get the formate from the user to store images in
        '''
        return ()
    def _set_font(self):
        '''
        Set the font to use to write the names on the certificate
        '''
        pass
    def generate_image(self, file_name):
        '''
        Generate an image with a unique filename which also includes
        a custom name to make searching easier
        '''
        pass
    def generate_pdf(self):
        '''
        Convert all the images from the output dir into a pdf to make
        printing easier
        '''
        pass
def main():
    x = Generator("temp/Book1.xlsx", "Sheet1", "Name", "", "")

if __name__=='__main__':
    main()