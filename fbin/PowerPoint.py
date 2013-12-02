__author__ = 'LPC'
import win32com.client as com
import pythoncom
import os

class PowerPoint:

    def __init__(self, filename):
        self.filename = filename
        self.application, self.presentation = self.create()

    def create(self):

        pythoncom.CoInitialize()
        application = com.Dispatch("PowerPoint.Application")
        return application, application.Presentations.Open(os.path.abspath(self.filename))

    def get_images_from_ppt(self, exportdir):
        self.presentation.SaveAs(os.path.abspath(exportdir), 17)  # 17 = jpg


    def quit(self):
        self.presentation.Close()
        self.application.Quit()
        pythoncom.CoUninitialize()

    def images_in_ppt(self):
        if not self.presentation:
            self.create()
        for slide in self.presentation:
            if self.image_in_slide(slide):
                return True
        return False

    @staticmethod
    def image_in_slide(slide):
        for shape in slide.Shapes:
                if shape.HasTextFrame == 0:
                    return True
        return False

    def slides_with_images(self):
        imageslides = []
        for slide in self.presentation.Slides:
            if self.image_in_slide(slide):
                imageslides.append(slide.SlideNumber)
        return imageslides

    def slides(self):
        return self.presentation.Slides
        #TODO: Generator?