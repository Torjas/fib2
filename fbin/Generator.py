import json

__author__ = 'LPC'

from PowerPoint import PowerPoint
from FileUpload import FileUpload

from datetime import date
import os
import shutil
import re


#Helpermethods to sort filelist in a human way
def tryint(s):
    try:
        return int(s)
    except:
        return s

def alphanum_key(s):
    """ Turn a string into a list of string and number chunks.
        "z23a" -> ["z", 23, "a"]
    """
    return [ tryint(c) for c in re.split('([0-9]+)', s) ]

def sort_nicely(l):
    """ Sort the given list in the way that humans expect.
    """
    l.sort(key=alphanum_key)





class Generator:
    def __init__(self, filename, template, exportdir, outfilename, remotePicturePath, pictureseverywhere=True):
        self.remotepicturepath = remotePicturePath
        self.file = filename
        self.templatefile = template
        self.exportdir = exportdir
        self.outfilename = outfilename
        self.powerpoint = PowerPoint(self.file)
        self.pictureseverywhere = pictureseverywhere

        self.templatesite = "news/template_site.html"
        self.templateitem = "news/template_item.html"

    def picture_generator(self, host, user, key):
        try:
            shutil.rmtree(self.exportdir)  # loesche altes exportDir
        except WindowsError:
            pass
        self.powerpoint.get_images_from_ppt(self.exportdir)

        imgstring = ''

        filelist = [os.path.join(self.exportdir, self.outfilename)]
        oslistdir = os.listdir(self.exportdir)
        sort_nicely(oslistdir)

        for f in oslistdir:
            imgstring += (" " * 50) + "{ url: '" + f + "'},\n"
            filelist.append(os.path.join(self.exportdir, f))

        template = open(self.templatefile, 'r').read()
        template = template.replace('IMAGE_LIST', imgstring[:-2])
        f = open(os.path.join(self.exportdir, self.outfilename), 'w+')
        f.write(template)
        f.close()
        fileupload = FileUpload(host, user, key)
        fileupload.multi_upload(filelist, self.remotepicturepath)
        fileupload.close()

    def dict_generator(self):

        tmppath = "tmp"
        if self.pictureseverywhere:
            slidelist = range(len(self.powerpoint.slides())+1)
        else:
            slidelist = self.powerpoint.slides_with_images()

        try:
            shutil.rmtree(self.exportdir)
        except WindowsError:
            pass  # TODO: Abfangen
        self.powerpoint.get_images_from_ppt(tmppath)
        filelist = []
        for f in os.listdir(tmppath):
            filelist.append(f)
        sort_nicely(filelist)

        presentationcontent = {s: {'file': 'pics/' + filelist[s-1]} for s in slidelist}

        slides = {}
        for slide in self.powerpoint.slides():
            shapedict = {}
            for shape in slide.Shapes:
                if shape.HasTextFrame and not "Datums" in shape.Name and not "Foliennummer" in shape.Name:
                    shapedict[shape.Name] = shape.TextFrame.TextRange.Text
            slides[slide.SlideNumber] = shapedict
        for s in slides:
            if s in presentationcontent:
                slides[s].update(presentationcontent[s])

        #from pprint import pprint
        #pprint(slides)
        return slides

    def to_json(self, content):
        if len(content) > 0 and type(content) == dict:
            for key, slide in content.items():
                for key, shape in slide.items():
                    shape = self.replacenewline(shape, '\n')
                    shape = shape.encode('latin-1', 'replace')
            #from pprint import pprint
            #pprint(content)
            return content
        return None


    def to_html(self, content):

        if len(content) > 0 and type(content) == dict:
            slidehtml = ""
            for slide in range(1, len(content)+1):
                image = ""
                if 'file' in content[slide]:

                    image = '''<div class="image">
                                <a href="{0}">
                                <img class="slideimage" src="{0}" alt="Slide {1}" >
                                </a>
                            </div>'''.format(content[slide].get('file', ""), slide)
                    del content[slide]['file']
                shapeshtml = ""
                for shape in content[slide]:
                    replaced = self.replacenewline(content[slide][shape], '<br/>')
                    shapeshtml += '<div class="shape">' + replaced + '</div> \n'
                shapeshtml += image

                template = ""
                with open(self.templateitem, 'r') as f:
                    template += f.read()\
                        .replace("PLACEHOLDER_SLIDECONTENT", str(shapeshtml.encode('latin-1', 'replace')))
                slidehtml += template
            template = ""
            with open(self.templatesite, "r") as f:
                template += f.read().replace("PLACEHOLDER_ITEMS", slidehtml)\
                    .replace('PLACEHOLDER_DATE', date.today().strftime('%d.%m.%Y'))

            return template
        return None

    def site_generator(self, host, user, key):
        content = self.dict_generator()

        with open('data.json', 'w') as f:
            json.dump(self.to_json(content), f)

        with open('index.html', 'w') as f:
            f.write(self.to_html(content))



        fileupload = FileUpload(host, user, key)
        fileupload.upload('index.html', "www")
        fileupload.upload('data.json', "www")
        fileupload.close()


    def close_presentation(self):
        self.powerpoint.quit()

    def replacenewline(self, string, nlchar):
        string = string[0] + string[1:-1].replace(chr(11), nlchar) + string[-1]
        string = string.replace(chr(11), '')
        return string

