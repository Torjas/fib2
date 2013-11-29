__author__ = 'LPC'
import os
import pysftp


class FileUpload:

    def __init__(self, host, user, key):

        self.host = host
        self.user = user
        self.key = key
        self.srv = self.connect()

    def connect(self):
        try:
            return pysftp.Connection(self.host, self.user, self.key)
        except Exception as e:
            print("No Connection")



    def upload(self, filename, remotepath=""):
        p = os.path.normpath(os.path.join(remotepath, filename))
        self.srv.put(filename, p.replace(chr(92), "/"))


    def multi_upload(self, filelist, remotepath=""):
        for fn in filelist:
            self.upload(fn, remotepath)


    def close(self):
        self.srv.close()