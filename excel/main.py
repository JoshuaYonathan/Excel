import os
import win32com.client

class OpenExcel():
    def __init__(self, filepath, visible=True, save=True):
        self.filepath = filepath
        self.visible = visible
        self.save = save

    def __call__(self):
        self.xlapp = win32com.client.gencache.EnsureDispatch("Excel.Application")
        self.xlapp.Visible = self.visible
        self.xlwb = xlapp.Workbooks.Open(os.path.abspath(self.filepath))
        return self.xlwb

    def __enter__(self):
        self.xlapp = win32com.client.gencache.EnsureDispatch("Excel.Application")
        self.xlapp.Visible = self.visible
        self.xlwb = self.xlapp.Workbooks.Open(os.path.abspath(self.filepath))
        return self.xlwb

    def __exit__(self, exc_type, exc_value, traceback):
        self.xlwb.Close(self.save)
        self.xlapp.Quit()
        del self.xlwb
        del self.xlapp
