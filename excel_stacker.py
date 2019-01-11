__author__ = 'shodges'

import wx, os
import pandas as pd

path = ''

class ExcelStacker(wx.Frame):

    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, None, title="Excel Stacker", size=(370,250))
        self.MainUI()

    def MainUI(self):

        panel = wx.Panel(self)

        x = 30
        x2 = 210

        self.folderSelectLabel = wx.StaticText(panel, pos=(x, 20), label="Choose Folder", size=(150, -1))
        self.folderSelect = wx.Button(panel, pos=(x2, 20), label="...", size=(30, -1))
        self.folderSelect.Bind(wx.EVT_BUTTON, self.onDir)

        self.folderSelectResult = wx.StaticText(panel, pos=(x, 50), label="", size=(150, -1))
        font = wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.folderSelectResult.SetFont(font)

        self.skipRowLabel = wx.StaticText(panel, pos=(x, 80), label="Skip Rows", size=(150, -1))
        self.skipRows = wx.SpinCtrl(panel, pos=(x2, 80), value="0", size=(60, -1), style=wx.SP_ARROW_KEYS, min=0, initial=0,
                                    name="wxSpinCtrl")
        self.outputFileLabel = wx.StaticText(panel, pos=(x, 110), label="Output File Name", size=(150, -1))
        self.outputFile = wx.TextCtrl(panel, pos=(x2, 110), value="result")

        self.stackFilesButton = wx.Button(panel, pos=(110, 140), label="Stack Files", size=(150,-1))
        self.stackFilesButton.Bind(wx.EVT_BUTTON, self.stackFiles)

        self.errorMessage = wx.StaticText(panel, pos=(0, 170), label="", size=(370, -1), style=wx.ALIGN_CENTER)

        self.Centre()
        self.Show()

    def onDir(self, event):
        global path
        dlg = wx.DirDialog(self, "Choose a directory:", defaultPath=os.getcwd(),
                           style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.folderSelectResult.SetLabel("Selected Folder: " + dlg.GetPath().split("\\")[-1])
        dlg.Destroy()

    def stackFiles(self, e):
        global path
        self.stackFilesButton.Disable()
        loopBreak = 0
        outputFile = self.outputFile.GetValue()
        skipRows = self.skipRows.GetValue()
        self.errorMessage.SetLabel("Loading...")

        if path == "" or outputFile == "":
            loopBreak = 1
            self.errorMessage.SetLabel("Error: Please make sure you have entered in all details")

        while loopBreak == 0:

            # Get file names from folder
            try:
                excel_files = []
                for file in os.listdir(path):
                    if file.endswith(".xls") or file.endswith(".xlsx"):
                        excel_files.append(file)

            except Exception, e:
                self.errorMessage.SetLabel("Error: Folder doesn't exist. Try again.")
                loopBreak = 1
                break

            # Read files in
            if len(excel_files) == 0:
                loopBreak = 1
                self.errorMessage.SetLabel("Error: No Excel files in folder")
                break

            try:
                excels = [pd.ExcelFile(path + '/' + file) for file in excel_files]
            except Exception, e:
                self.errorMessage.SetLabel("Error: Please make sure none of the files are open")
                loopBreak = 1
                break

            # Read first sheets and add to dataframe
            frames = [sheet.parse(sheet.sheet_names[0], header=None, skiprows=skipRows, index_col=None) for sheet in
                      excels]

            # Join dataframes together
            combined = pd.concat(frames)

            # Export Excel file
            try:
                combined.to_excel(outputFile + ".xlsx", header=False, index=False)
            except Exception, e:
                self.errorMessage.SetLabel("Error: Output file with that name is open")
                loopBreak = 1
                break


            # Break Loop
            loopBreak = 1

            self.errorMessage.SetLabel("complete")
            os.startfile(outputFile + ".xlsx")
            # Clean Vars

            path = ""
            self.folderSelectResult.SetLabel("")

        self.stackFilesButton.Enable()



def main():
    app = wx.App()
    ExcelStacker(None)
    app.MainLoop()

main()