import wx
import wx.xrc


###########################################################################
## Class MyFrame3
###########################################################################

class MyFrame3(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title='我的GUI程序', pos=wx.DefaultPosition,
                          size=wx.Size(500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_button2 = wx.Button(self, wx.ID_ANY, u"打开文件", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_button2.SetFont(wx.Font(18, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "微软雅黑"))
        bSizer2.Add(self.m_button2, 1, wx.ALL | wx.EXPAND, 5)

        self.SetSizer(bSizer2)
        self.Layout()
        self.m_statusBar2 = self.CreateStatusBar(1, wx.STB_SIZEGRIP, wx.ID_ANY)

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_button2.Bind(wx.EVT_BUTTON, self.m_button2OnButtonClick)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def m_button2OnButtonClick(self, event):
        openFileDialog = wx.FileDialog(frame, "请选择要打开的Excel文件", "", "",
                                       "Excel格式 (*.xls)|*.xls",
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)

        if openFileDialog.ShowModal() == wx.ID_OK:
            filePath = openFileDialog.GetPath()
            if wx.MessageBox("数据处理完成", "提示", wx.OK | wx.ICON_INFORMATION) == wx.OK:
                self.m_statusBar2.SetStatusText(filePath)

        openFileDialog.Destroy()


app = wx.App(False)
frame = MyFrame3(None)
frame.Show()
app.MainLoop()