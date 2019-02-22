#!/usr/bin/python3
# -*- coding: UTF-8 -*-
#
# generated by wxGlade 0.8.3 on Mon Jul  2 15:38:31 2018
#

import wx

# begin wxGlade: dependencies
# end wxGlade

from excel2csv import Excel2csv


# begin wxGlade: extracode
# end wxGlade


class FileDropTarget(wx.FileDropTarget):

    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, file_names):
        e2c = Excel2csv(str(file_names[0]))
        e2c.output_csv_files()
        return True


class MyFrame(wx.Frame):

    def __init__(self, *args, **kwds):
        # begin wxGlade: MyFrame.__init__
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((400, 300))

        self.__set_properties()
        self.__do_layout()
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MyFrame.__set_properties
        self.SetTitle("Excel->CSV直行便 v1.00")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MyFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add((400, 120), 0, wx.EXPAND, 0)

        label_1 = wx.StaticText(self, wx.ID_ANY,
                                u"この文字に合わせてExcelファイルをドラッグ＆ドロップ",
                                style=wx.ALIGN_CENTER)
        label_1.SetMinSize((400, 20))
        label_1.SetFont(wx.Font(13, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))

        drop_target = FileDropTarget(label_1)
        label_1.SetDropTarget(drop_target)

        sizer_1.Add(label_1, 0, wx.ALIGN_CENTER, 0)
        sizer_1.Add((400, 130), 0, wx.EXPAND, 0)

        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade


# end of class MyFrame

class MyApp(wx.App):

    def OnInit(self):
        self.frame = MyFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


# end of class MyApp

if __name__ == "__main__":
    msemiApp = MyApp(0)
    msemiApp.MainLoop()
