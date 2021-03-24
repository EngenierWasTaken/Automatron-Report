
import wx
from AutomazioneReportPerLuigiMainframe import AutomazioneReportPerLuigiMainframe


class AutomazioneApp(wx.App):
    def __init__(self):
        wx.App.__init__(self)

    def OnInit(self):

        self.MainFrame = AutomazioneReportPerLuigiMainframe(None)

        self.SetTopWindow(self.MainFrame)
        self.MainFrame.Show()

        return True


if __name__ == "__main__":
    automazione = AutomazioneApp()
    automazione.MainLoop()
