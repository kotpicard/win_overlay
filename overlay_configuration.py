import wx
from custom_controls import SaveFilePicker, FontPickerCtrl, NamedColourPicker, GlobalClickPicker, HotkeyCtrl
import ctypes
from hotkeys import *
import json
import threading
import win32api
import overlay
import win32con

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except Exception:
    pass

TYPE_TEXT = "text"
TYPE_WILDCARD = "wildcard"
TYPE_MESSAGE = "message"
TYPE_VISIBLE = "visible"
DEFAULT_HOTKEYS = {
    HOTKEY_SHOW: ["Ctrl", "Shift", "W"],
    HOTKEY_HIDE: ["Ctrl", "Shift", "Q"],
    HOTKEY_NEXTTEXT: ["Ctrl", "Shift", "T"],
    HOTKEY_PREVTEXT: ["Ctrl", "Shift", "R"],
    HOTKEY_SHOWPPT: ["Ctrl", "Shift", "G"],
    HOTKEY_HIDEPPT: ["Ctrl", "Shift", "A"],
    HOTKEY_NEXTSLIDE: ["Ctrl", "Shift", "F"],
    HOTKEY_PREVSLIDE: ["Ctrl", "Shift", "D"],
    HOTKEY_STARTTIMER: ["Ctrl", "Shift", "S"],
    HOTKEY_QUIT: ["Ctrl", "Shift", "X"]
}


def SetMessage(widget, text):
    widget._dlg_message = text


def SetWildcard(widget, wildcard):
    widget._dlg_wildcard = wildcard


def ToggleVisibility(sizer, visible):
    if visible:
        UnhideSizer(sizer)
    else:
        HideSizer(sizer)


def HideSizer(sizer):
    for i in range(sizer.GetItemCount()):
        item = sizer.GetItem(i)

        # If the item is a window, hide it
        window = item.GetWindow()
        if window:
            window.Hide()

        # If the item is a sizer, recurse
        child_sizer = item.GetSizer()
        if child_sizer:
            HideSizer(child_sizer)


def UnhideSizer(sizer):
    for i in range(sizer.GetItemCount()):
        item = sizer.GetItem(i)

        # If the item is a window, show it
        window = item.GetWindow()
        if window:
            window.Show()

        # If the item is a sizer, recurse
        child_sizer = item.GetSizer()
        if child_sizer:
            UnhideSizer(child_sizer)


def clear_sizer(sizer, delete_windows=True):
    for i in reversed(range(sizer.GetItemCount())):
        item = sizer.GetItem(i)
        window = item.GetWindow()
        child_sizer = item.GetSizer()

        if window:
            sizer.Detach(window)
            if delete_windows:
                window.Destroy()
        elif child_sizer:
            sizer.Detach(child_sizer)
            if delete_windows:
                clear_sizer(child_sizer, delete_windows)
        else:
            sizer.Remove(i)


class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Overlay Configuration", size=(500, 800),
                         style=wx.STAY_ON_TOP | wx.DEFAULT_FRAME_STYLE)

        self.UsePowerPoint = True
        self.SimultaneousChange = True
        self.ToggleWithPPT = True
        self.CreateLog = True
        self.hotkey_num = 10
        self.capturinghotkey = False
        self.overlay_window = None
        self.OverlayRunning = False

        self.PPTVALUES = {
            "textsource_setting": "Presentation to use as text source:",
            "textsource_wildcard": "PowerPoint files (*.ppt;*.pptx)|*.ppt;*.pptx",
            "textsource_msg": "Choose a PowerPoint file",
            "pptsettings_visible": True
        }

        self.NONPPTVALUES = {
            "textsource_setting": "File to use as text source:",
            "textsource_wildcard": "Text files (*.txt)|*.txt",
            "textsource_msg": "Choose a text file",
            "pptsettings_visible": False
        }

        self.CURRENTVALUES = self.PPTVALUES

        self.mainPanel = wx.Panel(self, name="mainPanel")
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.notebook = wx.Notebook(self.mainPanel)
        self.mainSizer.Add(self.notebook, 1, wx.ALL, 0)
        buttonsizer = self.SetUpButtons()
        self.mainSizer.Add(buttonsizer, 0, wx.ALL, 0)
        self.mainPanel.SetSizer(self.mainSizer)
        self.Refreshable = []
        self.HOTKEYLABELS = {
            HOTKEY_NEXTTEXT: "Next Overlay Text:",
            HOTKEY_PREVTEXT: "Previous Overlay Text:",
            HOTKEY_SHOW: "Overlay On:",
            HOTKEY_HIDE: "Overlay Off:",
            HOTKEY_STARTTIMER: "Start Log Timer:",
            HOTKEY_QUIT: "Quit:",
            HOTKEY_NEXTSLIDE: "Next PPT Slide:",
            HOTKEY_PREVSLIDE: "Previous PPT Slide:",
            HOTKEY_SHOWPPT: "Go To PPT:",
            HOTKEY_HIDEPPT: "Minimize PPT:"
        }
        self.HOTKEYORDER = [
            HOTKEY_NEXTTEXT,
            HOTKEY_PREVTEXT,
            HOTKEY_SHOW,
            HOTKEY_HIDE,
            HOTKEY_STARTTIMER,
            HOTKEY_QUIT,
            HOTKEY_NEXTSLIDE,
            HOTKEY_PREVSLIDE,
            HOTKEY_SHOWPPT,
            HOTKEY_HIDEPPT
        ]
        self.hotkey_controls = []
        self.hotkey_values = {
            HOTKEY_SHOW: ["Ctrl", "Shift", "W"],
            HOTKEY_HIDE: ["Ctrl", "Shift", "Q"],
            HOTKEY_NEXTTEXT: ["Ctrl", "Shift", "T"],
            HOTKEY_PREVTEXT: ["Ctrl", "Shift", "R"],
            HOTKEY_SHOWPPT: ["Ctrl", "Shift", "G"],
            HOTKEY_HIDEPPT: ["Ctrl", "Shift", "A"],
            HOTKEY_NEXTSLIDE: ["Ctrl", "Shift", "F"],
            HOTKEY_PREVSLIDE: ["Ctrl", "Shift", "D"],
            HOTKEY_STARTTIMER: ["Ctrl", "Shift", "S"],
            HOTKEY_QUIT: ["Ctrl", "Shift", "X"]
        }

        self.SetUpAll()

    def SetUpButtons(self):
        buttonsizer = wx.BoxSizer(wx.HORIZONTAL)
        buttonsizer.AddStretchSpacer(2)
        btnPreview = wx.Button(self.mainPanel, label="Preview Overlay", name="previewbutton")
        btnPreview.Bind(wx.EVT_BUTTON, self.PreviewOverlay)
        btnSave = wx.Button(self.mainPanel, label="Save All Settings")
        btnSave.Bind(wx.EVT_BUTTON, self.SaveSettings)
        buttonsizer.Add(btnPreview, 1, wx.ALL, 10)
        buttonsizer.Add(btnSave, 1, wx.ALL, 10)
        return buttonsizer

    def SetUpAll(self):
        self.SetUpGeneralSettings()
        self.SetUpTextSettings()
        self.SetUpHotkeySettings()

    def SetUpGeneralSettings(self):
        panel = wx.Panel(self.notebook, name="GeneralSettingsPanel")
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(wx.StaticText(panel, label="Use PowerPoint presentation as text source?"), 0, wx.ALL, 5)

        sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.pptselect = wx.RadioButton(panel, label="Yes", style=wx.RB_GROUP)
        self.pptselect.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        self.pptselect.changed = "UsePowerPoint"
        self.pptselect.changeto = True
        self.alttextselect = wx.RadioButton(panel, label="No")
        self.alttextselect.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        self.alttextselect.changed = "UsePowerPoint"
        self.alttextselect.changeto = False
        sizer.Add(self.pptselect, 0, wx.ALL | wx.CENTER, 10)
        sizer.Add(self.alttextselect, 0, wx.ALL | wx.CENTER, 10)
        mainSizer.Add(sizer)

        label_textsource = wx.StaticText(panel, label=self.CURRENTVALUES["textsource_setting"])
        self.AddToRefreshable(label_textsource, label_textsource.SetLabelText, TYPE_TEXT, "textsource_setting")
        mainSizer.Add(label_textsource, 0, wx.ALL, 5)

        picker = wx.FilePickerCtrl(panel, message=self.CURRENTVALUES["textsource_msg"],
                                   wildcard=self.CURRENTVALUES["textsource_wildcard"], name="textsource")
        picker.SetMinSize((450, -1))
        picker._dlg_message = self.CURRENTVALUES["textsource_msg"]
        picker._dlg_wildcard = self.CURRENTVALUES["textsource_wildcard"]

        def on_picker_button(evt):
            msg = getattr(picker, "_dlg_message", "Choose a file")
            wild = getattr(picker, "_dlg_wildcard", "All files (*.*)|*.*")

            dlg = wx.FileDialog(
                parent=panel,
                message=msg,
                wildcard=wild,
                style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
            )
            if dlg.ShowModal() == wx.ID_OK:
                picker.SetPath(dlg.GetPath())
            dlg.Destroy()

        btn = picker.GetPickerCtrl()
        btn.SetLabel("Browse")
        btn.Bind(wx.EVT_BUTTON, on_picker_button)

        self.AddToRefreshable(picker, SetMessage, TYPE_MESSAGE, "textsource_msg")
        self.AddToRefreshable(picker, SetWildcard, TYPE_WILDCARD, "textsource_wildcard")
        mainSizer.Add(picker, 0, wx.ALL, 5)

        pptSettingsSizer = wx.BoxSizer(wx.VERTICAL)
        pptSettingsSizer.Add(wx.StaticText(panel, label="PowerPoint Control Settings"), 0, wx.ALL, 5)
        pptSettingsSizer.Add(wx.StaticText(panel, label="Change PowerPoint slides and Overlay Text together?"), 0,
                             wx.LEFT, 5)
        rbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        simultaneous_change_yes = wx.RadioButton(panel, label="Yes", style=wx.RB_GROUP)
        simultaneous_change_yes.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        simultaneous_change_yes.changed = "SimultaneousChange"
        simultaneous_change_yes.changeto = True
        simultaneous_change_no = wx.RadioButton(panel, label="No")
        simultaneous_change_no.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        simultaneous_change_no.changed = "SimultaneousChange"
        simultaneous_change_no.changeto = False
        rbsizer1.Add(simultaneous_change_yes, 0, wx.ALL | wx.CENTER, 10)
        rbsizer1.Add(simultaneous_change_no, 0, wx.ALL | wx.CENTER, 10)
        pptSettingsSizer.Add(rbsizer1)

        pptSettingsSizer.Add(wx.StaticText(panel, label="Hide Overlay when PowerPoint open?"), 0,
                             wx.LEFT, 5)
        rbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        toggle_together_yes = wx.RadioButton(panel, label="Yes", style=wx.RB_GROUP)
        toggle_together_yes.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        toggle_together_yes.changed = "ToggleWithPPT"
        toggle_together_yes.changeto = True
        toggle_together_no = wx.RadioButton(panel, label="No")
        toggle_together_no.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        toggle_together_no.changed = "ToggleWithPPT"
        toggle_together_no.changeto = False
        rbsizer2.Add(toggle_together_yes, 0, wx.ALL | wx.CENTER, 10)
        rbsizer2.Add(toggle_together_no, 0, wx.ALL | wx.CENTER, 10)
        pptSettingsSizer.Add(rbsizer2)
        self.AddToRefreshable(pptSettingsSizer, ToggleVisibility, TYPE_VISIBLE, "pptsettings_visible")

        mainSizer.Add(pptSettingsSizer)

        logSettingsSizer = wx.BoxSizer(wx.VERTICAL)
        logSettingsSizer.Add(wx.StaticText(panel, label="Log Control Settings"), 0, wx.ALL, 5)
        logSettingsSizer.Add(wx.StaticText(panel, label="Create Log?"), 0,
                             wx.LEFT, 5)
        rbsizer3 = wx.BoxSizer(wx.HORIZONTAL)
        create_log_yes = wx.RadioButton(panel, label="Yes", style=wx.RB_GROUP)
        create_log_yes.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        create_log_yes.changed = "CreateLog"
        create_log_yes.changeto = True
        create_log_no = wx.RadioButton(panel, label="No")
        create_log_no.Bind(wx.EVT_RADIOBUTTON, self.SwitchAttribute)
        create_log_no.changed = "CreateLog"
        create_log_no.changeto = False
        rbsizer3.Add(create_log_yes, 0, wx.ALL | wx.CENTER, 10)
        rbsizer3.Add(create_log_no, 0, wx.ALL | wx.CENTER, 10)
        logSettingsSizer.Add(rbsizer3)
        loglabel = wx.StaticText(panel, label="Log File Path")
        logSettingsSizer.Add(loglabel, 0, wx.ALL, 10)
        log_picker = SaveFilePicker(panel, message="Pick location for log file",
                                    wildcard="Text files (*.txt)|*.txt|All files (*.*)|*.*", name="logpicker")
        log_picker.SetMinSize((450, -1))
        logSettingsSizer.Add(log_picker)
        create_log_yes.logpicker = log_picker
        create_log_no.logpicker = log_picker
        create_log_yes.loglabel = loglabel
        create_log_no.loglabel = loglabel
        mainSizer.Add(logSettingsSizer)

        panel.SetSizer(mainSizer)
        self.notebook.AddPage(panel, "General Settings")

    def SetUpTextSettings(self):
        panel = wx.Panel(self.notebook, name="TextSettingsPanel")
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(wx.StaticText(panel, label="Select Overlay Font"), 0, wx.ALL, 10)
        fontpicker = FontPickerCtrl(panel, name="fontpicker")
        mainSizer.Add(fontpicker, 0, wx.LEFT | wx.BOTTOM, 5)
        mainSizer.Add(wx.StaticText(panel, label="Select Font Size"), 0, wx.LEFT | wx.TOP, 10)
        spinctrl = wx.SpinCtrl(panel, min=0, max=100, initial=20, name="fontsize")

        def PreviewFontSettings(evt):
            fontpicker.current_font.SetPointSize(spinctrl.GetValue())
            fontpicker.font_display.SetForegroundColour(colorpicker.GetColour())
            fontpicker.UpdateFontDisplay()
            panel.Layout()
            fontpicker.GetSelectedFont()

        spinctrl.Bind(wx.EVT_SPINCTRL, PreviewFontSettings)

        mainSizer.Add(spinctrl, 0, wx.ALL, 10)
        mainSizer.Add(wx.StaticText(panel, label="Select Font Color"), 0, wx.LEFT | wx.TOP, 10)
        colorpicker = NamedColourPicker(panel, style=wx.CLRP_USE_TEXTCTRL, name="colorpicker")
        colorpicker.Bind(wx.EVT_COLOURPICKER_CHANGED, PreviewFontSettings)
        mainSizer.Add(colorpicker, 0, wx.ALL, 10)

        pospicker = GlobalClickPicker(panel, name="pospicker")
        mainSizer.Add(pospicker, 0, wx.LEFT, 5)

        panel.SetSizer(mainSizer)
        self.notebook.AddPage(panel, "Text Settings")

    def SetUpHotkeySettings(self):
        hotkey_panel = wx.Panel(self.notebook)

        sizer = self.SetUpHotkeySizer(hotkey_panel)

        hotkey_panel.SetSizer(sizer)
        self.notebook.AddPage(hotkey_panel, "Hotkeys")

    def SetUpHotkeySizer(self, parent):
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.hotkey_controls = []
        for i in range(self.hotkey_num):
            row_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row_sizer.hotkey_id = self.HOTKEYORDER[i]
            label = wx.StaticText(parent, label=self.HOTKEYLABELS[row_sizer.hotkey_id])
            hotkey_input = HotkeyCtrl(parent)
            hotkey_input.SetValue(" + ".join(self.hotkey_values[row_sizer.hotkey_id]))
            row_sizer.Add(label, 0, wx.ALL | wx.CENTER, 5)
            row_sizer.AddStretchSpacer(1)
            row_sizer.Add(hotkey_input, 0, wx.ALL, 5)
            sizer.Add(row_sizer, 0, wx.EXPAND)
            self.hotkey_controls.append(row_sizer)
        return sizer

    def UpdateHotkeys(self):
        hotkey_panel = self.notebook.GetPage(2)
        sizer = hotkey_panel.GetSizer()
        clear_sizer(sizer, delete_windows=True)
        sizer = self.SetUpHotkeySizer(hotkey_panel)
        hotkey_panel.SetSizer(sizer)
        hotkey_panel.Layout()

    def SwitchAttribute(self, evt):
        widget = evt.GetEventObject()
        attr_name = widget.changed
        attr_value = widget.changeto
        self.__setattr__(attr_name, attr_value)
        print(f"Changed {attr_name} value to {attr_value}")
        if attr_name == "UsePowerPoint":
            self.ApplyPPTSettings()
        elif attr_name == "CreateLog":
            if attr_value:
                widget.logpicker.Show()
                widget.loglabel.Show()
            else:
                widget.logpicker.Hide()
                widget.loglabel.Hide()

    def AddToRefreshable(self, widget, func, change_type, arg):
        self.Refreshable.append([widget, func, change_type, arg])

    def ApplyPPTSettings(self):
        self.CURRENTVALUES = self.PPTVALUES if self.UsePowerPoint else self.NONPPTVALUES
        self.hotkey_num = 10 if self.UsePowerPoint else 6
        self.UpdateHotkeys()
        for elem in self.Refreshable:
            widget = elem[0]
            func = elem[1]
            if elem[2] == TYPE_TEXT:
                arg = self.CURRENTVALUES[elem[3]]
                func(arg)
            if elem[2] == TYPE_MESSAGE or elem[2] == TYPE_WILDCARD:
                arg = self.CURRENTVALUES[elem[3]]
                func(widget, arg)
            if elem[2] == TYPE_VISIBLE:
                arg = self.CURRENTVALUES[elem[3]]
                func(widget, arg)
                self.notebook.GetCurrentPage().Layout()

    def PreviewOverlay(self, evt):
        if self.OverlayRunning:
            self.StopOverlay()
            self.OverlayRunning = False
            wx.FindWindowByName("previewbutton", self.mainPanel).SetLabel("Preview Overlay")
        else:
            if self.CheckOverlayPositionSet():
                self.OverlayRunning = True
                self.RunOverlay()
                wx.FindWindowByName("previewbutton", self.mainPanel).SetLabel("Stop Preview")
            else:
                wx.MessageBox(
                "Overlay position not set!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )



    def SaveSettings(self, evt):
        if self.ValidateSettings():
            print("Settings validated. Saving...")
            settings = self.GetAllSettings()
            fh = open("settings.json", "w")
            json.dump(settings, fh, indent=2)
            fh.close()

    def GetHotkeyValues(self):
        return [self.hotkey_controls[i].GetChildren()[2].GetWindow().GetValue() for i in
                range(self.hotkey_num)]

    def GetFontColorValue(self):
        r, g, b, a, = wx.FindWindowByName("colorpicker", self.notebook.GetPage(1)).GetColour()
        return [r, g, b]

    def SetHotkeyValueAttr(self):
        values = [[self.hotkey_controls[i].GetChildren()[2].GetWindow().GetValue(), self.hotkey_controls[i].hotkey_id]
                  for i in
                  range(self.hotkey_num)]
        for elem in values:
            hotkey_id = elem[1]
            value = elem[0].split(" + ")
            self.hotkey_values[hotkey_id] = value

    def GetAllSettings(self):
        overlay_text_settings = {
            "font_path": wx.FindWindowByName("fontpicker", self.notebook.GetPage(1)).GetSelectedFont(),
            "font_size": wx.FindWindowByName("fontsize", self.notebook.GetPage(1)).GetValue(),
            "text_color": self.GetFontColorValue(),
            "x_pos": wx.FindWindowByName("pospicker", self.notebook.GetPage(1)).GetValue()[0],
            "y_pos": wx.FindWindowByName("pospicker", self.notebook.GetPage(1)).GetValue()[1],
        }

        ppt_settings = {
            "ppt_path": wx.FindWindowByName("textsource",
                                            self.notebook.GetPage(0)).GetPath() if self.UsePowerPoint else ""
        }
        hotkey_settings = self.hotkey_values

        general_settings = {
            "use_ppt": self.UsePowerPoint,
            "simultaneous_change": self.SimultaneousChange,
            "toggle_overlay_with_ppt": self.ToggleWithPPT,
            "texts_path": wx.FindWindowByName("textsource",
                                              self.notebook.GetPage(0)).GetPath() if not self.UsePowerPoint else "",
            "hotkeys": hotkey_settings,
            "create_log": self.CreateLog,
            "log_path": wx.FindWindowByName("logpicker", self.notebook.GetPage(0)).GetPath() if self.CreateLog else ""
        }

        settings = {
            "general": general_settings,
            "text": overlay_text_settings,
            "ppt": ppt_settings
        }
        return settings

    def ValidateSettings(self):
        if not self.CheckTextSourceSet():
            wx.MessageBox(
                "Text source path not set!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return False
        elif not self.CheckLogPathSet():
            wx.MessageBox(
                "Log path not set!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return False

        elif not self.CheckAllHotkeysSet():
            wx.MessageBox(
                "Missing hotkeys!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return False

        elif not self.CheckAllHotkeysUnique():
            wx.MessageBox(
                "Some hotkeys are not unique!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return False
        elif not self.CheckOverlayPositionSet():
            wx.MessageBox(
                "Overlay position not set!",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return False

        else:
            return True

    def CheckTextSourceSet(self):
        return True if wx.FindWindowByName("textsource", self.notebook.GetPage(0)).GetPath() else False

    def CheckLogPathSet(self):
        if self.CreateLog and wx.FindWindowByName("logpicker", self.notebook.GetPage(0)).GetPath():
            return True
        elif not self.CreateLog:
            return True
        else:
            return False

    def CheckAllHotkeysSet(self):
        return all(
            [self.hotkey_controls[i].GetChildren()[2].GetWindow().GetValue() != "" for i in range(self.hotkey_num)])

    def CheckAllHotkeysUnique(self):
        hotkeys = self.GetHotkeyValues()
        return len(list(set(hotkeys))) == len(hotkeys)

    def CheckOverlayPositionSet(self):
        if wx.FindWindowByName("pospicker", self.notebook.GetPage(1)).GetValue():
            return True

    def RunOverlay(self):
        if self.CheckOverlayPositionSet():
            settings = self.GetAllSettings()
            text_config = settings["text"]

            def overlay_thread_target():
                import pythoncom
                pythoncom.CoInitialize()

                self.overlay_window = overlay.OverlayWindow(
                    extended_style=win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
                    class_name="OverlayWindow",
                    window_name="OverlayWindow",
                    style=win32con.WS_POPUP,
                    x_pos=0,
                    y_pos=win32api.GetSystemMetrics(1) - 200,
                    width=win32api.GetSystemMetrics(0),
                    height=win32api.GetSystemMetrics(1),
                    parent_window=None,
                    h_menu=None,
                    lp_void=None,
                    text_settings=text_config,
                    texts=["Overlay Preview"],
                    hotkey_list=[],
                    simultaneous_change=False,
                    simulated_hotkeys=[],
                    create_log=False,
                    log_path=""
                )

                self.overlay_window.Run()

            self.overlay_thread = threading.Thread(
                target=overlay_thread_target,
                daemon=True
            )
            self.overlay_thread.start()

    def StopOverlay(self):
        self.overlay_window.Stop()



class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame()
        frame.Show()
        return True


def main():
    app = MyApp(False)
    app.MainLoop()


if __name__ == "__main__":
    main()
