import wx
from matplotlib import font_manager
from PIL import ImageFont
from pynput import mouse, keyboard
from hotkey_logger import HotkeyLogger
import threading

CUSTOM_COLORS = {
    "cosmos": (20, 38, 68),
    "sky": (139, 195, 224),
    "lime": (205, 232, 92),
    "mint": (120,250,170),
    "flame":(245, 80, 105),
    "polar":(20,0,255),
    "glow":(255, 135, 85),
    "graphite":(160,170,180),
    "dusk":(125,95,235),
    "trunk":(120,110,90),
    "moss":(0,175,135),
    "dawn":(235,130,200)
}


def get_color(value):
    """Convert a name or tuple to wx.Colour."""
    if isinstance(value, str) and value in CUSTOM_COLORS:
        return wx.Colour(*CUSTOM_COLORS[value])
    return wx.Colour(value)  # falls back to wx's own parser


class NamedColourPicker(wx.ColourPickerCtrl):
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("style", wx.CLRP_USE_TEXTCTRL)
        super().__init__(parent, **kwargs)

        self.text_ctrl, self.colorbutton = self.GetChildren()
        if isinstance(self.text_ctrl, wx.TextCtrl):
            self.text_ctrl.Bind(wx.EVT_TEXT, self.on_text_change)

    def on_text_change(self, evt):
        text = self.text_ctrl.GetValue().strip()
        text_lower = text.lower()
        if text_lower in CUSTOM_COLORS:
            rgb = CUSTOM_COLORS[text_lower]
            colour = wx.Colour(*rgb)
            self.text_ctrl.Unbind(wx.EVT_TEXT)
            self.SetColour(colour)
            self.text_ctrl.SetValue(text)
            self.text_ctrl.SetInsertionPointEnd()
            self.text_ctrl.Bind(wx.EVT_TEXT, self.on_text_change)

            # Manually emit colour changed event
            evt = wx.ColourPickerEvent()
            evt.SetEventType(wx.wxEVT_COLOURPICKER_CHANGED)
            evt.SetId(self.GetId())
            evt.SetColour(self.GetColour())
            evt.SetEventObject(self)
            wx.PostEvent(self, evt)
        evt.Skip()


class FontListBox(wx.VListBox):
    def __init__(self, parent, font_names, font_data):
        super().__init__(parent)
        self.font_names = font_names
        self.font_data = font_data
        self.SetItemCount(len(font_names))
        self.selected_index = -1
        self.Bind(wx.EVT_LEFT_DOWN, self.OnMouseLeftDown)

    def OnMeasureItem(self, index):
        return 24

    def OnDrawBackground(self, dc, rect, index):
        if self.IsSelected(index):
            dc.SetBrush(wx.Brush(wx.Colour(51, 153, 255)))
            dc.SetPen(wx.Pen(wx.Colour(51, 153, 255)))
            dc.DrawRectangle(rect)
        else:
            dc.SetBrush(wx.Brush(self.GetBackgroundColour()))
            dc.SetPen(wx.Pen(self.GetBackgroundColour()))
            dc.DrawRectangle(rect)

    def OnDrawItem(self, dc, rect, index):
        font_value = self.font_names[index]
        font_name = self.font_data[font_value][0]
        font_data = self.font_data[font_value][1]
        fd = font_data.lower()
        if "extrablack" in fd:
            weight = wx.FONTWEIGHT_EXTRAHEAVY
        elif "black" in fd:
            weight = wx.FONTWEIGHT_HEAVY
        elif "extrabold" in fd:
            weight = wx.FONTWEIGHT_EXTRABOLD
        elif "semibold" in fd or "demibold" in fd:
            weight = wx.FONTWEIGHT_SEMIBOLD
        elif "bold" in fd:
            weight = wx.FONTWEIGHT_BOLD
        elif "medium" in fd:
            weight = wx.FONTWEIGHT_MEDIUM
        elif "extralight" in fd:
            weight = wx.FONTWEIGHT_EXTRALIGHT
        elif "light" in fd:
            weight = wx.FONTWEIGHT_LIGHT
        elif "thin" in fd:
            weight = wx.FONTWEIGHT_THIN
        else:
            weight = wx.FONTWEIGHT_NORMAL

        if "italic" in fd:
            style = wx.FONTSTYLE_ITALIC
        elif "roman" in fd:
            style = wx.FONTSTYLE_SLANT
        else:
            style = wx.FONTSTYLE_NORMAL

        font = wx.Font(14, wx.FONTFAMILY_DEFAULT, style, weight, False, faceName=font_name)
        dc.SetFont(font)
        dc.SetTextForeground(wx.WHITE if self.IsSelected(index) else wx.BLACK)
        text_x = rect.x + 5
        text_y = rect.y + (rect.height - dc.GetTextExtent(font_value)[1]) // 2
        dc.DrawText(font_value, text_x, text_y)

    def OnMouseLeftDown(self, event):
        pos = event.GetPosition()  # relative to the VListBox control
        index = self.VirtualHitTest(pos.y)
        if index != wx.NOT_FOUND:
            self.selected_index = index
            self.SetSelection(index)
            self.Refresh()
        else:
            event.Skip()

    def GetSelectedFontName(self):
        if 0 <= self.selected_index < len(self.font_names):
            return self.font_names[self.selected_index]
        return None


class FontSelectDialog(wx.Dialog):
    def __init__(self, parent, font_names, font_data, initial_selection=None):
        super().__init__(parent, title="Select Font", size=(400, 300))
        self.font_list = FontListBox(self, font_names, font_data)

        if initial_selection in font_names:
            idx = font_names.index(initial_selection)
            self.font_list.selected_index = idx
            self.font_list.SetSelection(idx)

        select_btn = wx.Button(self, label="Select")
        cancel_btn = wx.Button(self, label="Cancel")

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(select_btn)
        btn_sizer.Add(cancel_btn)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(self.font_list, 1, wx.EXPAND | wx.ALL, 10)
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        self.SetSizer(main_sizer)

        select_btn.Bind(wx.EVT_BUTTON, self.OnSelect)
        cancel_btn.Bind(wx.EVT_BUTTON, self.OnCancel)

    def OnSelect(self, event):
        if self.font_list.GetSelectedFontName() is not None:
            self.EndModal(wx.ID_OK)
        else:
            wx.MessageBox("Please select a font before clicking Select.", "No Selection", wx.OK | wx.ICON_WARNING)

    def OnCancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def GetSelectedFontName(self):
        return self.font_list.GetSelectedFontName()


class FontPickerCtrl(wx.Panel):
    def __init__(self, parent, name, initial_font=None):
        super().__init__(parent, name=name)
        self.parent = parent

        sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.font_display = wx.TextCtrl(self, style=wx.TE_READONLY)
        self.font_display.SetMinSize((400, 40))
        self.pick_btn = wx.Button(self, label="Choose Font")
        self.pick_btn.SetMinSize((-1, 40))
        sizer.Add(self.font_display, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.pick_btn, 0, wx.ALL, 5)
        self.SetSizer(sizer)

        self.font_names = []
        self.font_dict = {}
        self.font_data = {}
        for elem in font_manager.findSystemFonts():
            fontname = ImageFont.FreeTypeFont(elem).getname()
            name = " ".join(fontname)
            self.font_names.append(name)
            self.font_dict[name] = elem
            self.font_data[name] = fontname
        self.font_names = list(set(self.font_names))
        self.font_names.sort()
        self.current_font_name = initial_font or (
            "Arial Regular" if "Arial Regular" in self.font_names else (self.font_names[0] if self.font_names else "Default"))
        self.current_font = wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False,
                                    self.current_font_name)
        self.UpdateFontDisplay()

        self.pick_btn.Bind(wx.EVT_BUTTON, self.OnPickFont)

    def OnPickFont(self, event):
        dlg = FontSelectDialog(self, self.font_names, self.font_data, initial_selection=self.current_font_name)
        if dlg.ShowModal() == wx.ID_OK:
            selected = dlg.GetSelectedFontName()
            if selected:
                self.current_font_name = selected
                name,fd = self.font_data[self.current_font_name]
                fd = fd.lower()
                if "extrablack" in fd:
                    weight = wx.FONTWEIGHT_EXTRAHEAVY
                elif "black" in fd:
                    weight = wx.FONTWEIGHT_HEAVY
                elif "extrabold" in fd:
                    weight = wx.FONTWEIGHT_EXTRABOLD
                elif "semibold" in fd or "demibold" in fd:
                    weight = wx.FONTWEIGHT_SEMIBOLD
                elif "bold" in fd:
                    weight = wx.FONTWEIGHT_BOLD
                elif "medium" in fd:
                    weight = wx.FONTWEIGHT_MEDIUM
                elif "extralight" in fd:
                    weight = wx.FONTWEIGHT_EXTRALIGHT
                elif "light" in fd:
                    weight = wx.FONTWEIGHT_LIGHT
                elif "thin" in fd:
                    weight = wx.FONTWEIGHT_THIN
                else:
                    weight = wx.FONTWEIGHT_NORMAL

                if "italic" in fd:
                    style = wx.FONTSTYLE_ITALIC
                elif "roman" in fd:
                    style = wx.FONTSTYLE_SLANT
                else:
                    style = wx.FONTSTYLE_NORMAL
                if self.current_font:
                    size = self.current_font.GetPointSize()
                else:
                    size = 20

                self.current_font = wx.Font(size, wx.FONTFAMILY_DEFAULT, style, weight, False, name)
                self.UpdateFontDisplay()
        dlg.Destroy()

    def UpdateFontDisplay(self):
        self.font_display.SetValue(self.current_font_name)
        self.font_display.SetFont(self.current_font)
        text_height = self.font_display.GetTextExtent("Ag")[1] + 10
        self.font_display.SetMinSize((400, text_height))
        self.GetSizer().Layout()
        self.parent.Layout()
        self.Refresh()

    def GetSelectedFont(self):
        file = self.font_dict[self.current_font_name]
        return file


class SaveFilePicker(wx.Panel):
    def __init__(self, parent, name, message="Choose location for log file", wildcard="*.*"):
        super().__init__(parent, name=name)
        self.message = message
        self.wildcard = wildcard

        sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.text_ctrl = wx.TextCtrl(self)
        self.browse_btn = wx.Button(self, label="Browse")
        sizer.Add(self.text_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.browse_btn, 0, wx.ALL, 5)
        self.SetSizer(sizer)

        self.browse_btn.Bind(wx.EVT_BUTTON, self.OnBrowse)

    def OnBrowse(self, event):
        dlg = wx.FileDialog(self, message=self.message, wildcard=self.wildcard,
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            self.text_ctrl.SetValue(dlg.GetPath())
        dlg.Destroy()

    def GetPath(self):
        return self.text_ctrl.GetValue()


class GlobalClickPicker(wx.Panel):
    def __init__(self, parent, name):
        super().__init__(parent, name=name)

        sizer = wx.BoxSizer(wx.VERTICAL)

        self.text_x = wx.TextCtrl(self)
        self.text_y = wx.TextCtrl(self)

        self.text_x.SetValue("20")
        self.text_y.SetValue("120")


        self.pick_button = wx.Button(self, label="Pick Overlay Position")
        self.pick_button.Bind(wx.EVT_BUTTON, self.OnPickButton)

        sizer.Add(wx.StaticText(self, label="X Coordinate:"), 0, wx.ALL, 5)
        sizer.Add(self.text_x, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(wx.StaticText(self, label="Y Coordinate:"), 0, wx.ALL, 5)
        sizer.Add(self.text_y, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.pick_button, 0, wx.ALL | wx.CENTER, 5)

        self.SetSizer(sizer)

        self.listener = None
        self.listening = False

    def OnPickButton(self, event):
        if not self.listening:
            self.pick_button.SetLabel("Waiting for click...")
            self.listening = True
            self.listener = mouse.Listener(on_click=self.OnGlobalClick)
            self.listener.start()

    def OnGlobalClick(self, x, y, button, pressed):
        if pressed:
            if self.listener:
                self.listener.stop()
                self.listener = None
            wx.CallAfter(self.SetCoordinates, x, y)

    def SetCoordinates(self, x, y):
        self.text_x.SetValue(str(x))
        display_size = wx.GetDisplaySize()
        y = display_size[1] - y
        self.text_y.SetValue(str(y))
        self.pick_button.SetLabel("Pick Overlay Position")
        self.listening = False

    def GetValue(self):
        try:
            return int(self.text_x.GetValue()), int(self.text_y.GetValue())
        except:
            return False


class HotkeyCtrl(wx.TextCtrl):
    def __init__(self, parent):
        super().__init__(parent, style=wx.TE_READONLY)
        self.parent = parent
        self.capturing = False
        self.hotkey_str = ""
        self.logger = HotkeyLogger()
        self.Bind(wx.EVT_LEFT_DOWN, self.on_click)

    def on_click(self, event):
        if not self.capturing:
            print("starting capture")
            self.StartCapture()
        event.Skip()

    def StartCapture(self):
        if not self.parent.GetParent().GetParent().GetParent().capturinghotkey:
            self.SetValue("Press hotkey...")
            self.capturing = True
            self.parent.GetParent().GetParent().GetParent().capturinghotkey = True

            def capture_thread():
                hotkey = self.logger.start_capture()
                wx.CallAfter(self.finish_capture, hotkey)

            threading.Thread(target=capture_thread, daemon=True).start()

    def finish_capture(self, hotkey):
        self.capturing = False
        self.parent.GetParent().GetParent().GetParent().capturinghotkey = False

        self.SetValue(hotkey)
