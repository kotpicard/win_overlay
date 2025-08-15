import win32api
import win32con
import win32gui
import win32com.client
from hotkeys import *
import threading
import atexit


class PPTController:
    def __init__(self, ppt_path, hotkey_list=None, toggle_overlay_with_ppt=False, simulated_hotkeys=None):
        self.hotkey_list = hotkey_list or []
        self.toggle_overlay_with_ppt = toggle_overlay_with_ppt
        self.simulated_hotkeys = simulated_hotkeys
        self.ppt_path = ppt_path

        self.stop_event = threading.Event()
        self.ppt_app = None
        self.presentation = None
        self.window = None
        self._cleaned_up = False  # Flag to prevent multiple cleanup calls

        # Register cleanup function to run on exit
        atexit.register(self._clean_up)

        try:
            self.OpenPPT()
            self.current_slide = 0
            self.slides = self.GetSectionsAndTitles()
        except Exception as e:
            print(f"Error initializing PPTController: {e}")
            self._clean_up()
            raise

    def Stop(self):
        print("PPTController: Stop signal received.")
        self.stop_event.set()
        win32gui.PostQuitMessage(0)

    def OpenPPT(self):
        try:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            self.ppt_app.Visible = True
            self.presentation = self.ppt_app.Presentations.Open(self.ppt_path, WithWindow=True)
            self.window = win32gui.FindWindow("PPTFrameClass", None)
        except Exception as e:
            print(f"Error opening PowerPoint: {e}")
            raise

    def GetSectionsAndTitles(self):
        if not self.presentation:
            return []

        try:
            section_props = self.presentation.SectionProperties
            result = []

            for i in range(section_props.Count):
                section_name = section_props.Name(i + 1)
                first_slide = section_props.FirstSlide(i + 1)
                num_slides = section_props.SlidesCount(i + 1)

                for j in range(first_slide, first_slide + num_slides):
                    slide = self.presentation.Slides(j)

                    title = None
                    for shape in slide.Shapes:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            title = shape.TextFrame.TextRange.Text.strip()
                            break

                    title = title or f"Untitled Slide {j}"
                    if title == section_name:
                        result.append(f"Section {i}: {section_name}")
                    else:
                        result.append(f"{section_name}: {title}")

            return result
        except Exception as e:
            print(f"Error getting sections and titles: {e}")
            return []

    def ShowPPT(self):
        if self.window:
            win32gui.ShowWindow(self.window, win32con.SW_RESTORE)
            win32gui.ShowWindow(self.window, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(self.window)

    def HidePPT(self):
        if self.window:
            win32gui.ShowWindow(self.window, win32con.SW_MAXIMIZE)
            win32gui.SetWindowPos(
                self.window,
                win32con.HWND_BOTTOM,
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            )

    def UpdateSlide(self, amt):
        self.current_slide += amt
        if self.current_slide == -1:
            self.current_slide = 0  # can't go back from first text
        elif self.current_slide == len(self.slides):
            self.current_slide -= 1  # can't go forward from last text
        else:
            self.MoveToSlide(self.current_slide)

    def MoveToSlide(self, number):
        if self.presentation and self.ppt_app:
            try:
                slide = self.presentation.Slides(number + 1)  # ppt is indexed from 1
                view = self.ppt_app.ActiveWindow.View
                view.GotoSlide(slide.SlideIndex)
            except Exception as e:
                print(f"Error moving to slide: {e}")

    def ProcessMessages(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == HOTKEY_SHOWPPT:
                if self.toggle_overlay_with_ppt:
                    self.SimulateHotkey(self.simulated_hotkeys[HOTKEY_HIDE])
                self.ShowPPT()
            elif wparam == HOTKEY_HIDEPPT:
                if self.toggle_overlay_with_ppt:
                    self.SimulateHotkey(self.simulated_hotkeys[HOTKEY_SHOW])
                self.HidePPT()
            elif wparam == HOTKEY_NEXTSLIDE:
                print("Got Hotkey Next")
                self.UpdateSlide(1)
            elif wparam == HOTKEY_PREVSLIDE:
                print("Got Hotkey Prev")
                self.UpdateSlide(-1)

    def RegisterHotkey(self, hotkey_id, keys):
        try:
            modifiers = 0
            vk = 0
            for key in keys:
                if key == "Shift":
                    modifiers |= win32con.MOD_SHIFT
                elif key == "Alt":
                    modifiers |= win32con.MOD_ALT
                elif key == "Ctrl":
                    modifiers |= win32con.MOD_CONTROL
                elif key == "Win":
                    modifiers |= win32con.MOD_WIN
                else:
                    vk = ord(key)
            print(f"Registering ppt hotkey {hotkey_id} with keys {keys}")
            result = win32gui.RegisterHotKey(None, hotkey_id, modifiers, vk)
            if not result:
                print(f"Failed to register ppt hotkey {hotkey_id} with keys {keys}")
        except Exception as e:
            print(f"Error registering ppt hotkey {hotkey_id} with keys {keys}: {e}")

    def UnregisterHotkeys(self):
        """Unregister all hotkeys"""
        try:
            for hotkey in self.hotkey_list:
                hotkey_id = hotkey[0]
                win32gui.UnregisterHotKey(None, hotkey_id)
                print(f"Unregistered ppt hotkey {hotkey_id}")
        except Exception as e:
            print(f"Error unregistering ppt hotkeys: {e}")

    def SimulateHotkey(self, keys):
        if not keys:
            return

        keys_to_press = []
        for key in keys:
            if key == "Ctrl":
                keys_to_press.append(win32con.VK_CONTROL)
            elif key == "Alt":
                keys_to_press.append(win32con.VK_MENU)  # alt key
            elif key == "Shift":
                keys_to_press.append(win32con.VK_SHIFT)
            elif key == "Win":
                keys_to_press.append(win32con.VK_LWIN)
            else:
                keys_to_press.append(ord(key.upper()))

        # Press keys down
        for key in keys_to_press:
            win32api.keybd_event(key, 0, 0, 0)
        # Release keys
        for key in keys_to_press:
            win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)

    def Run(self):
        try:
            # Register hotkeys
            for hotkey in self.hotkey_list:
                hotkey_id = hotkey[0]
                hotkey_keys = hotkey[1]
                print(f"Registering {hotkey_id}, {hotkey_keys}")
                self.RegisterHotkey(hotkey_id, hotkey_keys)

            while not self.stop_event.is_set():
                ret = win32gui.PeekMessage(None, 0, 0, win32con.PM_REMOVE)
                if ret[0]:  # Message available
                    msg = ret[1]
                    if msg[1] == win32con.WM_QUIT:
                        print("PPTController: WM_QUIT received.")
                        break
                    self.ProcessMessages(None, msg[1], msg[2], msg[3])
                    win32gui.TranslateMessage(msg)
                    win32gui.DispatchMessage(msg)
                else:
                    # No message available
                    import time
                    time.sleep(0.01)

        except KeyboardInterrupt:
            print("Keyboard interrupt received")
        except Exception as e:
            print(f"Error in Run(): {e}")
        finally:
            print("PPTController: Run() finishing, calling cleanup")
            self._clean_up()

    def _clean_up(self):
        if self._cleaned_up:
            return  # Already cleaned up

        print("PPTController: Cleaning up PowerPoint.")
        self._cleaned_up = True
        self.stop_event.set()

        # Unregister hotkeys first
        self.UnregisterHotkeys()

        # Close PowerPoint
        try:
            if self.ppt_app:
                self.ppt_app.DisplayAlerts = False

                if self.presentation:
                    print("Closing presentation...")
                    self.presentation.Close()
                    self.presentation = None

                print("Quitting PowerPoint application...")
                self.ppt_app.Quit()

                # Release COM objects
                del self.ppt_app
                self.ppt_app = None

        except Exception as e:
            print(f"Error during cleanup: {e}")

        # Force garbage collection to help release COM objects
        import gc
        gc.collect()
        print("PPTController: Cleanup completed.")


if __name__ == "__main__":
    try:
        pptc = PPTController(r"C:\Users\adria\PycharmProjects\pythonProject3\testsections.pptx")

        hotkeys = [
            (HOTKEY_NEXTSLIDE, ["Ctrl", "Shift", "F"]),
            (HOTKEY_PREVSLIDE, ["Ctrl", "Shift", "D"]),
            (HOTKEY_SHOWPPT, ["Ctrl", "Shift", "G"]),
            (HOTKEY_HIDEPPT, ["Ctrl", "Shift", "A"]),
        ]

        pptc.hotkey_list = hotkeys
        pptc.Run()  # This will handle the message loop and cleanup

    except KeyboardInterrupt:
        print("Program interrupted by user")
    except Exception as e:
        print(f"Error in main: {e}")
    finally:
        print("Program finished")
