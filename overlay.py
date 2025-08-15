import ctypes
import win32gui
import win32con
import win32api
from wintypestructs import BITMAPINFO, BITMAPINFOHEADER
from textimagecreator import create_text_image
from hotkeys import *
import datetime
import threading
import atexit


class OverlayWindow:
    def __init__(self, extended_style, class_name, window_name, style, x_pos, y_pos, width, height, parent_window,
                 h_menu, lp_void, text_settings, texts, hotkey_list, simultaneous_change, simulated_hotkeys, create_log,
                 log_path):
        self.hotkey_list = hotkey_list
        self.simultaneous_change = simultaneous_change
        self.simulated_hotkeys = simulated_hotkeys
        self.create_log = create_log
        self.log_path = log_path
        self.log_file = None

        self.extended_style = extended_style
        self.class_name = class_name
        self.window_name = window_name
        self.style = style
        self.x_pos = x_pos
        self.y_pos = y_pos
        self.width = width
        self.height = height
        self.parent_window = parent_window
        self.h_menu = h_menu
        self.h_instance = win32api.GetModuleHandle(None)
        self.lp_void = lp_void

        self.font_path = text_settings["font_path"]
        self.font_size = text_settings["font_size"]
        self.text_color = text_settings["text_color"]
        self.text_x_pos = text_settings["x_pos"]
        self.text_y_pos = -1 * text_settings["y_pos"]

        self.texts = texts
        self.counter = 0
        self.current_text = self.texts[self.counter]

        # Thread safety and cleanup
        self.stop_event = threading.Event()
        self.window = None
        self.registered_class = None
        self._cleaned_up = False

        # Register cleanup function
        atexit.register(self._clean_up)

        try:
            self.registered_class = self.RegisterClass()
            self.window = self.CreateWindow()
            self.ShowWindow()
            self.CreateOverlayContent(self.current_text)
            self.starttime = datetime.datetime.now()
        except Exception as e:
            print(f"Error initializing OverlayWindow: {e}")
            self._clean_up()
            raise

    def Stop(self):
        """Signal the message loop to exit and destroy the window."""
        print("OverlayWindow: Stop signal received.")
        self.stop_event.set()
        #
        # # Post quit message to exit the message loop
        # try:
        #     win32gui.PostQuitMessage(0)
        # except Exception as e:
        #     print(f"Error posting quit message: {e}")

    def CreateWindow(self):
        return win32gui.CreateWindowEx(
            self.extended_style,
            self.registered_class,
            self.window_name,
            self.style,
            self.x_pos,
            self.y_pos,
            self.width,
            self.height,
            self.parent_window,
            self.h_menu,
            self.h_instance,
            self.lp_void
        )

    def RegisterClass(self):
        wndClass = win32gui.WNDCLASS()
        wndClass.lpfnWndProc = self.ProcessMessages
        wndClass.hInstance = self.h_instance
        wndClass.lpszClassName = self.class_name
        wndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        return win32gui.RegisterClass(wndClass)

    def HideOverlay(self):
        if self.window:
            win32gui.ShowWindow(self.window, win32con.SW_HIDE)

    def ShowOverlay(self):
        if self.window:
            win32gui.ShowWindow(self.window, win32con.SW_SHOW)

    def ProcessMessages(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == HOTKEY_HIDE:
                self.HideOverlay()
            elif wparam == HOTKEY_SHOW:
                self.ShowOverlay()
            elif wparam == HOTKEY_NEXTTEXT:
                if self.simultaneous_change:
                    self.SimulateHotkey(self.simulated_hotkeys[HOTKEY_NEXTSLIDE])
                self.UpdateCounterAndText(1)
                if self.create_log:
                    self.CreateLog(datetime.datetime.now())
            elif wparam == HOTKEY_PREVTEXT:
                if self.simultaneous_change:
                    self.SimulateHotkey(self.simulated_hotkeys[HOTKEY_PREVSLIDE])
                self.UpdateCounterAndText(-1)
                if self.create_log:
                    self.CreateLog(datetime.datetime.now())
            elif wparam == HOTKEY_STARTTIMER:
                self.starttime = datetime.datetime.now()
                if self.create_log:
                    print("Starting log at", self.starttime)
                    self.CreateLog(self.starttime)
        elif msg == win32con.WM_DESTROY:
            print("OverlayWindow: WM_DESTROY received")
            win32gui.PostQuitMessage(0)

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def CreateLog(self, endtime):
        if not self.create_log:
            return

        try:
            with open(self.log_path, "a") as fh:
                timestamp = self.calc_time(endtime)
                if timestamp != "0:00:00":
                    try:
                        fh.write("—" + timestamp + "\n")
                    except Exception as e:
                        print(f"Cannot write timestamp to file: {e}")
                        self.create_log = False
                        return

                try:
                    fh.write(self.current_text + ": " + timestamp + "\n")
                except Exception as e:
                    print(f"Cannot write text to file: {e}")
                    self.create_log = False
        except Exception as e:
            print(f"Cannot open log file: {e}")
            self.create_log = False

    def calc_time(self, endtime):
        td = endtime - self.starttime
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        formatted = f"{hours}:{minutes:02d}:{seconds:02d}"
        return formatted

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
                    vk = ord(key.upper())  # Convert to uppercase for consistency

            result = win32gui.RegisterHotKey(None, hotkey_id, modifiers, vk)
            if not result:
                print(f"Failed to register hotkey {hotkey_id} with keys {keys}")
            else:
                print(f"Successfully registered hotkey {hotkey_id} with keys {keys}")
        except Exception as e:
            print(f"Error registering overlay hotkey {hotkey_id} with keys {keys}: {e}")

    def UnregisterHotkeys(self):
        """Unregister all hotkeys"""
        try:
            for hotkey in self.hotkey_list:
                hotkey_id = hotkey[0]
                result = win32gui.UnregisterHotKey(None, hotkey_id)
                if result:
                    print(f"Unregistered overlay hotkey {hotkey_id}")
                else:
                    print(f"Failed to unregister overlay hotkey {hotkey_id}")
        except Exception as e:
            print(f"Error unregistering overlay hotkeys: {e}")

    def SimulateHotkey(self, keys):
        if not keys:
            return

        keys_to_press = []
        for key in keys:
            if key == "Ctrl":
                keys_to_press.append(win32con.VK_CONTROL)
            elif key == "Alt":
                keys_to_press.append(win32con.VK_MENU)  # Alt key
            elif key == "Shift":
                keys_to_press.append(win32con.VK_SHIFT)
            elif key == "Win":
                keys_to_press.append(win32con.VK_LWIN)
            else:
                keys_to_press.append(ord(key.upper()))

        # Press keys down
        for key in keys_to_press:
            win32api.keybd_event(key, 0, 0, 0)
        # Release keys in reverse order
        for key in reversed(keys_to_press):
            win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)

    def ShowWindow(self):
        if self.window:
            win32gui.ShowWindow(self.window, win32con.SW_SHOWNORMAL)

    def UpdateCounterAndText(self, amt):
        self.counter += amt
        if self.counter == -1:
            self.counter = 0  # can't go back from first text
        elif self.counter == len(self.texts):
            self.counter -= 1  # can't go forward from last text
        else:
            self.current_text = self.texts[self.counter]
            self.CreateOverlayContent(self.current_text)

    def ConvertImageToBitmap(self, img):
        hdc = win32gui.GetDC(0)
        memdc = win32gui.CreateCompatibleDC(hdc)
        width, height = img.size
        raw_data = img.tobytes("raw", "BGRA")

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = win32con.BI_RGB
        bmi.bmiHeader.biSizeImage = 0
        bmi.bmiHeader.biXPelsPerMeter = 0
        bmi.bmiHeader.biYPelsPerMeter = 0
        bmi.bmiHeader.biClrUsed = 0
        bmi.bmiHeader.biClrImportant = 0
        ppvBits = ctypes.c_void_p()
        hbitmap = ctypes.windll.gdi32.CreateDIBSection(
            memdc, ctypes.byref(bmi), win32con.DIB_RGB_COLORS,
            ctypes.byref(ppvBits), None, 0
        )
        if not hbitmap:
            raise ctypes.WinError()

        ctypes.memmove(ppvBits, raw_data, len(raw_data))

        return hdc, memdc, hbitmap

    def CreateOverlayContent(self, text):
        if not self.window:
            return

        try:
            img = create_text_image(self.width, self.height, text, self.text_color, self.font_path, self.font_size)
            hdc, memdc, hbitmap = self.ConvertImageToBitmap(img)
            win32gui.SelectObject(memdc, hbitmap)
            blend = (win32con.AC_SRC_OVER, 0, 255, win32con.AC_SRC_ALPHA)

            win32gui.UpdateLayeredWindow(
                self.window,
                hdc,
                (self.text_x_pos, self.text_y_pos),
                (self.width, self.height),
                memdc,
                (0, 0),
                0,
                blend,
                win32con.ULW_ALPHA
            )

            win32gui.ReleaseDC(self.window, hdc)
            win32gui.DeleteObject(hbitmap)
            win32gui.DeleteDC(memdc)
        except Exception as e:
            print(f"Error creating overlay content: {e}")

    def Run(self):
        try:
            # Register hotkeys
            for hotkey in self.hotkey_list:
                hotkey_id = hotkey[0]
                hotkey_keys = hotkey[1]
                print(f"Registering overlay hotkey {hotkey_id}, {hotkey_keys}")
                self.RegisterHotkey(hotkey_id, hotkey_keys)

            # Message loop with timeout to check stop_event periodically
            while not self.stop_event.is_set():
                # Use PeekMessage with timeout instead of GetMessage to avoid blocking indefinitely
                ret = win32gui.PeekMessage(None, 0, 0, win32con.PM_REMOVE)
                if ret[0]:  # Message available
                    msg = ret[1]
                    if msg[1] == win32con.WM_QUIT:
                        print("OverlayWindow: WM_QUIT received.")
                        break
                    self.ProcessMessages(None, msg[1], msg[2], msg[3])
                    win32gui.TranslateMessage(msg)
                    win32gui.DispatchMessage(msg)
                else:
                    # No message available, sleep briefly to avoid busy waiting
                    import time
                    time.sleep(0.01)  # 10ms sleep

        except KeyboardInterrupt:
            print("Keyboard interrupt received in overlay")
        except Exception as e:
            print(f"Error in overlay Run(): {e}")
        finally:
            print("OverlayWindow: Run() finishing, calling cleanup")
            self._clean_up()

    def _clean_up(self):
        if self._cleaned_up:
            return  # Already cleaned up

        print("OverlayWindow: Cleaning up overlay window.")
        self._cleaned_up = True

        # Set stop event to ensure any running loops exit
        self.stop_event.set()

        # Unregister hotkeys first
        self.UnregisterHotkeys()

        # Close and destroy the window
        try:
            if self.window:
                print("Hiding and destroying overlay window...")
                self.HideOverlay()

                # Destroy the window
                result = win32gui.DestroyWindow(self.window)
                if not result:
                    print("Warning: Failed to destroy window")
                self.window = None

        except Exception as e:
            print(f"Error destroying window: {e}")

        # Unregister the window class
        try:
            if self.registered_class and self.h_instance:
                result = win32gui.UnregisterClass(self.class_name, self.h_instance)
                if not result:
                    print("Warning: Failed to unregister window class")
                self.registered_class = None
        except Exception as e:
            print(f"Error unregistering window class: {e}")

        print("OverlayWindow: Cleanup completed.")


if __name__ == "__main__":
    try:
        window = OverlayWindow(
            extended_style=win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
            class_name="OverlayWindow",
            window_name="OverlayWindow",
            style=win32con.WS_POPUP,
            x_pos=0,
            y_pos=win32api.GetSystemMetrics(1) - 200,
            width=win32api.GetSystemMetrics(0),
            height=100,
            parent_window=None,
            h_menu=None,
            lp_void=None,
            text_settings={
                "font_path": r"C:\Users\adria\PycharmProjects\pythonProject3\overlay\Cera-Pro-Regular.otf",
                "font_size": 20,
                "text_color": (120, 138, 168),
                "x_margin": int(win32api.GetSystemMetrics(0) * 0.98),
                "y_margin": 20,
                "x_pos": 0,
                "y_pos": win32api.GetSystemMetrics(1) - 200
            },
            texts=["text1", "text2", "text3"],
            hotkey_list=[
                (HOTKEY_NEXTTEXT, ["Ctrl", "Shift", "T"]),
                (HOTKEY_PREVTEXT, ["Ctrl", "Shift", "R"]),
                (HOTKEY_SHOW, ["Ctrl", "Shift", "W"]),
                (HOTKEY_HIDE, ["Ctrl", "Shift", "Q"])
            ],
            simultaneous_change=False,
            simulated_hotkeys=None,
            create_log=False,
            log_path=""
        )

        window.Run()  # This will handle the message loop and cleanup

    except KeyboardInterrupt:
        print("Program interrupted by user")
    except Exception as e:
        print(f"Error in main: {e}")
    finally:
        print("Overlay program finished")
