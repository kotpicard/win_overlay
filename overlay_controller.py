import overlay
import pptcontroller
import threading
import json
from hotkeys import *
import pythoncom
import win32con
import win32api
import win32gui
import sys


class OverlayController:
    def __init__(self, config_path):
        # Get config
        self.general_config, self.text_config, self.ppt_config = self.ReadConfig(config_path)
        self.additional_hotkeys_for_overlay = None
        self.additional_hotkeys_for_ppt = None
        self.stopped = False
        self.Run()

    def stop(self):
        print("Shutting down...")

        try:
            win32gui.PostQuitMessage(0)  # Exit message loop
        except Exception as e:
            print("Error posting quit message:", e)

        for hotkey_id in [HOTKEY_QUIT, HOTKEY_STARTTIMER]:
            try:
                win32gui.UnregisterHotKey(None, hotkey_id)
            except Exception as e:
                print(f"Failed to unregister hotkey {hotkey_id}: {e}")

        # Clean up overlay window
        if hasattr(self, "overlay_window"):
            try:
                self.overlay_window.Stop()
            except Exception as e:
                print("Error stopping overlay:", e)
            if self.overlay_thread and self.overlay_thread.is_alive():
                self.overlay_thread.join(timeout=5.0)  # Wait up to 5 seconds

                if self.overlay_thread.is_alive():
                    print("Warning: Overlay thread didn't stop cleanly")

        # Stop PPT controller
        if hasattr(self, "ppt_controller"):
            try:
                self.ppt_controller.Stop()
            except Exception as e:
                print("Error stopping ppt controller:", e)
            if self.ppt_thread and self.ppt_thread.is_alive():
                self.ppt_thread.join(timeout=5.0)  # Wait up to 5 seconds

                if self.ppt_thread.is_alive():
                    print("Warning: PPT thread didn't stop cleanly")
        print("Shutting down")
        sys.exit(0)

    def ReadConfig(self, config_path):
        fh = open(config_path)
        config = json.load(fh)
        fh.close()
        return config["general"], config["text"], config["ppt"]

    def GetHotkeys(self):
        return [[HOTKEY_QUIT, self.general_config["hotkeys"][str(HOTKEY_QUIT)]]]

    def ProcessMessages(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == HOTKEY_QUIT:
                print("quit")
                self.stop()
            elif wparam == HOTKEY_STARTTIMER:
                print("Start timer")
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

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
            print(f"Registering controller hotkey {hotkey_id} with keys {keys}")
            result = win32gui.RegisterHotKey(None, hotkey_id, modifiers, vk)
            if not result:
                print(f"Failed to register controller hotkey {hotkey_id} with keys {keys}")
        except Exception as e:
            print(f"Error registering controller hotkey {hotkey_id} with keys {keys}: {e}")

    def GetTextsFromFile(self):
        fh = open(self.general_config["texts_path"])
        texts = fh.readlines()
        texts = [x.strip() for x in texts]
        fh.close()
        return texts

    def RunOverlay(self):
        def overlay_thread_target():
            import pythoncom
            pythoncom.CoInitialize()
            overlay_hotkeys = [
                [HOTKEY_SHOW, self.general_config["hotkeys"][str(HOTKEY_SHOW)]],
                [HOTKEY_HIDE, self.general_config["hotkeys"][str(HOTKEY_HIDE)]],
                [HOTKEY_NEXTTEXT, self.general_config["hotkeys"][str(HOTKEY_NEXTTEXT)]],
                [HOTKEY_PREVTEXT, self.general_config["hotkeys"][str(HOTKEY_PREVTEXT)]],
                [HOTKEY_STARTTIMER, self.general_config["hotkeys"][str(HOTKEY_STARTTIMER)]]
            ]

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
                text_settings=self.text_config,
                texts=self.overlay_texts,
                hotkey_list=overlay_hotkeys,
                simultaneous_change=self.general_config["simultaneous_change"],
                simulated_hotkeys=self.additional_hotkeys_for_overlay,
                create_log=self.general_config["create_log"],
                log_path=self.general_config["log_path"]
            )

            self.overlay_window.Run()

        self.overlay_thread = threading.Thread(
            target=overlay_thread_target,
            daemon=False
        )
        self.overlay_thread.start()

    def RunPPTController(self):
        def ppt_thread_target():
            pythoncom.CoInitialize()

            ppt_hotkey_list = [
                [HOTKEY_SHOWPPT, self.general_config["hotkeys"][str(HOTKEY_SHOWPPT)]],
                [HOTKEY_HIDEPPT, self.general_config["hotkeys"][str(HOTKEY_HIDEPPT)]],
                [HOTKEY_NEXTSLIDE, self.general_config["hotkeys"][str(HOTKEY_NEXTSLIDE)]],
                [HOTKEY_PREVSLIDE, self.general_config["hotkeys"][str(HOTKEY_PREVSLIDE)]]
            ]

            self.ppt_controller = pptcontroller.PPTController(self.ppt_config["ppt_path"], ppt_hotkey_list,
                                                              self.general_config["toggle_overlay_with_ppt"],
                                                              self.additional_hotkeys_for_ppt)
            self.ppt_ready_event.set()
            self.ppt_controller.Run()

        self.ppt_thread = threading.Thread(target=ppt_thread_target, daemon=False)
        self.ppt_thread.start()

    def Run(self):
        # Run PPTController if configured,otherwise get texts from provided file
        if self.general_config["use_ppt"]:
            print("HERE")
            if self.general_config["toggle_overlay_with_ppt"]:
                print("HERE")
                self.additional_hotkeys_for_ppt = {
                    HOTKEY_SHOW: self.general_config["hotkeys"][str(HOTKEY_SHOW)],
                    HOTKEY_HIDE: self.general_config["hotkeys"][str(HOTKEY_HIDE)]

                }
            self.ppt_controller = None
            self.ppt_ready_event = threading.Event()
            self.RunPPTController()
            self.ppt_ready_event.wait(timeout=5)  # wait max 5s
            self.overlay_texts = self.ppt_controller.slides
            # If changing texts should change slides automatically, create configuration to pass to overlay
            if self.general_config["simultaneous_change"]:
                self.additional_hotkeys_for_overlay = {
                    HOTKEY_NEXTSLIDE: self.general_config["hotkeys"][str(HOTKEY_NEXTSLIDE)],
                    HOTKEY_PREVSLIDE: self.general_config["hotkeys"][str(HOTKEY_PREVSLIDE)]

                }
        else:
            self.overlay_texts = self.GetTextsFromFile()
            self.general_config["toggle_overlay_with_ppt"] = False
            self.general_config["simultaneous_change"] = False

        self.RunOverlay()
        self.hotkey_list = self.GetHotkeys()
        for key in self.hotkey_list:
            self.RegisterHotkey(key[0], key[1])
        while True:
            ret, msg = win32gui.GetMessage(None, 0, 0)
            if ret > 0:
                self.ProcessMessages(None, msg[1], msg[2], msg[3])
                win32gui.TranslateMessage(msg)
                win32gui.DispatchMessage(msg)


if __name__ == "__main__":
    controller = OverlayController("settings.json")
