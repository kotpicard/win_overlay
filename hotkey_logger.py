from pynput import keyboard


class HotkeyLogger:
    def __init__(self):
        self.pressed_keys = set()
        self.last_combo = set()
        self.listener = None
        self.captured_combo = None
        self.mods = ["Ctrl", "Shift", "Alt"]

    def key_name(self, key):
        """Return readable name + vk number for any key."""
        print(key)
        if isinstance(key, keyboard.Key):
            print("aaaa")
            if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
                return "Ctrl"
            elif key in (keyboard.Key.alt_l, keyboard.Key.alt_r):
                return "Alt"
            elif key in (keyboard.Key.shift_l, keyboard.Key.shift_r):
                return "Shift"
            else:
                return str(key)
        elif isinstance(key, keyboard.KeyCode):
            if key.char:
                return f"{chr(key.vk)}"
            else:
                num = int(str(key)[1:-1])
                return chr(num)
        else:
            return chr(key[1:-1])

    def on_press(self, key):
        self.pressed_keys.add(key)
        self.last_combo = self.pressed_keys.copy()

    def on_release(self, key):
        self.pressed_keys.discard(key)
        if not self.pressed_keys and self.last_combo:
            mod_names = [self.key_name(k) for k in self.last_combo if self.key_name(k) in self.mods]
            key_name = [self.key_name(k) for k in self.last_combo if self.key_name(k) not in self.mods]
            ordered_mods = [x for x in self.mods if x in mod_names]
            self.captured_combo = " + ".join(ordered_mods + key_name)
            print("Captured hotkey:", self.captured_combo)
            self.last_combo.clear()

            # Stop listener after first capture
            if self.listener:
                self.listener.stop()

    def start_capture(self):
        with keyboard.Listener(
                on_press=self.on_press,
                on_release=self.on_release) as self.listener:
            self.listener.join()
        return self.captured_combo


if __name__ == "__main__":
    print("Press your hotkey combination...")
    combo = HotkeyLogger().start_capture()
    print("Final combo was:", combo)
