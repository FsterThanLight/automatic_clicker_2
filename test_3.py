import ctypes


def set_caps_lock_state(activate):
    """Set the state of the Caps Lock key.

    :param activate: If True, the Caps Lock key is activated. If False, it is deactivated.
    """
    user32 = ctypes.WinDLL('user32.dll')
    VK_CAPITAL = 0x14

    is_caps_lock_on = user32.GetKeyState(VK_CAPITAL) != 0

    if activate and not is_caps_lock_on:
        # Caps Lock is not on, but we want to activate it
        user32.keybd_event(VK_CAPITAL, 0, 0, 0)  # Press the Caps Lock key
        user32.keybd_event(VK_CAPITAL, 0, 2, 0)  # Release the Caps Lock key
    elif not activate and is_caps_lock_on:
        # Caps Lock is on, but we want to deactivate it
        user32.keybd_event(VK_CAPITAL, 0, 0, 0)  # Press the Caps Lock key
        user32.keybd_event(VK_CAPITAL, 0, 2, 0)  # Release the Caps Lock key


set_caps_lock_state(False)
