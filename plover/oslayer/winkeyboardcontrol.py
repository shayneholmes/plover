#!/usr/bin/env python
# Copyright (c) 2011 Hesky Fisher.
# See LICENSE.txt for details.
#
# winkeyboardcontrol.py - capturing and injecting keyboard events in windows.

"""Keyboard capture and control in windows.

This module provides an interface for basic keyboard event capture and
emulation. Set the key_up and key_down functions of the
KeyboardCapture class to capture keyboard input. Call the send_string
and send_backspaces functions of the KeyboardEmulation class to
emulate keyboard input.

"""

import re
import functools
import ctypes
import pyHook
import pythoncom
import threading
import collections
import win32api
import win32con

# For the purposes of this class, we'll only report key presses that
# result in these outputs in order to exclude special key combos.
KEY_TO_ASCII = {
    41: '`', 2: '1', 3: '2', 4: '3', 5: '4', 6: '5', 7: '6', 8: '7',
    9: '8', 10: '9', 11: '0', 12: '-', 13: '=', 16: 'q',
    17: 'w', 18: 'e', 19: 'r', 20: 't', 21: 'y', 22: 'u', 23: 'i',
    24: 'o', 25: 'p', 26: '[', 27: ']', 43: '\\',
    30: 'a', 31: 's', 32: 'd', 33: 'f', 34: 'g', 35: 'h', 36: 'j',
    37: 'k', 38: 'l', 39: ';', 40: '\'', 44: 'z', 45: 'x',
    46: 'c', 47: 'v', 48: 'b', 49: 'n', 50: 'm', 51: ',',
    52: '.', 53: '/', 57: ' ',
}

# "Narrow python" unicode objects store chracters in UTF-16 so we
# can't iterate over characters in the standard way. This workaround
# let's us iterate over full characters in the string.


def characters(s):
    encoded = s.encode('utf-32-be')
    characters = []
    for i in xrange(len(encoded)/4):
        start = i * 4
        end = start + 4
        character = encoded[start:end].decode('utf-32-be')
        yield character

"""
SendInput code and classes based off of code
from user "inControl" on StackOverflow:

http://stackoverflow.com/questions/11906925/python-simulate-keydown
"""

LONG = ctypes.c_long
DWORD = ctypes.c_ulong
ULONG_PTR = ctypes.POINTER(DWORD)
WORD = ctypes.c_ushort
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_SCANCODE = 0x0008
KEYEVENTF_UNICODE = 0x0004
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1


class MOUSEINPUT(ctypes.Structure):
    _fields_ = (('dx', LONG),
                ('dy', LONG),
                ('mouseData', DWORD),
                ('dwFlags', DWORD),
                ('time', DWORD),
                ('dwExtraInfo', ULONG_PTR))


class KEYBDINPUT(ctypes.Structure):
    _fields_ = (('wVk', WORD),
                ('wScan', WORD),
                ('dwFlags', DWORD),
                ('time', DWORD),
                ('dwExtraInfo', ULONG_PTR))


class _INPUTunion(ctypes.Union):
    _fields_ = (('mi', MOUSEINPUT),
                ('ki', KEYBDINPUT))


class INPUT(ctypes.Structure):
    _fields_ = (('type', DWORD),
                ('union', _INPUTunion))


"""
NOTE: This is copied from the osxkeyboardcontrol.py and modified
for Windows using:

http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx
"""

# This mapping only works on keyboards using the ANSI standard layout. Each
# entry represents a sequence of keystrokes that are needed to achieve the
# given symbol. First, all keydown events are sent, in order, and then all
# keyup events are send in reverse order.
KEYNAME_TO_KEYCODE = collections.defaultdict(list, {
    # Adding media controls for windows
    'Vol_Mute': [0xAD], 'Vol_Down': [0xAE], 'Vol_Up': [0xAF],
    'Media_Next': [0xB0], 'Media_Prev': [0xB1], 'Media_Stop': [0xB2],
    'Media_Play_Pause': [0xB3], 'Sleep': [0x5F],

    # The order follows that of the plover guide, roughly
    '0': [0x30], '1': [0x31], '2': [0x32], '3': [0x33], '4': [0x34],
    '5': [0x35], '6': [0x36], '7': [0x37], '8': [0x38], '9': [0x39],

    'a': [0x41], 'b': [0x42], 'c': [0x43], 'd': [0x44], 'e': [0x45],
    'f': [0x46], 'g': [0x47], 'h': [0x48], 'i': [0x49], 'j': [0x4A],
    'k': [0x4B], 'l': [0x4C], 'm': [0x4D], 'n': [0x4E], 'o': [0x4F],
    'p': [0x50], 'q': [0x51], 'r': [0x52], 's': [0x53], 't': [0x54],
    'u': [0x55], 'v': [0x56], 'w': [0x57], 'x': [0x58], 'y': [0x59],
    'z': [0x5A],

    # Same as above, plus Shift_R for capital
    'A': [0xA1, 0x41], 'B': [0xA1, 0x42], 'C': [0xA1, 0x43], 'D': [0xA1, 0x44],
    'E': [0xA1, 0x45], 'F': [0xA1, 0x46], 'G': [0xA1, 0x47], 'H': [0xA1, 0x48],
    'I': [0xA1, 0x49], 'J': [0xA1, 0x4A], 'K': [0xA1, 0x4B], 'L': [0xA1, 0x4C],
    'M': [0xA1, 0x4D], 'N': [0xA1, 0x4E], 'O': [0xA1, 0x4F], 'P': [0xA1, 0x50],
    'Q': [0xA1, 0x51], 'R': [0xA1, 0x52], 'S': [0xA1, 0x53], 'T': [0xA1, 0x54],
    'U': [0xA1, 0x55], 'V': [0xA1, 0x56], 'W': [0xA1, 0x57], 'X': [0xA1, 0x58],
    'Y': [0xA1, 0x59], 'Z': [0xA1, 0x5A],

    'Alt_L': [0x12], 'Alt_R': [0x12], 'Control_L': [0xA2], 'Control_R': [0xA3],
    'Hyper_L': [], 'Hyper_R': [], 'Meta_L': [], 'Meta_R': [],
    'Shift_L': [0xA0], 'Shift_R': [0xA1], 'Super_L': [0x5B], 'Super_R': [0x5C],

    'Caps_Lock': [0x14], 'Num_Lock': [0x90], 'Scroll_Lock': [0x91],
    'Shift_Lock': [],

    'Return': [0x0D], 'Tab': [0x09], 'BackSpace': [0x08], 'Delete': [0x2E],
    'Escape': [0x1B], 'Break': [0x03], 'Insert': [0x2D], 'Pause': [0x13],
    'Print': [0x2C], 'Sys_Req': [],

    'Up': [0x26], 'Down': [0x28], 'Left': [0x25], 'Right': [0x27],
    'Page_Up': [0x21], 'Page_Down': [0x22], 'Home': [0x24], 'End': [0x23],

    'F1': [0x70], 'F2': [0x71], 'F3': [0x72], 'F4': [0x73], 'F5': [0x74],
    'F6': [0x75], 'F7': [0x76], 'F8': [0x77], 'F9': [0x78], 'F10': [0x79],
    'F11': [0x7A], 'F12': [0x7B], 'F13': [0x7C], 'F14': [0x7D], 'F15': [0x7E],
    'F16': [0x7F], 'F17': [0x80], 'F18': [0x81], 'F19': [0x82], 'F20': [0x83],
    'F21': [0x84], 'F22': [0x85], 'F23': [0x86], 'F24': [0x87], 'F25': [],
    'F26': [], 'F27': [], 'F28': [], 'F29': [], 'F30': [], 'F31': [],
    'F32': [], 'F33': [], 'F34': [], 'F35': [],

    'L1': [], 'L2': [], 'L3': [], 'L4': [], 'L5': [], 'L6': [],
    'L7': [], 'L8': [], 'L9': [], 'L10': [],

    'R1': [], 'R2': [], 'R3': [], 'R4': [], 'R5': [], 'R6': [],
    'R7': [], 'R8': [], 'R9': [], 'R10': [], 'R11': [], 'R12': [],
    'R13': [], 'R14': [], 'R15': [],

    'KP_0': [0x60], 'KP_1': [0x61], 'KP_2': [0x62], 'KP_3': [0x63],
    'KP_4': [0x64], 'KP_5': [0x65], 'KP_6': [0x66], 'KP_7': [0x67],
    'KP_8': [0x68], 'KP_9': [0x69], 'KP_Add': [0xA1, 0xBB], 'KP_Begin': [],
    'KP_Decimal': [0x6E], 'KP_Delete': [0x2E], 'KP_Divide': [0x6F],
    'KP_Down': [], 'KP_End': [], 'KP_Enter': [0x0D], 'KP_Equal': [0xBB],
    'KP_F1': [], 'KP_F2': [], 'KP_F3': [], 'KP_F4': [], 'KP_Home': [],
    'KP_Insert': [], 'KP_Left': [], 'KP_Multiply': [0x6A], 'KP_Next': [],
    'KP_Page_Down': [], 'KP_Page_Up': [], 'KP_Prior': [], 'KP_Right': [],
    'KP_Separator': [], 'KP_Space': [], 'KP_Subtract': [0x6D], 'KP_Tab': [],
    'KP_Up': [],

    'ampersand': [0xA1, 0x37], 'apostrophe': [0xDE],
    'asciitilde': [0xA1, 0xC0], 'asterisk': [0xA1, 0x38], 'at': [0xA1, 0x32],
    'backslash': [0xDC], 'braceleft': [0xA1, 0xDB], 'braceright': [0xA1, 0xDD],
    'bracketleft': [0xDB], 'bracketright': [0xDD], 'colon': [0xA1, 0xBA],
    'comma': [0xBC], 'division': [], 'dollar': [0xA1, 0x34], 'equal': [0xBB],
    'exclam': [0xA1, 0x31], 'greater': [0xA1, 0xBE], 'hyphen': [0xBD],
    'less': [0xA1, 0xBC], 'minus': [0xBD], 'multiply': [0x6A],
    'numbersign': [0xA1, 0x33], 'parenleft': [0xA1, 0x39],
    'parenright': [0xA1, 0x30], 'percent': [0xA1, 0x35], 'period': [0xBE],
    'plus': [0xA1, 0xBB], 'question': [0xA1, 0xBF], 'quotedbl': [0xA1, 0xDE],
    'quoteleft': [], 'quoteright': [], 'semicolon': [0xBA], 'slash': [0xBF],
    'space': [0x20], 'underscore': [0xA1, 0xBD], 
    'grave': [0xC0], 'asciicircum': [0xA1, 0x36], 'bar': [0xA1, 0xDC],

    'Help': [0x2F], 'Mode_switch': [0x1F], 'Menu': [0x5D],


    'Begin': [], 'Cancel': [0x03], 'Clear': [0x0C], 'Execute': [0x2B],
    'Find': [], 'Linefeed': [],
    'Multi_key': [], 'MultipleCandidate': [], 'Next': [0x22],
    'PreviousCandidate': [], 'Prior': [0x21], 'Redo': [], 'Select': [0x29],
    'SingleCandidate': [], 'Undo': [],

    'Eisu_Shift': [], 'Eisu_toggle': [], 'Hankaku': [], 'Henkan': [],
    'Henkan_Mode': [], 'Hiragana': [], 'Hiragana_Katakana': [],
    'Kana_Lock': [], 'Kana_Shift': [], 'Kanji': [], 'Katakana': [],
    'Mae_Koho': [], 'Massyo': [], 'Muhenkan': [], 'Romaji': [],
    'Touroku': [], 'Zen_Koho': [], 'Zenkaku': [], 'Zenkaku_Hankaku': []
})

KEYNAME_TO_UNICODE = {
    # Decimal values of Unicode chars
    'plusminus': 177, 'aring': 229, 'yen': 165, 'ograve': 242,
    'adiaeresis': 228, 'Ntilde': 209, 'questiondown': 191, 'Yacute': 221,
    'Atilde': 195, 'ccedilla': 231, 'copyright': 169, 'ntilde': 241,
    'otilde': 245, 'masculine': 9794, 'Eacute': 201, 'ocircumflex': 244,
    'guillemotright': 187, 'ecircumflex': 234, 'uacute': 250, 'cedilla': 184,
    'oslash': 248, 'acute': 237, 'ssharp': 223, 'Igrave': 204,
    'twosuperior': 178, 'udiaeresis': 252, 'notsign': 172, 'exclamdown': 161,
    'ordfeminine': 9792, 'Otilde': 213, 'agrave': 224, 'ection': 167,
    'egrave': 232, 'macron': 175, 'Icircumflex': 206, 'diaeresis': 168,
    'ucircumflex': 251, 'atilde': 227, 'Acircumflex': 194, 'degree': 176,
    'THORN': 222, 'acircumflex': 226, 'Aring': 197, 'Ooblique': 216,
    'Ugrave': 217, 'Agrave': 192, 'ydiaeresis': 255, 'threesuperior': 179,
    'Egrave': 200, 'Idiaeresis': 207, 'igrave': 236, 'ETH': 208,
    'Ecircumflex': 202, 'Aacute': 193, 'cent': 162, 'registered': 174,
    'Oacute': 211, 'Adiaeresis': 228, 'guillemotleft': 171, 'ediaeresis': 235,
    'Ograve': 210, 'mu': 956, 'paragraph': 182, 'Ccedilla': 199, 'thorn': 254,
    'threequarters': 190, 'ae': 230, 'brokenbar': 166, 'nobreakspace': 32,
    'currency': 164, 'ugrave': 249, 'Ucircumflex': 219, 'odiaeresis': 246,
    'periodcentered': 183, 'Uacute': 218, 'idiaeresis': 239, 'yacute': 253,
    'sterling': 163, 'AE': 198, 'Ediaeresis': 203, 'onequarter': 188,
    'onehalf': 189, 'Thorn': 222, 'aacute': 225, 'icircumflex': 238,
    'Udiaeresis': 220, 'eacute': 233, 'Eth': 240, 'eth': 240, 'Iacute': 205,
    'onesuperior': 185, 'Ocircumflex': 212, 'Odiaeresis': 214, 'oacute': 243
}

# Maps from literal characters to their key names.
LITERALS = collections.defaultdict(str, {
    '`': 'grave', '~': 'asciitilde', '!': 'exclam', '@': 'at',
    '#': 'numbersign', '$': 'dollar', '%': 'percent', '^': 'asciicircum',
    '&': 'ampersand', '*': 'asterisk', '(': 'parenleft', ')': 'parenright',
    '-': 'minus', '_': 'underscore', '=': 'equal', '+': 'plus',
    '[': 'bracketleft', ']': 'bracketright', '{': 'braceleft',
    '}': 'braceright', '\\': 'backslash', '|': 'bar', ';': 'semicolon',
    ':': 'colon', '\'': 'apostrophe', '"': 'quotedbl', ',': 'comma',
    '<': 'less', '.': 'period', '>': 'greater', '/': 'slash',
    '?': 'question', '\t': 'Tab', ' ': 'space'
})


class KeyboardCapture(threading.Thread):
    """Listen to all keyboard events."""

    CONTROL_KEYS = set(('Lcontrol', 'Rcontrol'))
    SHIFT_KEYS = set(('Lshift', 'Rshift'))
    ALT_KEYS = set(('Lmenu', 'Rmenu'))

    def __init__(self):
        threading.Thread.__init__(self)

        self.suppress_keyboard(True)

        self.shift = False
        self.ctrl = False
        self.alt = False

        # NOTE(hesky): Does this need to be more efficient and less
        # general if it will be called for every keystroke?
        def on_key_event(func_name, event):
            ascii = KEY_TO_ASCII.get(event.ScanCode, None)
            if not event.Injected:
                if event.Key in self.CONTROL_KEYS:
                    self.ctrl = func_name == 'key_down'
                if event.Key in self.SHIFT_KEYS:
                    self.shift = func_name == 'key_down'
                if event.Key in self.ALT_KEYS:
                    self.alt = func_name == 'key_down'
                if ascii and not self.ctrl and not self.alt and not self.shift:
                    getattr(self, func_name, lambda x: True)(
                        KeyboardEvent(ascii))
                    return not self.is_keyboard_suppressed()

            return True

        self.hm = pyHook.HookManager()
        self.hm.KeyDown = functools.partial(on_key_event, 'key_down')
        self.hm.KeyUp = functools.partial(on_key_event, 'key_up')

    def run(self):
        self.hm.HookKeyboard()
        pythoncom.PumpMessages()

    def cancel(self):
        if self.is_alive():
            self.hm.UnhookKeyboard()
            win32api.PostThreadMessage(self.ident, win32con.WM_QUIT)

    def can_suppress_keyboard(self):
        return True

    def suppress_keyboard(self, suppress):
        self._suppress_keyboard = suppress

    def is_keyboard_suppressed(self):
        return self._suppress_keyboard


class KeyboardEmulation:

    # Sends input types to buffer
    def _SendInput(self, *inputs):
        nInputs = len(inputs)
        LPINPUT = INPUT * nInputs
        pInputs = LPINPUT(*inputs)
        cbSize = ctypes.c_int(ctypes.sizeof(INPUT))
        return ctypes.windll.user32.SendInput(nInputs, pInputs, cbSize)

    # Input type (can be mouse, keyboard)
    def _Input(self, structure):
        if isinstance(structure, MOUSEINPUT):
            return INPUT(INPUT_MOUSE, _INPUTunion(mi=structure))
        if isinstance(structure, KEYBDINPUT):
            return INPUT(INPUT_KEYBOARD, _INPUTunion(ki=structure))
        raise TypeError('Cannot create INPUT structure!')

    # Container to send mouse input
    # Not used, but maybe one day it will be useful
    def _MouseInput(flags, x, y, data):
        return MOUSEINPUT(x, y, data, flags, 0, None)

    # KEYBoarD input type to send key input
    def _KeybdInput(self, code, flags):
        if flags == KEYEVENTF_UNICODE:
            # special handling of Unicode characters
            return KEYBDINPUT(0, code, flags, 0, None)
        return KEYBDINPUT(code, code, flags, 0, None)

    # Abstraction to set flags to 0 and create an input type
    def _Keyboard(self, code, flags=0):
        return self._Input(self._KeybdInput(code, flags))

    # Presses a key down
    def _key_down(self, keyname):
        # Press all keys
        for keycode in KEYNAME_TO_KEYCODE[keyname]:
            self._SendInput(self._Keyboard(keycode))

    # Releases a key
    def _key_up(self, keyname):
        # Release all keys backwards
        for keycode in reversed(KEYNAME_TO_KEYCODE[keyname]):
            self._SendInput(self._Keyboard(keycode, KEYEVENTF_KEYUP))

    # Press and release a key
    def _key_press(self, keyname):
        self._key_down(keyname)
        self._key_up(keyname)

    # Send a Unicode character to application
    def _key_unicode(self, code):
        self._SendInput(self._Keyboard(code, KEYEVENTF_UNICODE))

    def send_backspaces(self, number_of_backspaces):
        for _ in xrange(number_of_backspaces):
            self._key_press("BackSpace")

    def send_string(self, s):
        for c in characters(s):

            # We normalize characters
            # Like . to period, * to asterisk
            if c in LITERALS:
                c = LITERALS[c]

            # We check if we know the character
            # If we do we can do a manual keycode
            if c in KEYNAME_TO_KEYCODE:
                self._key_press(c)

            # Otherwise, we send it as a Unicode character
            else:
                self._key_unicode(ord(c))

    def send_key_combination(self, combo_string):
        """Emulate a sequence of key combinations.

        Argument:

        combo_string -- A string representing a sequence of key
        combinations. Keys are represented by their names in the
        KEYNAME_TO_KEYCODE above. For example, the
        left Alt key is represented by 'Alt_L'. Keys are either
        separated by a space or a left or right parenthesis.
        Parentheses must be properly formed in pairs and may be
        nested. A key immediately followed by a parenthetical
        indicates that the key is pressed down while all keys enclosed
        in the parenthetical are pressed and released in turn. For
        example, Alt_L(Tab) means to hold the left Alt key down, press
        and release the Tab key, and then release the left Alt key.

        """

        # We will go through and press down keys
        # When encountering (, we will add to the held stack
        # ) will release these from the stack

        # There is a problem. If the user defines something like:
        #   Shift_L(ampersand x)
        # Shift_L( will press the shift key, but ampersand will release it
        # too early. x output will be lowercase.
        # In order to combat this, ampersand and other shifted characters
        # use Shift_R as most entries seem to use Shift_L.
        # That being said, we can consider a shifted-shifted character to be
        # undefined behavior...

        keycode_events = []
        key_down_stack = []
        current_command = []
        for c in combo_string:
            if c in (' ', '(', ')'):
                # Keystring is the variable (l, b, Alt_L, etc.)
                keystring = ''.join(current_command)
                # Clear out current command
                current_command = []

                # Handle unicode characters by pressing them
                if keystring in KEYNAME_TO_UNICODE:
                    self._key_unicode(KEYNAME_TO_UNICODE[keystring])
                    # Reset keystring to nothing to prevent further presses
                    keystring = ''

                if c == ' ':
                    # Record press and release for command's keys.
                    if keystring:
                        self._key_press(keystring)
                elif c == '(':
                    # Record press for command's key.
                    if keystring:
                        self._key_down(keystring)
                    key_down_stack.append(keystring)
                elif c == ')':
                    # Record press and release for command's key and
                    # release previously held keys.
                    if keystring:
                        self._key_press(keystring)
                    if key_down_stack:
                        self._key_up(key_down_stack.pop())
            else:
                current_command.append(c)

        # Record final command key.
        keystring = ''.join(current_command)
        if keystring in KEYNAME_TO_UNICODE:
                    self._key_unicode(KEYNAME_TO_UNICODE[keystring])
                    # Reset keystring to nothing to prevent further presses
                    keystring = ''
        else:
            self._key_press(keystring)
        # Release all keys.
        # Should this be legal in the dict (lack of closing parens)?
        for keystring in key_down_stack:
            if keystring:
                self._key_up(keystring)


class KeyboardEvent(object):
    """A keyboard event."""

    def __init__(self, char):
        self.keystring = char

if __name__ == '__main__':
    kc = KeyboardCapture()
    ke = KeyboardEmulation()

    def test(event):
        print event.keystring
        ke.send_backspaces(1)
        ke.send_string(' you pressed: "' + event.keystring + '" ')

    kc.key_up = test
    kc.start()
    print 'Press CTRL-c to quit.'
    try:
        while True:
            pass
    except KeyboardInterrupt:
        kc.cancel()
