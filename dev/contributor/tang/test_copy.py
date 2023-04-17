import win32clipboard
from ctypes import *


class DROPFILES(Structure):
    _fields_ = [
        ("pFiles", c_uint),
        ("x", c_long),
        ("y", c_long),
        ("fNC", c_int),
        ("fWide", c_bool),
    ]


def setClipboardFiles(paths):
    files = ("\0".join(paths)).replace("/", "\\")
    data = files.encode("U16")[2:] + b"\0\0"
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(
            win32clipboard.CF_HDROP, matedata + data)
    finally:
        win32clipboard.CloseClipboard()


def readClipboardFilePaths():
    win32clipboard.OpenClipboard()
    try:
        return win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
    finally:
        win32clipboard.CloseClipboard()


pDropFiles = DROPFILES()
pDropFiles.pFiles = sizeof(DROPFILES)
pDropFiles.fWide = True
matedata = bytes(pDropFiles)
filename = [r"d:\a.pptx"]
setClipboardFiles(filename)

# 作者：18037128621
# 链接：https: // juejin.cn / post / 7123461124162322445
# 来源：稀土掘金
# 著作权归作者所有。商业转载请联系作者获得授权，非商业转载请注明出处。