"""
PyXLL Examples: ActiveX Controls

These examples show how Excel ActiveX controls can be created
by PyXLL to host widgets created using Python's main UI toolkits.

The supported Python UI Toolkits are:

- Tkinter
- Qt (PySide6, PyQt6, PySide2 and PyQt5)
- WxWindows

Tkinter is usually installed as a standard library with Python
but the other UI toolkits will need to be installed using pip
or conda for the examples to work.

ActiveX controls are embedded directly into the Excel workbook.
These can be used for creating advanced UI Python tools within Excel.
"""
from pyxll import xl_menu

def _check_pywin32():
    try:
        import win32com
    except ImportError:
        raise RuntimeError("""

pywin32 needs to be installed for this example.
Install either using pip or conda and try again.

""")

def show_tk_activex():
    _check_pywin32()
    try:
        from . import tk_activex
        tk_activex.show_tk_activex()
    except ImportError:
        raise RuntimeError("""

Tkinter needs to be installed for this example.
Install either using pip or conda and try again.

""")

def show_qt_activex():
    _check_pywin32()
    try:
        from . import qt_activex
        qt_activex.show_qt_activex()
    except ImportError:
        raise RuntimeError("""

Either PyQt or PySide needs to be installed for this example.
Install either using pip or conda and try again.

""")

def show_wx_activex():
    _check_pywin32()
    try:
        from . import wx_activex
        wx_activex.show_wx_activex()
    except ImportError:
        raise RuntimeError("""

wxPython to be installed for this example.
Install either using pip or conda and try again.

""")

#
# Ribbon actions for the example ribbon toolbar
#

def tk_activex_ribbon_action(control):
    show_tk_activex()

def qt_activex_ribbon_action(control):
    show_qt_activex()

def wx_activex_ribbon_action(control):
    show_wx_activex()

#
# Menu functions for the 'Add-ins' tab
#

@xl_menu("Tk", sub_menu="ActiveX Controls")
def tk_activex_menu():
    show_tk_activex()

@xl_menu("Qt", sub_menu="ActiveX Controls")
def qt_activex_menu():
    show_qt_activex()

@xl_menu("Wx", sub_menu="ActiveX Controls")
def wx_activex_menu():
    show_wx_activex()