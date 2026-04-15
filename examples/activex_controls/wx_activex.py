"""
PyXLL Examples: Wx ActiveX control

This example shows how a WxPython window can be hosted
in an Excel ActiveX control.

ActiveX controls are embedded directly into the Excel workbook.
These can be used for creating advanced UI Python tools within Excel.
"""
from pyxll import create_activex_control, xl_app
import logging
import wx
import os

_log = logging.getLogger(__name__)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

class WxActiveXPanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        # Main sizer for everything in the window
        sizer = wx.BoxSizer(orient=wx.VERTICAL)

        # Load the Tk icon and add the header to the widget
        wx.InitAllImageHandlers()
        image_path = os.path.join(os.path.dirname(__file__), "..", "ribbon", "icons", "wx.png")
        icon = wx.Icon(image_path, type=wx.BITMAP_TYPE_PNG)
        static_bitmap = wx.StaticBitmap(self)
        static_bitmap.SetIcon(icon)
        label = wx.StaticText(self, label="Wx ActiveX Control")

        header = wx.BoxSizer(orient=wx.HORIZONTAL)
        header.Add(static_bitmap, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=10)
        header.Add(label, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=10)
        sizer.Add(header)

        # Step 1, getting the address of the currently selected cell
        label = wx.StaticText(self, label="1. Press button to get the current selection")
        sizer.Add(label, flag=wx.LEFT | wx.RIGHT | wx.ALIGN_LEFT, border=10)

        button = wx.Button(self, label=">>", size=(40, 30))
        button.Bind(wx.EVT_BUTTON, self.get_selected_cell_address)
        self.address = wx.TextCtrl(self)
        address_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)
        address_sizer.Add(button, flag=wx.ALIGN_CENTER_VERTICAL)
        address_sizer.Add(self.address, proportion=1, flag=wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(address_sizer, flag=wx.ALL | wx.EXPAND, border=10)
        sizer.AddSpacer(20)

        # Step 2, getting the value of the cell
        label = wx.StaticText(self, label="2. Press the button to get cell value")
        sizer.Add(label, flag=wx.LEFT | wx.RIGHT | wx.ALIGN_LEFT, border=5)

        button = wx.Button(self, label="Get Cell Value")
        button.Bind(wx.EVT_BUTTON, self.get_cell_value)
        sizer.Add(button, flag=wx.ALL | wx.EXPAND, border=10)
        sizer.AddSpacer(20)

        self.value = wx.TextCtrl(self)
        sizer.Add(self.value, flag=wx.ALL | wx.EXPAND, border=10)
        sizer.AddSpacer(20)

        # Step 3, setting the value of the cell
        label = wx.StaticText(self, label="3. Edit the text above and set it")
        sizer.Add(label, flag=wx.LEFT | wx.RIGHT | wx.ALIGN_LEFT, border=10)

        button = wx.Button(self, label="Set Cell Value")
        button.Bind(wx.EVT_BUTTON, self.set_cell_value)
        sizer.Add(button, flag=wx.ALL | wx.EXPAND, border=10)

        # Finish laying out the panel
        sizer.AddStretchSpacer()
        self.SetSizer(sizer)
        self.Layout()

    def get_selected_cell_address(self, event):
        """Get the address of the currently selected cell."""
        xl = xl_app(com_package="win32com")
        address = xl.Selection.Address
        self.address.SetValue(address)

    def get_cell_value(self, event):
        address = self.address.GetValue()
        if not address:
            wx.MessageBox("Select a cell with the button above first.", style=wx.ICON_ERROR)

        xl = xl_app(com_package="win32com")
        cell = xl.Range(address)
        self.value.SetValue(str(cell.Value))

    def set_cell_value(self, event):
        address = self.address.GetValue()
        if not address:
            wx.MessageBox("Error", "Select a cell with the button above first.", style=wx.ICON_ERROR)

        xl = xl_app(com_package="win32com")
        cell = xl.Range(address)
        cell.Value = self.value.GetValue()

class WxActiveXFrame(wx.Frame):

    def __init__(self, name):
        wx.Frame.__init__(self, parent=None)
        self.SetTitle(name)
        self.control = WxActiveXPanel(parent=self)

def show_wx_activex():
    """Create a Wx window and embed it in Excel as a Custom Task Pane."""

    # Make sure the wx App has been created
    app = wx.App.Get()
    if app is None:
        app = wx.App()

    # Create the frame to use as the Custom Task Pane
    frame = WxActiveXFrame("Wx Example")

    # Use PyXLL's 'create_activex_control' function to create the ActiveX control.
    # By using a named control, if this is called again with the same name the
    # existing control will be updated with the new widget.
    create_activex_control(frame,
                           name="wx_activex_example",
                           top=50,
                           left=600,
                           width=250,
                           height=300)

if __name__ == "__main__":
    app = wx.App()
    frame = WxActiveXFrame("Wx Example")
    frame.Show()
    app.MainLoop()