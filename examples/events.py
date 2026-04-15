"""
PyXLL Examples: Excel Events

Using the PyXLL @xl_event decorators we can register Python
functions to be called when certain Excel COM events* are fired.

Each event handler function takes the same arguments as the
corresponding VBA event handler, as documented in the Excel
online documentation.

Excel objects passed to the event handlers are passed as
COM objects, and a third party COM wrapper package such
as pywin32 or comtypes is required for this feature.

The pywin32 package can be installed by running::

    pip install pywin32

* Excel COM support was added in Office 2000. If you are
  using an earlier version these COM examples won't work.
"""

from pyxll import xl_event
import logging

_log = logging.getLogger(__name__)

@xl_event.sheet_activate
def on_sheet_activated(sheet):
    """This handles the Excel Workbook.SheetActivate event
    and is passed a COM Sheet object."""

    # The sheet is an Excel Sheet object as has the same properties as the corresponding VBA Sheet object.
    sheet_name = sheet.Name

    # Log a message to show the event has been handled.
    # This module should be removed from pyxll.cfg file once you are finished with this example.
    _log.info("EXAMPLE: Workbook.SheetActivate event for '%s' handled by %s" % (sheet_name, __file__))