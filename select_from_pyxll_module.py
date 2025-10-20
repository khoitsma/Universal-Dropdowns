import sys
import os
import re

from pyxll import (
    xl_app,
    xl_func,
    xlfCaller,
    schedule_call,
)

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QSize
from PySide6.QtGui import QScreen, QFont
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QInputDialog,
    QPushButton,
)


class Form0(QDialog):
    def __init__(self, parent=None):
        super(Form0, self).__init__(parent)
        self.setWindowTitle("Form0")

        # Create a custom font
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(24)

        # invoke the font
        self.setFont(font)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)  # Connect to reject()

        # A QVBoxLayout is a layout manager in the Qt framework that arranges widgets vertically,
        # in a top-to-bottom sequence.
        # As more widgets are added, they are placed one after the other in a column.
        #
        # Key characteristics:
        #
        # Vertical arrangement: It lines up widgets one above the other.
        # Sequential addition: New widgets are placed at the bottom
        #    of the column as they are added.
        # Dynamic resizing: The layout and its contained widgets
        #    resize automatically when the parent window is resized.
        # Flexible spacing: You can manage the empty space between
        #    widgets using spacers (addStretch()) or by setting
        #    the spacing with setSpacing().
        # Layout hierarchy: A QVBoxLayout must be set on a parent widget
        #    using setLayout() to be displayed.
        layout = QVBoxLayout()
        layout.addWidget(cancel_button)

        # Set the dialog layout
        self.setLayout(layout)


# for my demo function
# build a demo list of lists
#
# the lists are
#    apples
#    countries
#    three letter names
#    cars

my_dropdown_lists = {
    "apples": [
        "gala",
        "honeycrisp",
        "macintosh",
        "cosmic crisp",
        "fuji",
        "granny smith",
        "pink lady",
    ],
    "countries": [
        "Uganda",
        "Ukraine",
        "United Arab Emirates",
        "United Kingdom",
        "United States of America",
        "Uruguay",
        "Uzbekistan",
    ],
    "three letter names": ["Bob", "Joe", "Leo", "Ned", "Ray", "Sam", "Tom"],
    "cars": [
        "alpha romeo",
        "buick",
        "chrysler",
        "dodge",
        "elantra",
        "fiat",
        "gremlin",
        "honda",
        "isuzu",
        "jaguar",
        "kia",
        "lincoln",
        "mazda",
        "nissan",
        "opel",
        "pontiac",
        "quattro",
        "rolls royce",
        "saturn",
        "toyota",
        "uplander",
        "volkswagen",
        "wrx",
        "xantia",
        "yaris",
        "zagato",
    ],
}


@xl_func(
    "str db_name, str r_filter_value: var",
    none_value="  » » »",
    volatile=False,
    macro=True,
)
def select_from_PyXLL_dropdown_function_TR(dropdown_name=None, filter_string=""):
    """[PyXLL] Execute a PySide6 dialog based on a (pre-coded) list of lists

    :param dropdown_name:  [apples,countries,three letter names,cars] Defaults to None
    :param r_filter_value: a filter pattern (regex allowed). Defaults to ""
    """

    def update_func():
        xl = xl_app()

        # get the cell address from which the function is called
        r_entry = xlfCaller().to_range()

        # for the demo, hardwire this
        # the answer will be placed
        # in the cell to the right of the caller call
        r_entry_answer = r_entry.GetOffset(RowOffset=0, ColumnOffset=1)

        r_dropdown = dropdown_name

        # set the default for picked_list
        picked_list = my_dropdown_lists["three letter names"]

        if r_dropdown:
            picked_list = my_dropdown_lists[r_dropdown.lower()]

        # process the filter parameter
        if filter_string:
            try:
                regex = re.compile(filter_string.upper())
                filtered_list = filter(lambda x: regex.search(x.upper()), picked_list)
                picked_list = list(filtered_list)
            except:
                pass

        # create the custom form
        dlg = Form0()

        # set current index value for the dropdown dialog
        try:
            # use existing r_entry_answer value if possible
            dropdown_index = picked_list.index(r_entry_answer.Value)
        except:
            dropdown_index = 0

        try:
            dialog_title = "".join(["[", r_entry.Name.Name, "]", " PyQt Dropdown"])
        except:
            dialog_title = "".join(["[", r_entry.Address, "]", " PyQt Dropdown"])

        # get text answer from the dialog
        text = QInputDialog.getItem(
            dlg,
            dialog_title,
            "Selection: ",
            picked_list,
            current=max(0, min(dropdown_index, len(picked_list) - 1)),
        )

        # set default for picked item
        picked_item = r_entry_answer.Value

        # text is a tuple, our desired answer is text[0]
        # text[1] is zero if no selection is made, or if dialog is cancelled
        # print(text)
        if text[1] != 0:
            picked_item = text[0]

        # post the user-selected results
        r_entry_answer.Value = picked_item

    # Schedule calling the update function
    schedule_call(update_func)
