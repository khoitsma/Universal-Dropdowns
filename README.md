# Universal-Dropdowns
### Universal Dropdowns Using **PyXLL**

What does that mean?

The idea is so simple: what if a user wanted to access standard dropdown choices on all his workbooks, with pure workbooks (no macros, no validation lists, no piddlin’ setups).

References

- [Universal Dropdowns in Excel Dec 2024](https://www.mathpax.com/universal-dropdowns-in-excel-dec-2024/)

- [Excel Dropdown with NO worksheet data list (with video) Nov 2020
](https://www.mathpax.com/excel-dropdown-with-no-worksheet-data-list/)

---

20 October 2025

### And now I offer an updated, more compact version of the universal dropdown.

My previous version was based on a PyXLL macro; the new version is based on a PyXLL function (set as macro=True).

Link to the new function: (https://github.com/khoitsma/Universal-Dropdowns/blob/main/)select_from_pyxll_module.py)

```
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
```

When called, the function presents a PySide6 QDialog dropdown incorporating the following features:
* short and simple
* dropdown lists can be set in Python as a global variable (executed once, no recalculation)
* avoids the restrictions of Excel built-in dropdowns
* dropdown is available in any workbook, any sheet
* regex filtering
* dropdown index is set to current value
* custom None value is returned to avoid formula overwrite
* calling cell address or name (if named) is built into the dialog title
* a default list is coded
* no macro buttons, clean interface
* thanks to [Tony Roberts](https://www.pyxll.com/blog/author/tony/) for assistance with the `PyXll schedule_call`

***Video/Screenshot***

> - [Excel video snippet](https://khoitsmahq.firstcloudit.com/images/universal_dropdown_select_from_PyXLL_dropdown_function_TR.mp4)
> 
> - ![Excel screenshot](https://khoitsmahq.firstcloudit.com/images/universal_dropdown_select_from_PyXLL_dropdown_function_TR.png)
