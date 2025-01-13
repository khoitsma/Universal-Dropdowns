@xl_func("str reg_pattern, str reg_text: bool")
def reg_klh_tf(reg_pattern, reg_text):
    return True if re.search(reg_pattern, reg_text) else False


@xl_func(
    "string pathname_inp, string filename_inp, string dataname_inp, string[] rang_inp: string"
)
def build_dropdown_dict(pathname_inp, filename_inp, dataname_inp, rang_inp):
    """build a JSON file from an Excel list,
    the JSON file is used later by my PyXLL function
    to populate my Pyside QT dropdown"""
    result = "FAIL"
    try:
        filename = pathname_inp + filename_inp + ".json"
        # data_name = cell_inp
        data_list = [e for e in rang_inp]
        with open(filename, "w") as f:
            f.write("".join(['{"', dataname_inp, '": [\n']))
            data_list_len = len(data_list)
            # all but last item
            for e in data_list[:-1]:
                # with comma
                f.write("".join(['"', e, '",\n']))
            # last item, therefore no comma
            f.write("".join(['"', data_list[-1], '"\n']))
            f.write("]}\n")
        result = "SUCCESS"
    except:
        pass
    return result


@xl_func("string pathname_inp, string filename_inp: string")
def append_to_USER_PyXLL_dropdown(pathname_inp, filename_inp):
    """
    Append a new dropdown to the existing user dictionary of dropdowns
    """
    result = "FAIL"
    try:

        cfg = pyxll.get_config()
        path_fn = "E:/my_pyxll_modules/user_dict.json"
        temp_tuple = "KARL", "user_dict_path"
        if cfg.has_option(*temp_tuple):
            path_fn = cfg.get(*temp_tuple)

        # reading the data from the existing file
        with open("E:/my_pyxll_modules/user_dict.json") as f:
            data = f.read()

        # reconstruct the data as a dictionary
        js_exist_dict = js.loads(data)
        print("Data type js_exist_dict type: ", type(js_exist_dict))
        print("Data type js_exist_dict: ", js_exist_dict)

        # reading the data from the new file (to be appended)
        filename = os.path.join(pathname_inp, filename_inp + ".json")
        with open(filename) as f:
            data = f.read()

        # reconstruct the data as a dictionary
        js_new_dict = js.loads(data)
        print("Data type js_new_dict type: ", type(js_new_dict))
        print("Data type js_new_dict: ", js_new_dict)

        # see if first_key in new dict is found in exist dict
        first_key = list(js_new_dict)[0]
        if list(js_new_dict)[0] in js_exist_dict:
            xl = pyxll.xl_app()
            title = "* IMPORTANT *"
            prompt = (
                "Are you sure you want to replace ["
                + filename_inp
                + "] ? (True/False or 1/0)"
            )  # set prompt
            mtype = 4  # boolean
            response = xl.Application.InputBox(Prompt=prompt, Title=title, Type=mtype)
            if response:
                # user says okay; therefore proceed
                # first delete
                del js_exist_dict[first_key]
                # then replace
                js_exist_dict.update(js_new_dict)

                # Convert and write JSON object to file
                with open(path_fn, "w") as outfile:
                    js.dump(js_exist_dict, outfile)
                result = "SUCCESS"
            else:
                pass
        else:
            # didn't find existing; therefore proceed
            js_exist_dict.update(js_new_dict)

            # Convert and write JSON object to file
            with open(path_fn, "w") as outfile:
                js.dump(js_exist_dict, outfile)
            result = "SUCCESS"

    except:
        pass
    return result


def index_containing_substring(the_list, substring):
    for i, s in enumerate(the_list):
        if substring in s.upper():
            return i
    return -1


@xl_macro()
def select_from_USER_PyXLL_dropdown():
    """
    Execute a PyQt dialog based on a list from a dictionary
    """
    xl = pyxll.xl_app()

    # cfg = pyxll.get_config()
    # path_fn = "E:/my_pyxll_modules/user_dict.json"
    # temp_tuple = "KARL", "user_dict_path"
    # if cfg.has_option(*temp_tuple):
    #     path_fn = cfg.get(*temp_tuple)

    # # print(temp_tuple)
    # # print(path_fn)
    # # reading the data from the file
    # with open(path_fn) as f:
    #     data = f.read()

    # # print(data)
    # # print("Data type before reconstruction: ", type(data))

    # # reconstructing the data as a dictionary
    # js_main_dict = js.loads(data)
    # # print("Data type js_main_dict type: ", type(js_main_dict))
    # # print("Data type js_main_dict: ", js_main_dict)

    # js_main_dict is a GLOBAL variable now

    r_entry = xl.Selection
    # next row up
    r_init_last_range = r_entry.GetOffset(RowOffset=-1, ColumnOffset=0)
    # next row up
    r_filter_value = r_entry.GetOffset(RowOffset=-2, ColumnOffset=0).Value
    # next row up
    r_named_variable = r_entry.GetOffset(RowOffset=-3, ColumnOffset=0).Value
    # default situation is use the default_dict named in the pyxll.cfg
    if r_named_variable is None:
        r_named_variable = default_dict

    r_init_last_text = None
    rilr = r_init_last_range.Value
    if rilr:
        if isinstance(rilr, (int, float)):
            rilr = str(rilr)
        r_init_last_text = rilr.upper()

    # set default for picked_list
    # first remove whitespace
    r_named_variable_clean = r_named_variable.strip()

    # now determine if the json reference
    # is a list or a 'table' (list of lists)

    splits = r_named_variable_clean.split("#")
    # if the split produces a list then the
    # r_named_variable is a 'table'
    # splits[0] will be the json entry
    # splits[1] will be the json column
    table_dict_flag = False
    if type(splits) == list:
        table_dict_flag = True
        table_column = int(splits[1])
        if type(table_column) is int:
            picked_list = [e[table_column] for e in js_main_dict[splits[0]]]

    else:
        picked_list = js_main_dict[r_named_variable]

    # in a dropdown, all entries in the list must be strings
    picked_list = [str(item) for item in picked_list]

    def is_list_of_lists(inp_list):
        return type(inp_list[0]) == list if inp_list else False

    if r_named_variable:
        if r_filter_value:
            try:
                regex = re.compile(r_filter_value.upper())
                filtered_list = filter(lambda x: regex.search(x.upper()), picked_list)
                picked_list = list(filtered_list)
            except:
                pass

        # make sure Qt has been initialized
        app = get_qt_app()

        # create the custom form
        dlg = Form0()

        # determine index position within dialog list
        # set default value
        dropdown_index = 0
        if r_init_last_text:
            try:
                dropdown_index = index_containing_substring(
                    picked_list, r_init_last_text
                )
            except:
                pass

        if table_dict_flag:
            selection_text = "".join([splits[0], "[", splits[1], "]", " selection:"])
        else:
            selection_text = "".join([r_named_variable, " selection:"])

        text = QInputDialog.getItem(
            dlg,
            "PyQt Dropdown",
            selection_text,
            picked_list,
            max(0, min(dropdown_index, len(picked_list) - 1)),
        )

        # set default for picked item
        picked_item = r_entry
        if text[1]:
            picked_item = text[0]

        # post entry results
        r_entry.Value = picked_item
        r_init_last_range.Value = picked_item
    else:
        r_init_last_range.Value = "ERROR"
