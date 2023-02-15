import logging
from xlsxwriter.utility import xl_cell_to_rowcol


import ffxiv_automated_collectible_tracker.lodestone as lodestoneapi
from ffxiv_automated_collectible_tracker.gsheets import BatchUpdate, GSheets, DEFAULT_SHEET_NAME


logger = logging.getLogger(__name__)


def get_colours(sheets_config):
    """Prepare colour dictionaries for Google Spreadsheets using the details from the sheet config.

    :param sheets_config: The configuration dictionary for the Spreadsheets.
    :return: All the colours.
    """
    colours = sheets_config["Colours"]
    colourheading = {
        "rgbColor": {
            "red": colours["ColourHeading"]["r"] / 255.0,
            "green": colours["ColourHeading"]["g"] / 255.0,
            "blue": colours["ColourHeading"]["b"] / 255.0,
            "alpha": 1
        }
    }
    colourcharcol = {
        "rgbColor": {
            "red": colours["ColourCharCol"]["r"] / 255.0,
            "green": colours["ColourCharCol"]["g"] / 255.0,
            "blue": colours["ColourCharCol"]["b"] / 255.0,
            "alpha": 1
        }
    }
    colourhasitem = {
        "rgbColor": {
            "red": colours["ColourHasItem"]["r"] / 255.0,
            "green": colours["ColourHasItem"]["g"] / 255.0,
            "blue": colours["ColourHasItem"]["b"] / 255.0,
            "alpha": 1
        }
    }
    colournotitem = {
        "rgbColor": {
            "red": colours["ColourNotItem"]["r"] / 255.0,
            "green": colours["ColourNotItem"]["g"] / 255.0,
            "blue": colours["ColourNotItem"]["b"] / 255.0,
            "alpha": 1
        }
    }
    colourallitem = {
        "rgbColor": {
            "red": colours["ColourAllItem"]["r"] / 255.0,
            "green": colours["ColourAllItem"]["g"] / 255.0,
            "blue": colours["ColourAllItem"]["b"] / 255.0,
            "alpha": 1
        }
    }
    return colourheading, colourcharcol, colourhasitem, colournotitem, colourallitem


def prepare_spreadsheet(
    gsheets: GSheets,
    spreadsheet_id: str,
    spreadsheet_config: dict,
    title_row_index: int,
    name_col_index: int,
    colourheading: dict,
    colourhasitem: dict,
    colournotitem: dict,
    colourallitem: dict,
    results: list,
):
    """Reset the specified Spreadsheet to a default state.
    Then set up the sheets, headings, and formatting for collectible tracking.

    :param gsheets: The gsheets connection object.
    :param spreadsheet_id: The Google Spreadsheet ID.
    :param spreadsheet_config: The configuration dictionary for the Spreadsheet.
    :param title_row_index: The index of the headings\title row.
    :param name_col_index: The index of the character names column.
    :param colourheading: Colour dict for the heading row.
    :param colourhasitem: Colour dict for having an item.
    :param colournotitem: Colour dict for not having an item.
    :param colourallitem: Colour dict for having a full row.
    :param results: The list of total API responses.
    :return: None
    """
    result = gsheets.reset_spreadsheet(spreadsheet_id)
    results.append(result)
    # Create and setup sheets with headings and formatting.
    logger.info("Creating and setting up sheets.")
    batchupdate = gsheets.create_new_batchupdate(spreadsheet_id)
    for sheet_config in spreadsheet_config["sheets"]:
        sheet_name = sheet_config["title"]
        headings = sheet_config["Values"]
        sheet_obj = gsheets.add_sheet(spreadsheet_id, sheet_name)
        sheet_id = sheet_obj["properties"]["sheetId"]
        sheet_config["sheetId"] = sheet_id

        # Headings
        logger.info(f"Adding headings to Sheet '{sheet_name}'.")
        existing_heading_vals = []
        existing_heading = []
        new_headings = [heading for heading in headings]
        if "rowData" in sheet_obj["data"][0]:
            existing_heading = sheet_obj["data"][0]["rowData"][title_row_index]["values"][name_col_index + 1:]
            for index, cell in enumerate(existing_heading_vals[-1::-1]):
                if "effectiveValue" in cell:
                    break
                existing_heading_vals.pop()
        for cell in existing_heading:
            cell_value = ""
            if "effectiveValue" in cell:
                cell_value = cell["effectiveValue"]["stringValue"]
            existing_heading_vals.append(cell_value)
        if len(existing_heading_vals) > len(headings):
            new_headings += [""] * (len(existing_heading_vals) - len(headings))
        cols = []
        for heading in new_headings:
            if heading == "":
                cell_data = {"userEnteredValue": {"stringValue": ""}}
            else:
                heading_display_name = heading["DisplayName"]
                heading_item_type = heading["ItemType"]
                heading_item_name = heading["ItemName"]

                cell_data = {
                    "userEnteredValue": {"stringValue": heading_display_name},
                    "note": heading_item_type + ":" + heading_item_name,
                    "userEnteredFormat": {
                        "borders": {
                            "top": {"style": "SOLID"},
                            "right": {"style": "SOLID"},
                            "bottom": {"style": "DOUBLE"},
                            "left": {"style": "SOLID"}
                        },
                        "textFormat": {"bold": True},
                    }
                }
                if colourheading:
                    cell_data["userEnteredFormat"]["backgroundColorStyle"] = colourheading
            cols.append(cell_data)

        batchupdate.add_row(sheet_id, title_row_index, name_col_index + 1, cols)

        logger.info(f"Adding formatting to Sheet '{sheet_name}'.")
        # Formatting
        batchupdate.freeze_row_col(sheet_id, title_row_index, name_col_index)
        if colourhasitem:
            batchupdate.add_cond_fmt_rule_text_eq(sheet_id, "Y", colourhasitem)
        if colournotitem:
            batchupdate.add_cond_fmt_rule_text_eq(sheet_id, "N", colournotitem)
        if colourallitem:
            batchupdate.add_cnd_fmt_rule_column_formula(
                sheet_id,
                name_col_index,
                '=AND(COUNTIF(C1:1,"N")=0,COUNTIF(C1:1,"Y")>1)',
                colourallitem
            )
    result = batchupdate.execute()
    results.append(result)
    logger.info("Completing Spreadsheet Reset...")

    spreadsheet_simple = gsheets.get_sheet_list(spreadsheet_id)
    if DEFAULT_SHEET_NAME in spreadsheet_simple:
        result = gsheets.delete_sheet(spreadsheet_id, DEFAULT_SHEET_NAME)
        results.append(result)


def update_char_row_in_sheet(
    batchupdate: BatchUpdate,
    sheet_id: int,
    sheet_obj: dict,
    headings: list,
    title_row_index: int,
    name_col_index: int,
    colourcharcol: dict,
    charnum: int,
    character_details: dict,
):
    """Add instructions for a single character's row details for a single sheet to the provided BatchUpdate object..

    :param batchupdate: A BatchUpdate object to append instructions to.
    :param sheet_id: The ID of the current sheet/tab.
    :param sheet_obj: A dictionary with additional details of the current sheet/tab.
    :param headings: The list of headings for this sheet/tab.
    :param title_row_index: The index of the headings\title row.
    :param name_col_index: The index of the character names column.
    :param colourcharcol: The colour for the character column.
    :param charnum: The character number - the relative row number to the heading row.
    :param character_details: A dictionary of the character's lodestone details.
    :return: None
    """
    fullname = character_details["Name"]
    world = character_details["World"]
    first_name, second_name = fullname.split()
    char_and_world = "@".join([fullname, world])
    char_id = character_details["ID"]
    character_mounts = character_details["Mounts"]
    character_achievements = character_details["Achievements"]
    cols = []

    hyperlink = f"https://eu.finalfantasyxiv.com/lodestone/character/{char_id}/"
    cell_data = {
        "userEnteredValue": {
            "formulaValue": f'=HYPERLINK("{hyperlink}", "{first_name}")'
        },
        "userEnteredFormat": {
            "textFormat": {"foregroundColorStyle": {"themeColor": "TEXT"}},
            "borders": {
                "top": {"style": "SOLID"},
                "right": {"style": "SOLID_MEDIUM"},
                "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}
            },
        },
        "note": char_and_world
    }
    if colourcharcol:
        cell_data["userEnteredFormat"]["backgroundColorStyle"] = colourcharcol
    cols.append(cell_data)

    existing_row_cells = []
    if "rowData" in sheet_obj["data"][0]:
        if len(sheet_obj["data"][0]["rowData"]) > title_row_index + charnum:
            if "values" in sheet_obj["data"][0]["rowData"][title_row_index + charnum]:
                if len(sheet_obj["data"][0]["rowData"][title_row_index + charnum]["values"]) > name_col_index:
                    existing_row_cells = sheet_obj["data"][0]["rowData"][title_row_index + charnum]["values"][name_col_index:]
    for index, cell in enumerate(existing_row_cells[-1::-1]):
        if "effectiveValue" in cell:
            break
        existing_row_cells.pop()

    values = [""] * len(headings)
    if len(existing_row_cells) > len(headings):
        values += [""] * (len(existing_row_cells) - len(headings))
    for heading_id, heading in enumerate(headings):
        if "ItemType" not in heading or "ItemName" not in heading:
            continue
        heading_item_type = heading["ItemType"]
        heading_item_name = heading["ItemName"]
        list_to_use = None
        if heading_item_type == "Mount":
            list_to_use = character_mounts
        elif heading_item_type == "Achievement":
            list_to_use = character_achievements

        if list_to_use:
            if heading_item_name in list_to_use:
                values[heading_id] = "Y"
            else:
                values[heading_id] = "N"

    for val_num, value in enumerate(values, start=1):
        if value:
            cell_data = {
                "userEnteredValue": {"stringValue": value},
                "userEnteredFormat": {
                    "borders": {
                        "right": {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                    }
                }
            }
        elif val_num < len(headings) and val_num < len(existing_row_cells):
            cell_data = existing_row_cells[val_num]
        else:
            cell_data = {"userEnteredValue": {"stringValue": ""}}
        cols.append(cell_data)

    batchupdate.add_row(sheet_id, title_row_index + charnum, name_col_index, cols)


async def update_formatted_spreadsheet(
    gsheets,
    spreadsheet_id,
    spreadsheet_config,
    title_row_index,
    name_col_index,
    characters_list,
    colourcharcol,
    results,
):
    """Loop through sheets and characters to update a spreadsheet.

    :param gsheets: The gsheets connection object.
    :param spreadsheet_id: The Google Spreadsheet ID.
    :param spreadsheet_config: The configuration dictionary for the Spreadsheet.
    :param title_row_index: The index of the headings\title row.
    :param name_col_index: The index of the character names column.
    :param characters_list: A list of Final Fantasy XIV character names.
    :param colourcharcol: Colour dict for character column.
    :param results: The list of total API responses.
    :return: None
    """
    for charnum, char_and_world in enumerate(characters_list, start=1):
        logger.info(f"{char_and_world=}")
        fullname, world = char_and_world.split("@")
        character_details = await lodestoneapi.get_char_details(fullname, world, achievements=True, mounts=True)

        batchupdate = gsheets.create_new_batchupdate(spreadsheet_id)
        for sheet_config in spreadsheet_config["sheets"]:
            sheet_name = sheet_config["title"]

            headings = sheet_config["Values"]
            sheet_id = sheet_config["sheetId"]
            sheet_obj = gsheets.get_sheet(spreadsheet_id, sheet_name)

            update_char_row_in_sheet(
                batchupdate,
                sheet_id,
                sheet_obj,
                headings,
                title_row_index,
                name_col_index,
                colourcharcol,
                charnum,
                character_details,
            )
        result = batchupdate.execute()
        results.append(result)


async def update_spreadsheets(
    cred_filename: str,
    sheets_config: dict,
    characters_list: [str] = None,
) -> list:
    """

    :param cred_filename: The filepath to the file with the google sheets credentials.
    :param sheets_config:
    :param characters_list:
    :return:
    """
    gsheets = GSheets(cred_filename)

    name_col = sheets_config["NamesColumn"]
    title_row = str(sheets_config["HeadingsRow"])
    title_row_index, name_col_index = xl_cell_to_rowcol(name_col + title_row)

    colourheading, colourcharcol, colourhasitem, colournotitem, colourallitem = get_colours(sheets_config)

    results = []
    for spreadsheet_config in sheets_config["Spreadsheets"]:
        spreadsheet_id = spreadsheet_config["spreadsheetId"]

        logger.info(f"{spreadsheet_id=}")

        prepare_spreadsheet(
            gsheets,
            spreadsheet_id,
            spreadsheet_config,
            title_row_index,
            name_col_index,
            colourheading,
            colourhasitem,
            colournotitem,
            colourallitem,
            results
        )

        await update_formatted_spreadsheet(
            gsheets,
            spreadsheet_id,
            spreadsheet_config,
            title_row_index,
            name_col_index,
            characters_list,
            colourcharcol,
            results,
        )

    return results
