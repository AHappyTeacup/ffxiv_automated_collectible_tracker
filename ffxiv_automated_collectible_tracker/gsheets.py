import json
import logging
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from xlsxwriter.utility import xl_cell_to_rowcol


logger = logging.getLogger(__name__)


DEFAULT_SHEET_NAME = "Sheet1"

# Re-used Strings
REQUESTS = "requests"
# Request Types
ADDSHEET = "addSheet"
AUTORESIZE = "autoResizeDimensions"
ADDCNDFMTRULE = "addConditionalFormatRule"
DELETESHEET = "deleteSheet"
UPDATESHEETPROPERTIES = "updateSheetProperties"
UPDATECELLS = "updateCells"
REPEATCELL = "repeatCell"
# Common Keywords
SHEETS = "sheets"
SHEETID = "sheetId"
CELL = "cell"
PROPERTIES = "properties"
FIELDS = "fields"


class BatchUpdate:
    """This object allows for construction BatchUpdate JSON blocks in a very dynamic but readable manner."""
    def __init__(self, spreadsheet_resource, spreadsheet_id) -> dict:
        """
        :param spreadsheet_resource: The spreadsheet_resource attribute of a GSheets object.
        :param spreadsheet_id: The ID of the Spreadsheet this update is intended for.
        """
        self.spreadsheet_resource = spreadsheet_resource
        self.spreadsheet_id = spreadsheet_id
        self.body = {REQUESTS: []}

    def execute(self):
        """POST the BatchUpdate JSON to the GoogleSheets API."""
        logger.info(json.dumps(self.body, indent=4))
        batchupdate = self.spreadsheet_resource.batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=self.body
        )
        results = batchupdate.execute()
        logger.debug(results)
        return results

    def add_new_request(self, request_obj: dict) -> None:
        """Allow the appending of an existing request object.

        :param request_obj: A dictionary.
        :return: None
        """
        self.body[REQUESTS].append(request_obj)

    def create_new_request(self, request_type: str) -> dict:
        """Create a new request as part of this update, and return the dictionary for the calling code to provide
        parameters.

        :param request_type: A String. The type of request to make. See the API documentation for expected values.
        :return: An empty dictionary. See the API documentation for expected contents.
        """
        request_params = {}
        new_request = {request_type: request_params}
        self.body[REQUESTS].append(new_request)
        return request_params

    def freeze_row_col(self, sheet_id: int, row_index: int, col_index: int) -> None:
        """Add requests to freeze the specified row and column.

        :param sheet_id: The ID of the sheet/tab of this Spreadsheet.
        :param row_index: The row to freeze.
        :param col_index: The column to freeze.
        :return: None
        """
        row_count = row_index + 1
        col_count = col_index + 1
        updateproperties_request = self.create_new_request(UPDATESHEETPROPERTIES)
        updateproperties_request[PROPERTIES] = {
            SHEETID: sheet_id,
            "gridProperties": {"frozenRowCount": row_count, "frozenColumnCount": col_count}
        }
        updateproperties_request[FIELDS] = "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"

    def add_cond_fmt_rule_text_eq(self, sheet_id: int, text: str, colour_object: dict) -> None:
        """Add a conditional formatting rule for setting the background and text colour to the specified colour,
        when the cell contents are equal to the specified text.

        :param sheet_id: The ID of the sheet/tab of this Spreadsheet.
        :param text: The value that the cell contents should match to trigger this formatting rule.
        :param colour_object: A dictionary describing the colour to apply.
        :return: None
        """
        addcndfmt_request = self.create_new_request(ADDCNDFMTRULE)
        addcndfmt_request["rule"] = {
            "ranges": [{"sheetId": sheet_id, "startRowIndex": 0, "startColumnIndex": 0}],
            "booleanRule": {
                "condition": {
                    "type": "TEXT_EQ",
                    "values": [{"userEnteredValue": text}]
                },
                "format": {
                    "backgroundColorStyle": colour_object,
                    "textFormat": {"foregroundColorStyle": colour_object}
                }
            }
        }
        addcndfmt_request["index"] = 0

    def add_cnd_fmt_rule_column_formula(self, sheet_id: int, col_index: int, formula: str, colour_object: dict) -> None:
        """Add a conditional formatting rule for setting the background colour of a column to the specified colour,
        when the specified formula returns TRUE.

        :param sheet_id: The ID of the sheet/tab of this Spreadsheet.
        :param col_index: The index of the column where the colour should be adjusted.
        :param formula: The formula to use to decide when to change the column colour.
        :param colour_object: A dictionary describing the colour to apply.
        :return: None
        """
        addcndfmt_request = self.create_new_request(ADDCNDFMTRULE)
        addcndfmt_request["rule"] = {
            "ranges": [
                {"sheetId": sheet_id, "startRowIndex": 0, "startColumnIndex": col_index, "endColumnIndex": col_index+1}
            ],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": formula}]
                },
                "format": {
                    "backgroundColorStyle": colour_object,
                }
            }
        }
        addcndfmt_request["index"] = 0

    def add_row(self, sheet_id: int, row_index: int, col_index: int, cell_list: list) -> None:
        """Add a row to the sheet.

        :param sheet_id: The ID of the sheet/tab of this Spreadsheet.
        :param row_index: The index to add the row at.
        :param col_index: The column the row should start at.
        :param cell_list: A list of cell objects.
        :return: None
        """
        updatecells_request = self.create_new_request(UPDATECELLS)
        # The FIELDS property basically determines what properties of the cells will be updated by this request.
        # If "*" is provided, and no value is provided for a field for the cell, then it will be set to default.
        # updatecells_request[FIELDS] = "userEnteredValue.stringValue,formattedValue,note"
        updatecells_request[FIELDS] = "*"
        updatecells_request["start"] = {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": col_index}
        # The cell_list, a list of cells, will be the values of a single row.
        updatecells_request["rows"] = [{"values": cell_list}]

    def add_col(self, sheet_id: int, row_index: int, col_index: int, cell_list: list) -> None:
        """Add a column to the sheet.

        :param sheet_id: The ID of the sheet/tab of this Spreadsheet.
        :param row_index: The row the column should start at.
        :param col_index: The index to add the column at.
        :param cell_list: A list of cell objects.
        :return: None
        """
        updatecells_request = self.create_new_request(UPDATECELLS)
        # The FIELDS property basically determines what properties of the cells will be updated by this request.
        # If "*" is provided, and no value is provided for a field for the cell, then it will be set to default.
        # updatecells_request[FIELDS] = "userEnteredValue.stringValue,formattedValue,note"
        updatecells_request[FIELDS] = "*"
        updatecells_request["start"] = {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": col_index}
        # The cell_list, a list of cells, needs to be separated into a cell per row.
        rows = []
        for cell in cell_list:
            rows.append({"values": [cell]})
        updatecells_request["rows"] = rows


class GSheets:
    """I could have used an exisitng API module, but this was more fun."""
    def __init__(self, service_account_filename: str):
        """
        :param service_account_filename: The path to the file containing your google API auth credentials.
        """
        logging.debug("Reading credentials for Google Sheets API from '%s'" % service_account_filename)
        creds = service_account.Credentials.from_service_account_file(service_account_filename)
        service_resource = build('sheets', 'v4', credentials=creds)
        self.spreadsheets_resource = service_resource.spreadsheets()
        self.spreadsheets_values_resource = self.spreadsheets_resource.values()

    def get_spreadsheet(self, spreadsheet_id: str, include_grid_data=True) -> dict:
        """Get a spreadsheet object.

        :param spreadsheet_id: The ID for the Google Spreadsheet to be gotten.
        :param include_grid_data: Boolean. Include a verbose amount of details per sheet.
        :return: Dictionary. Spreadsheet details.
        """
        logger.info(f"Getting spreadsheet: '{spreadsheet_id}'.")
        spreadsheet_get_http = self.spreadsheets_resource.get(
            spreadsheetId=spreadsheet_id,
            includeGridData=include_grid_data
        )
        result = spreadsheet_get_http.execute()
        logger.debug(result)
        return result

    def get_sheet(self, spreadsheet_id: str, sheet_name: str) -> dict:
        """Get a sheet/tab using the human readable name for it.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_name: The human readable name of a sheet.
        :return: A dictionary with the details of the requested sheet.
        """
        logger.info(f"Getting Sheet: '{sheet_name}'.")
        spreadsheets_values_get_http = self.spreadsheets_resource.get(
            spreadsheetId=spreadsheet_id,
            ranges=["'%s'" % sheet_name],
            includeGridData=True
        )
        result = None
        while result is None:
            try:
                result = spreadsheets_values_get_http.execute()
            except HttpError as e:
                if e.status_code != 429:
                    raise e
            time.sleep(1)
            logger.info(f"Too Many Requests to Google...")
            result = None
        logger.debug(result)
        sheet = result["sheets"][0]
        logger.debug(f"Found Sheet: {sheet}")
        return sheet

    def add_sheet(self, spreadsheet_id: str, sheet_name: str) -> dict:
        """Create a sheet/tab.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_name: The human readable name for the new sheet/tab.
        :return: A dictionary with the details of the requested sheet.
        """
        sheet_name = "%s" % sheet_name
        logger.info(f"Adding New Sheet: '{sheet_name}'.")
        try:
            batchupdate = BatchUpdate(self.spreadsheets_resource, spreadsheet_id)
            addsheet_request = batchupdate.create_new_request(ADDSHEET)
            addsheet_request[PROPERTIES] = {
                "title": sheet_name,
                "hidden": False
            }
            result = batchupdate.execute()
            logger.debug(result)
        except HttpError as e:
            logger.info(f"Did not add new Sheet: '{sheet_name}'.'")
            logger.debug(e)
        sheet = self.get_sheet(spreadsheet_id, sheet_name)
        # Result already logged by get_sheet()
        return sheet

    def reset_spreadsheet(self, spreadsheet_id: str, default_sheet_name: str = DEFAULT_SHEET_NAME) -> dict:
        """Reset a Google Spreadsheet to a single blank "Sheet1" sheet/tab.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param default_sheet_name: The name for the remaining 'default' sheet. Usually "Sheet1".
        :return: Dictionary containing the API response.
        """
        logger.info(f"Resetting Spreadsheet: '{spreadsheet_id}', to single Sheer: '{default_sheet_name}'")
        self.add_sheet(spreadsheet_id, default_sheet_name)
        spreadsheet = self.get_spreadsheet(spreadsheet_id)

        batchupdate = BatchUpdate(self.spreadsheets_resource, spreadsheet_id)
        for sheet in spreadsheet[SHEETS]:
            if sheet[PROPERTIES]["title"] != default_sheet_name:
                sheet_id = sheet[PROPERTIES][SHEETID]
                logger.info(f"Scheduling to delete Sheet: `{sheet_id}`")
                deletesheet_request = batchupdate.create_new_request(DELETESHEET)
                deletesheet_request[SHEETID] = sheet_id
        result = batchupdate.execute()
        logger.debug(result)
        return result

    def get_sheet_list(self, spreadsheet_id: str) -> list:
        """Get a list of the names of sheets/tabs in a Google Spreadsheet.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :return: The list of sheet names.
        """
        logger.info(f"Getting sheet list for Spreadsheet: '{spreadsheet_id}'.")
        spreadsheet_config_simple = self.get_spreadsheet(spreadsheet_id, include_grid_data=False)
        sheet_list = []
        for sheet in spreadsheet_config_simple["sheets"]:
            sheet_list.append(sheet['properties']["title"])
        logger.debug(f"Found sheets: {sheet_list}")
        return sheet_list

    def get_cells_simple(self, spreadsheet_id: str, sheet_range: str) -> dict:
        """Get a range of cells from a sheet.
        Do not includeGridData to return minimal result details.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_range: The query range. See the API documentation.
        :return: The API response.
        """
        logger.info(f"Getting Cells (Simple): '{sheet_range}'.")
        spreadsheets_values_get_http = self.spreadsheets_values_resource.get(
            spreadsheetId=spreadsheet_id,
            range=[sheet_range],
        )
        result = spreadsheets_values_get_http.execute()
        logger.debug(result)
        return result

    def get_cells_complex(self, spreadsheet_id: str, sheet_range: str) -> dict:
        """Get a range of cells from a sheet.
        includeGridData will include an enormous amount of information in the response.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_range: The query range. See the API documentation.
        :return: The API response.
        """
        logger.info(f"Getting Cells (Verbose): '{sheet_range}'.")
        get = self.spreadsheets_resource.get(
            spreadsheetId=spreadsheet_id,
            ranges=[sheet_range],
            includeGridData=True
        )
        results = get.execute()
        logger.debug(results)
        return results

    def delete_sheet(self, spreadsheet_id: str, sheet_name: str) -> dict:
        """Delete a sheet/tab from a Google Spreadsheet.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_name:  The human readable name for sheet/tab.
        :return: The API response.
        """
        logger.info(f"Deleting Sheet: `{sheet_name}`.")
        sheet = self.get_sheet(spreadsheet_id, sheet_name)
        if sheet:
            sheet_id = sheet[PROPERTIES][SHEETID]

            batchupdate = BatchUpdate(self.spreadsheets_resource, spreadsheet_id)
            deletesheet_request = batchupdate.create_new_request(DELETESHEET)
            deletesheet_request[SHEETID] = sheet_id
            result = batchupdate.execute()
            logger.debug(result)
            return result
        logger.debug(f"The sheet `{sheet_name}` was not found to be deleted.")

    def fix_column_width(self, spreadsheet_id: str, sheet_id: int, col_letter: str) -> dict:
        """Adjust the width of a column.

        :param spreadsheet_id: The ID for the Google Spreadsheet.
        :param sheet_id: The ID of the sheet/tab.
        :param col_letter: The letter for the column to adjust.
        :return: The API response.
        """
        logger.info(f"Fixing column width for column {col_letter}, of sheet `{sheet_id}`.")
        name_col_index = xl_cell_to_rowcol(col_letter+"0")[1]

        batchupdate = BatchUpdate(self.spreadsheets_resource, spreadsheet_id)
        autoresize_request = batchupdate.create_new_request(AUTORESIZE)
        autoresize_request["dimensions"] = {
            SHEETID: sheet_id,
            "dimension": "COLUMNS",
            "startIndex": name_col_index,
            "endIndex": name_col_index+1
        }
        result = batchupdate.execute()
        logger.debug(result)
        return result

    def create_new_batchupdate(self, spreadsheet_id: str) -> BatchUpdate:
        """Create and return a new BatchUpdate object."""
        batchupdate = BatchUpdate(self.spreadsheets_resource, spreadsheet_id)
        return batchupdate
