"""
Utility file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

import os
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from openpyxl.cell import Cell

from config import FIRST_COLUMN, LAST_COLUMN, LAST_ROW, SCOPES


def get_merged_range(cell, ws):
    """
    Attempts to find a merged cell which contains the specified cell.
    :param cell: the cell to check.
    :param ws: the worksheet.
    :return: a merged cell range or None.
    """
    for cell_range in ws.merged_cells.ranges:
        if cell.coordinate in cell_range:
            return cell_range
    return None


def get_last_in_range(cell_range, ws):
    """
    Determines the last cell in a cell range.
    :param cell_range: a cell range.
    :param ws: the worksheet.
    :return: the last cell in the cell range.
    """
    return Cell(ws, row=cell_range.max_row, column=cell_range.max_col)


def get_next_cell(cell, ws, check_merged=True):
    """
    Determines the next cell which we need to parse.
    :param cell: the current cell.
    :param ws: the worksheet.
    :param check_merged: whether we should check if the cell is merged.
    :return: the next cell or None if this was the last cell.
    """
    if cell.row == LAST_ROW and cell.column == LAST_COLUMN:
        return None

    cell_range = get_merged_range(cell, ws)
    if cell_range and check_merged:
        return get_next_cell(get_last_in_range(cell_range, ws), ws, False)
    else:
        if cell.column == LAST_COLUMN:
            return Cell(ws, row=cell.row + 1, column=FIRST_COLUMN)
        else:
            return Cell(ws, row=cell.row, column=cell.col_idx + 1)


def get_credentials():
    """
    Attempts to load the credentials from file or generates new ones.
    :return: the credentials to authenticate with Google.
    """
    credentials = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return credentials
