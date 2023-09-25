import smartsheet
import datetime
from collections import defaultdict
from smartsheet_dataframe import get_as_df, get_sheet_as_df
import sys

import config_secret
import config_secret as config



def test():
    NUVOLO_MUST_FOLDER_ID = 607072942352260 #8974313631049604

    siebel_sheet = {'name': 'Spire Must Archived 2020',
                    'columns': [
                        {
                            'title': 'Asset Tag',
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'GE System ID',
                            'primary': True,
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'System Name',
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'Job No',
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'Job Type',
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'Call Date',
                            'type': 'DATE'
                        }, {
                            'title': 'Nbr Debriefs',
                            'type': 'TEXT_NUMBER'
                        }, {
                            'title': 'Processed Date',
                            'type': 'DATE'
                        }
                    ]
                    }

    conn = SSConn(config.smartsheet_token)

    all_sheets_lst = conn.list_all_sheets()

    spire_lst = []

    # year : sheet_id
    sh_dict = {}
    all_rows = []

    # this creates an integer value of current year
    year_now = yr = datetime.datetime.now().year

    for sheet in all_sheets_lst:
        if 'Spire MUST Archived' in sheet.name:
            spire_lst.append(sheet)
            print(f"Added {sheet.name}")
            sh_year = datetime.datetime.strptime(sheet.name[-4:],"%Y").year
            # add Sheet object to dictionary if within last 2 years
            if sh_year >= year_now - 2:
                sh_dict[sh_year] = Sheet(config_secret.smartsheet_token,sheet.id)

    for l in spire_lst:
        print(f"Sheet details = {l.name},{l.id},{l.name[-4:]}")

    if year_now not in sh_dict.keys():
        # create new name with current year
        siebel_sheet['name'] = f'Spire MUST Archived {str(year_now)}'
        print('Adding new sheet')
        new_sheet = conn.create_sheet_in_folder(NUVOLO_MUST_FOLDER_ID,siebel_sheet)
        sh_dict[year_now] = Sheet(config_secret.smartsheet_token, new_sheet.id)
    else:
        print("All years already exist")

    # now check results
    for k, v in sh_dict.items():
        print(f"Key: {k} ,Sheet {v.sheet.name} with ID {v.sheet.id} has {v.get_row_count()} rows")
        all_rows = all_rows + v.get_col_list('Job No')

    [print(item) for item in all_rows]

    # folder = conn.get_folder(folder_must_imperial_interface)
    # print(type(folder))
    # print(folder.data)
    # for k,v in folder.items():
    #     print(k,v)
    # sh_summary = Sheet(8937600267380612)
    # headers = sh_summary.get_headers()
    # print(headers)


# def init_sheet(id):
#     sheet = Sheet(config.smartsheet_token, id)
#     return sheet

"""
1. Init client
2. Use client to init sheet, report etc
"""


class SSConn:

    def __init__(self, token):
        self._ss_client = smartsheet.Smartsheet(token)
        print("SSConn substantiated")

    def list_all_sheets(self):
        """Prints the name of all the sheets in your account"""
        response = self._ss_client.Sheets.list_sheets(include_all=True)
        # for sheet in response.data:
        #     print(f"Name : {sheet.name}, ID : {sheet.id}")
        return response.data

    def create_sheet_in_folder(self, folder_id, dict_sheet):
        """ Creates a new sheet in the specified folder"""
        sheet_spec = smartsheet.models.Sheet(dict_sheet)
        response = self._ss_client.Folders.create_sheet_in_folder(
            folder_id,
            sheet_spec
        )
        new_sheet = response.result
        print(f"Sheet added, name : {new_sheet.name}, ID : {new_sheet.id}")
        return new_sheet


    def get_folder(self, folder_id):
        folder = self._ss_client.Folders.get_folder(folder_id)
        return folder




class Sheet(SSConn):

    def __init__(self, token, sheet_id):
        super().__init__(token)
        self.sheet = self._ss_client.Sheets.get_sheet(sheet_id)  # , page_size=page_size)
        self._rows_to_update = []
        self._rows_to_add = []
        self._rows_to_delete = []
        self._col_map = {}
        for col in self.sheet.columns:
            self._col_map[col.title] = col.id
        print('Sheet "{}" initialised'.format(self.sheet.name))

    def add_to_delete(self, col_header, value):
        for row in self.sheet.rows:
            cell_value = self.get_cell_value(row, col_header)
            if cell_value == value:
                self._rows_to_delete.append(row.id)
                # print(f"Added {row} to be deleted")

    def get_df(self):
        return get_sheet_as_df(sheet_obj=self.sheet)

    def commit_delete_rows(self):
        if len(self._rows_to_delete) > 0:
            response = self._ss_client.Sheets.delete_rows(self.sheet.id, self._rows_to_delete)
            print("Deleted {} rows from {}".format(str(len(self._rows_to_delete)),self.sheet.name))

    def get_row_count(self):
        return self.sheet.total_row_count

    def rename(self, new_name):
        orig_name = self.sheet.name
        response = self._ss_client.Sheets.update_sheet(
            self.sheet.id,
            smartsheet.models.Sheet({
                'name': new_name,
            })
        )
        if response.message == 'SUCCESS':
            print("'{}' renamed to '{}'".format(orig_name,new_name))
        else:
            print('Rename failed')

    def backup(self, backup_sh_id):
        """Copies all rows to backup sheet"""
        this_sheet_rows = []
        for row in self.sheet.rows:
            this_sheet_rows.append(row.id)

        if len(this_sheet_rows) > 0:
            response = self._ss_client.Sheets.copy_rows(
                self.sheet.id,
                smartsheet.models.CopyOrMoveRowDirective({
                    'row_ids' : this_sheet_rows,
                    'to' : smartsheet.models.CopyOrMoveRowDestination({
                        'sheet_id': backup_sh_id
                    })
                })
            )

    def get_headers(self):
        """Returns a list of the headers"""
        headers = []
        for header in self._col_map.keys():
            headers.append(header)
        return headers

    def get_col_list(self, column_name):
        """
        Returns a list of all values in one column:: useful for lookups
        """
        column_list = []
        for row in self.sheet.rows:
            value = self.get_cell_value(row, column_name)
            column_list.append(value)
        return column_list

    def make_dict(self, col_key):
        """[col_key : { col1: col1val, col2: col2val ... }, ...]"""
        list_of_dicts = []
        for row in self.sheet.rows:
            target_dict = defaultdict(dict)
            headers = self.get_headers()
            key_idx = headers.index(col_key)
            key = headers.pop(key_idx)
            key_value = self.get_cell_value(row, key)
            for header in headers:
                target_dict[key_value][header] = self.get_cell_value(row, header)
            list_of_dicts.append(target_dict)

        return list_of_dicts

    def get_cell_value(self, row, column_name):
        """Used when iterating through a sheet"""
        cell = row.get_column(self._col_map[column_name])
        if cell.value:
            return cell.value
        else:
            return None

    def update_cell_row(self, row, col_name, value):
        """Modifies an individual cell and adds to the rows_to_update list for group commit"""
        new_cell = self._ss_client.models.Cell()  # first update the cell value according to the column
        new_cell.column_id = self._col_map[col_name]
        new_cell.value = value
        new_cell.strict = False
        # build the row to update
        new_row = self._ss_client.models.Row()  # now set the correct row
        new_row.id = row.id
        new_row.cells.append(new_cell)  # set this row with the updated cell
        self._rows_to_update.append(new_row)

    def update_cell_rowid(self, rowid, col_name, value):
        """Modifies an individual cell and adds to the rows_to_update list for group commit"""
        new_cell = self._ss_client.models.Cell()  # first update the cell value according to the column
        new_cell.column_id = self._col_map[col_name]
        new_cell.value = value
        new_cell.strict = False
        # build the row to update
        new_row = self._ss_client.models.Row()  # now set the correct row
        new_row.id = rowid
        new_row.cells.append(new_cell)  # set this row with the updated cell
        self._rows_to_update.append(new_row)

    def commit_update_rows(self):
        """Writes the rows with cell modifications to SS and empties the list on success"""
        num_rows_update = len(self._rows_to_update)
        if num_rows_update > 0:
            result = self._ss_client.Sheets.update_rows_with_partial_success(self.sheet.id, self._rows_to_update)
            if result.message == 'SUCCESS':
                print("Updated {} rows to {}".format(str(num_rows_update), self.sheet.name))
                [self._rows_to_update.remove(item) for item in list(self._rows_to_update)]
            else:
                print(result.message)

    def add_row(self, data_list):
        """Creates new cells and row, appends to the rows_to_add list for group commit
        Note: the data list needs to match the index and formats!!! """
        new_row = self._ss_client.models.Row()  # create a new row object
        new_row.to_top = True
        indexer = 0
        for col_id in self._col_map.values():
            new_row.cells.append({
                'column_id': col_id,
                'value': data_list[indexer],
                'strict': False
            })
            indexer += 1
        self._rows_to_add.append(new_row)


    def add_row_cells(self, cell_tup_list):
        """Creates new cells and row, appends to the rows_to_add list for group commit
        Note: the data list needs to match the index and formats!!!
        (col_name, value)
        """
        new_row = self._ss_client.models.Row()  # create a new row object
        new_row.to_top = True

        for tup in range(len(cell_tup_list)):
            # print(cell_tup_list[tup][1])
            # print(self._col_map[cell_tup_list[tup][0]])
            # print(f"Tuple {tup}")
            # print(f"Length {len(cell_tup_list)}")
            # print('**************')
            new_row.cells.append({
                'column_id': self._col_map[cell_tup_list[tup][0]],
                'value': str(cell_tup_list[tup][1]),
                'strict': False
            })

        self._rows_to_add.append(new_row)

    def commit_add_rows(self):
        """Writes the new rows to SS and empties the list on success"""
        num_rows_to_add = len(self._rows_to_add)
        if num_rows_to_add > 0:
            result = self._ss_client.Sheets.add_rows(self.sheet.id, self._rows_to_add)
            if result.message == "SUCCESS":
                print("Added {} rows to {}".format(str(num_rows_to_add), self.sheet.name))
                self._rows_to_add = []
            else:
                print(result.message)
        else:
            print("No rows to commit!!")

    def delete_all_rows(self):
        """Yep, deletes all the rows"""
        rows_to_delete = []
        total_rows = 0
        cnt = 0
        for row in self.sheet.rows:
            cnt += 1
            rows_to_delete.append(row.id)
            if cnt == 50 and len(rows_to_delete) > 0:
                self._ss_client.Sheets.delete_rows(self.sheet.id, rows_to_delete)
                rows_to_delete.clear()
                cnt = 0
                total_rows = total_rows + 50

        no_rows = len(rows_to_delete)
        total_rows = total_rows + no_rows
        if no_rows > 0:
            response = self._ss_client.Sheets.delete_rows(self.sheet.id, rows_to_delete)
            print("Deleted {} rows from {}".format(str(total_rows), self.sheet.name))

    def delete_rows(self, rows_to_delete):
        if len(rows_to_delete) > 0:
            response = self._ss_client.Sheets.delete_rows(self.sheet.id, rows_to_delete)
            print("Deleted {} rows from {}".format(str(len(rows_to_delete)),self.sheet.name))
        else: print("No rows added to delete list")

    def sort(self, col_name, direction='ASCENDING'):
        sort_specifier = smartsheet.models.SortSpecifier({
            'sort_criteria': [smartsheet.models.SortCriterion({
                'column_id': self._col_map[col_name],
                'direction': direction
            })]
        })
        sh_ret = self._ss_client.Sheets.sort_sheet(self.sheet.id, sort_specifier)

    def print_sheet(self):
        """Prints rows in csv"""
        for row in self.sheet.rows:
            values = []
            for col in self._col_map.values():
                values.append(row.get_column(col).value)
            print(*values, sep=', ')


if __name__ == '__main__':
    test()


