import pandas as pd
import smartsheet
from collections import defaultdict
from smartsheet_dataframe import get_sheet_as_df
import json


class SSConn:
    def __init__(self, token):
        self._ss_client = smartsheet.Smartsheet(token)

    def list_all_sheets(self):
        """Prints the name of all the sheets in your account
        IndexResult object containing an array of Sheet objects limited to the following attributes:
            id
            accessLevel
            createdAt
            modifiedAt
            name
            permalink
            source
        """
        response = self._ss_client.Sheets.list_sheets(include_all=True)
        return response.data

    def create_sheet_in_folder(self, folder_id, dict_sheet):
        """ Creates a new sheet in the specified folder
        :returns accessLevel,columns,id,name,permalink"""
        sheet_spec = self._ss_client.models.Sheet(dict_sheet)
        response = self._ss_client.Folders.create_sheet_in_folder(
            folder_id,
            sheet_spec
        )
        new_sheet = response.result
        print(f"Sheet added, name : {new_sheet.name}, ID : {new_sheet.id}")
        return new_sheet

    def create_sheet_in_folder_from_template(self, folder_id, template_sh_id, new_name):
        """ Create an empty sheet from a template
        "result": {
                "accessLevel": "OWNER",
                "id": 7960873114331012,
                "name": "newsheet",
                "permalink": "https://app.smartsheet.com/b/home?lx=lbKEF1UakfTNJTZ5XkpxWg"}
        """
        response = self._ss_client.Folders.create_sheet_in_folder_from_template(
            folder_id,
            self._ss_client.models.Sheet({
                'name': new_name,
                'from_id': template_sh_id
            })
        )
        if not response.message == "SUCCESS":
            raise "Failed to create new sheet from template"
        return response.result

    def copy_sheet(self, dest_type, dest_id, sh_id, new_name):
        """ Creates a copy of an existing sheet
        :returns    "message": "SUCCESS",
                    "resultCode": 0,
                    "result": {
                        "id": 4366633289443204,
                        "name": "newSheetName",
                        "accessLevel": "OWNER",
                        "permalink": "https://{base_url}?lx=lB0JaOh6AX1wGwqxsQIMaA" """
        response = self._ss_client.Sheets.copy_sheet(
            sh_id,
            self._ss_client.models.ContainerDestination({
                'destination_type': dest_type,
                'destination_id': dest_id,
                'new_name': new_name
            })
        )
        return response.result

    def get_folder(self, folder_id):
        folder = self._ss_client.Folders.get_folder(folder_id)
        return folder

    def _list_all_automation_rules(self, sheet_id):
        """Helper function to get a list of rule ids PER sheet"""
        response = self._ss_client.Sheets.list_automation_rules(
            sheet_id,
            include_all=True
        )
        # print(response.data)
        # parsed = json.loads(response.data)
        # print(json.dumps(parsed, indent=4, sort_keys=True))
        rule_ids = []
        for rule in response.data:
            parsed = json.loads(rule)
            print(json.dumps(parsed, indent=4))
            print(rule)
            rule_ids.append(rule.id)

        return rule_ids

    def disable_all_automation_rules_per_sheet(self, list_sheet_ids, ena_disable=False):
        """ Enables or disables (default) all automations for each sheet in the list
            passed in """

        # set the spec
        automation_spec = self._ss_client.models.AutomationRule({
            'enabled': ena_disable,
            'action': {
                'type': 'NOTIFICATION_ACTION'   # always needs to be included
            }
        })
        # enable/disable each rule for each sheet
        for sheet_id in list_sheet_ids:
            rule_ids = self._list_all_automation_rules(sheet_id)
            for rule_id in rule_ids:
                response = self._ss_client.Sheets.update_automation_rule(
                    sheet_id,
                    rule_id,
                    automation_spec
                )
                print(response.result)

    def get_workspace_share(self, worskspace_id):
        """ Returns a list of all shares of the passed in work space ID"""
        response = self._ss_client.Workspaces.list_shares(
            worskspace_id,
            include_all=True
        )

        shared = []
        for share in response.data:
            print(share)
            shared.append(share)

        [print(s.name, s.email, s.access_level) for s in shared]


class Sheet(SSConn):
    def __init__(self, token, id):
        super().__init__(token)
        self.sheet = self._ss_client.Sheets.get_sheet(id)  # , page_size=page_size)
        self._rows_to_update = []
        self._rows_to_add = []
        self._rows_to_delete = []
        self._cell_list = []
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

    def add_to_delete_rowid(self, rowid):
        self._rows_to_delete.append(rowid)

    def get_df(self):
        """Returns sheet as a dataframe with row_id and parent_id"""
        if len(self.sheet.rows) > 0:
            df = get_sheet_as_df(sheet_obj=self.sheet)
        else:
            # create an empty dataframe
            df = pd.DataFrame()
        return df

    def commit_delete_rows2(self):
        print(f"Sheet ID = {self.sheet.id}")
        print(self._rows_to_delete)
        cnt = 0
        delete_buffer = []

        while len(self._rows_to_delete) != 0:
            if len(self._rows_to_delete) > 100:
                for row in range(100):
                    delete_buffer.append(self._rows_to_delete.pop())
                response = self._ss_client.Sheets.delete_rows(self.sheet.id, delete_buffer)
                cnt += 100
                delete_buffer = []
            else:
                response = self._ss_client.Sheets.delete_rows(self.sheet.id, self._rows_to_delete)
                cnt += len(self._rows_to_delete)
                self._rows_to_delete = []

        print(f"Deleted {cnt} rows from sheet '{self.sheet.name}'")

    def commit_delete_rows(self):
        if len(self._rows_to_delete) > 0:
            response = self._ss_client.Sheets.delete_rows(self.sheet.id, self._rows_to_delete)
            print("Deleted {} rows from {}".format(str(len(self._rows_to_delete)), self.sheet.name))

    def get_row_count(self):
        return self.sheet.total_row_count

    def rename(self, new_name):
        orig_name = self.sheet.name
        response = self._ss_client.Sheets.update_sheet(
            self.sheet.id,
            self._ss_client.models.Sheet({
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
                self._ss_client.models.CopyOrMoveRowDirective({
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
        """ Modifies an individual cell and adds to the rows_to_update list for group commit
            Note: can only modify one cell in row """
        new_cell = self._ss_client.models.Cell()  # first update the cell value according to the column
        new_cell.column_id = self._col_map[col_name]
        new_cell.value = value
        new_cell.strict = False
        # build the row to update
        new_row = self._ss_client.models.Row()  # now set the correct row
        new_row.id = rowid
        new_row.cells.append(new_cell)  # set this row with the updated cell
        self._rows_to_update.append(new_row)

    def update_cell(self, col_name, value):
        """Updates a cell value and returns the new cell.  Needs to be used with a list"""
        new_cell = self._ss_client.models.Cell()
        new_cell.column_id = self._col_map[col_name]
        new_cell.value = value
        new_cell.strict = False
        return new_cell

    def update_multiple_cells_single_row(self, rowid, cell_list):
        """ Used with update_cell and requires a local cell_list that is initialised for each row """
        new_row = self._ss_client.models.Row()  # now set the correct row
        new_row.id = rowid
        new_row.cells = cell_list  # set this row with the updated cell
        return self._rows_to_update.append(new_row)

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
        else:
            print(f"No updates to commit {self.sheet.name}")

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

    def add_row_cells_dict(self, cells_dict):
        """
        Creates new cells and row, appends to the rows_to_add list for group commit
        """
        new_row = self._ss_client.models.Row()  # create a new row object
        new_row.to_top = True

        for col, value in cells_dict.items():
            # print(cell_tup_list[tup][1])
            # print(self._col_map[cell_tup_list[tup][0]])
            # print(f"Tuple {tup}")
            # print(f"Length {len(cell_tup_list)}")
            # print('**************')
            new_row.cells.append({
                'column_id': self._col_map[col],
                'value': value,
                'strict': False
            })

        self._rows_to_add.append(new_row)


    def add_row_cells_tup(self, cell_tup_list):
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
            print(f"No rows to commit {self.sheet.name}")

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

    # def delete_rows(self, rows_to_delete):
    #     if len(rows_to_delete) > 0:
    #         response = self._ss_client.Sheets.delete_rows(self.sheet.id, rows_to_delete)
    #         print("Deleted {} rows from {}".format(str(len(rows_to_delete)),self.sheet.name))
    #     else: print("No rows added to delete list")

    def sort(self, col_name, direction='ASCENDING'):
        sort_specifier = self._ss_client.models.SortSpecifier({
            'sort_criteria': [self._ss_client.models.SortCriterion({
                'column_id': self._col_map[col_name],
                'direction': direction
            })]
        })
        sh_ret = self._ss_client.Sheets.sort_sheet(self.sheet.id, sort_specifier)
        # print(f"Rows sorted by {col_name} on {self.sheet.name}")

    def print_sheet(self):
        """Prints rows in csv"""
        for row in self.sheet.rows:
            values = []
            for col in self._col_map.values():
                values.append(row.get_column(col).value)
            print(*values, sep=', ')








    # rule_ids = conn.list_all_automation_rules(4614223801149316)
    # for rule in rule_ids:
    #     print(rule)

    # conn.disable_all_automation_rules_per_sheet([4614223801149316])

