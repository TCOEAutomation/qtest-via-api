import pandas as pd
from pprint import pprint
import sys
import json
import userConfig

import json,requests, systemConfig,importlib
from test_case_details import TestCaseDetails

class TestCaseStructureBuilder():
    def __init__(self):
        if userConfig.default_automation_report:
            xls = pd.ExcelFile("DataSheet.xlsx")
            self.df = xls.parse(xls.sheet_names[1], keep_default_na=False, na_values=userConfig.nan_values)
            self.allowed_columns = ["qTest TC ID"]
        else:
            xls = pd.ExcelFile(userConfig.excelFileName)
            self.df = xls.parse(systemConfig.current_sheet, keep_default_na=False, na_values=userConfig.nan_values)
            self.allowed_columns = userConfig.allowed_columns
        self.test_data = self.df.to_dict()
        self.total_test_cases = len(self.df.index)
        self.test_case_details = TestCaseDetails()

    def set_payload(self, payload):
        self.testCaseStructure = payload
        self.test_case_details.setup()

    def validate_excel_file(self):
        self.validate_allowed_fields()
        self.validate_mandatory_fields()

    def get_qtest_id(self, test_case_id):
        test_case_index = self.get_test_case_index(test_case_id)
        qtest_id = self.test_data["qTest TC ID"][test_case_index]

        if pd.isnull(qtest_id) or qtest_id == "":
            return None
        return int(qtest_id)

    def get_qtest_cycle_id(self, test_case_id):
        test_case_index = self.get_test_case_index(test_case_id)
        cycle_id = self.test_data["Test Cycle ID"][test_case_index]

        if pd.isnull(cycle_id) or cycle_id == "":
            return None
        return int(cycle_id)

    def validate_allowed_fields(self):
        all_fields = self.test_case_details.get_all_fields()
        unallowed_fields = []

        for key in self.test_data.keys():
            if key not in all_fields and key not in self.allowed_columns:
                unallowed_fields.append(str(key))
        if len(unallowed_fields) == 0:
            return

        unallowed_fields = ", ".join(unallowed_fields)
        print (f"[ERR] Fields not allowed in current Project: {unallowed_fields}")
        sys.exit(-1)

    def validate_mandatory_fields(self):
        mandatory_fields = self.test_case_details.get_mandatory_fields()

        for key in self.allowed_columns:
            if key not in self.test_data.keys():
                print (f"[WARN] '{key}' should be in Excel File Headers")
        missing_fields = []

        for field in mandatory_fields:
            if field not in self.test_data.keys():
                missing_fields.append(field)

        if len(missing_fields) == 0:
            return

        missing_fields = ", ".join(missing_fields)
        print (f"[ERR] Mandatory Fields not found in excel: {missing_fields}")
        sys.exit(-1)

    def get_test_case_index(self, test_case_id):
        for index in self.test_data["Test Case ID"]:
            if str(test_case_id).replace("\n", "") == str(self.test_data["Test Case ID"][index]).replace("\n", ""):
                return index
        print (f"[ERR] {test_case_id} is not found in Test Data Sheet")
        return -1

    def set_test_property_by_test_cases_name(self, test_case_id):
        test_case_index = 0
        for index, key in enumerate(self.test_data.keys()):
            if index == 0:
                value = test_case_id
            else:
                value = str(self.test_data[key][0])

            if key in self.allowed_columns:
                continue
            if key == "Type":
                if value != "Automation" and userConfig.default_automation_report:
                    print (f'[WARN] Using Automated Test Runs but using a value of {value}')

            if self.test_case_details.fields_has_array_values(key):
                try:
                    value = value.split("\n")
                    value_id = []
                    for item in value:
                        id = self.test_case_details.get_allowed_value_id(key, item)
                        value_id.append(str(id))
                    value = f"[{','.join(value)}]"
                    value_id = f"[{','.join(value_id)}]"
                    self.set_test_property(key, value, value_id)
                except Exception as e:
                    print(f"[ERR] Encountered issue on '{key}' :: {e}")
                    sys.exit(-1)

                continue

            if self.test_case_details.fields_has_allowed_values(key):
                id = self.test_case_details.get_allowed_value_id(key, value)
                self.set_test_property(key, value, id)
            else:
                self.set_test_property(key, value)

    def get_test_property_by_name(self, name):
        for index, property in enumerate(self.testCaseStructure['properties']):
            if name == property['field_name']:
                return index
        return -1

    def set_test_property(self, name, value, id=None):
        if pd.isnull(value):
            return

        current_property = {}
        current_property["field_id"] = self.test_case_details.get_field_id_by_name(name)
        current_property["field_name"] = name
        if id is None:
            current_property["field_value"] = value
        else:
            current_property["field_value"] = str(id)
            current_property["field_value_name"] = value
        self.testCaseStructure["properties"].append(current_property)

    def print_current_tc(self):
        for index, property in enumerate(self.testCaseStructure['properties']):
            pprint (property)

    def get_test_case_structure(self):
        return self.testCaseStructure
