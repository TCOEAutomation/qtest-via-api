import json
import sys
from pprint import pprint
import importlib
import systemConfig
import pandas as pd

class TestCaseDetails():
    def __init__(self):
        self.all_fields = []
        self.mandatory_fields = []
        self.apiBaseStructures = {}

    def setup(self):
        self.apiBaseStructures = json.loads(systemConfig.testCaseFields)
        self.clean_duplicates()
        self.set_all_fields()

    def clean_duplicates(self):
        for index, field in enumerate(self.apiBaseStructures):
            if "allowed_values" in self.apiBaseStructures[index]:
                allowed_values = self.apiBaseStructures[index]["allowed_values"]
                all = self.apiBaseStructures[index]["allowed_values"]
                df = pd.DataFrame(all)
                df.drop_duplicates(subset=['label'], keep='last', inplace=True)
                all = df.to_dict("records")
                self.apiBaseStructures[index]["allowed_values"] = all

    def get_test_field_index_by_name(self, name):
        for index, field in enumerate(self.apiBaseStructures):
            if field["label"] == name:
                return index
        return -1

    def fields_has_array_values(self, name):
        index = self.get_test_field_index_by_name(name)
        if self.apiBaseStructures[index]["attribute_type"] == "ArrayNumber":
            return True
        return False

    def fields_has_allowed_values(self, name):
        index = self.get_test_field_index_by_name(name)
        if "allowed_values" in self.apiBaseStructures[index].keys():
            return True
        return False

    def get_field_id_by_name(self, name):
        for field in self.apiBaseStructures:
            if field["label"] == name:
                return field["id"]

    def get_allowed_values_by_field(self, name):
        index = self.get_test_field_index_by_name(name)
        return self.apiBaseStructures[index]["allowed_values"]

    def get_allowed_value_id(self, name, value):
        allowed_values = self.get_allowed_values_by_field(name)
        allowed_values_list = []
        for allowed_value in allowed_values:
            allowed_values_list.append(allowed_value["label"])
            if allowed_value["label"] == value:
                return allowed_value["value"]
        print (f"[ERR] '{value}' is not allowed for '{name}'. Allowed values are: {','.join(allowed_values_list)}")
        sys.exit(-1)

    def set_all_fields(self):
        for field in self.apiBaseStructures:
            self.all_fields.append(field["label"])
            if field["required"]:
                self.mandatory_fields.append(field["label"])

    def print_details(self):
        pprint(self.apiBaseStructures)

    def get_all_fields(self):
        return self.all_fields

    def get_mandatory_fields(self):
        return self.mandatory_fields
