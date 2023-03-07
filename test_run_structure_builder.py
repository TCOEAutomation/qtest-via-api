import pandas as pd
from pprint import pprint
import sys
import json
import userConfig
import systemConfig

import json,requests, systemConfig,importlib

class TestRunStructureBuilder():
    def __init__(self):
        self.test_run_structure = {}
        self.field_with_options = {}
        self.test_run_details = {}

    def set_payload(self, payload):
        self.test_run_structure = payload
        self.test_run_details = json.loads(systemConfig.testRunFields)
        self.set_fields_with_allowed_values()
        self.set_mandatory_properties()

    def reset_test_case_structure(self):
        self.test_run_structure = self.loaded_payload

    def get_property_by_name(self, name):
        for index, property in enumerate(self.test_run_structure['properties']):
            if name == property['field_name']:
                return index
        return -1

    def set_fields_with_allowed_values(self):
        for field in self.test_run_details:
            if "allowed_values" in field.keys():
                current_dict = {}
                for i in range(len(field['allowed_values'])):
                    current_dict[field["allowed_values"][i]["label"]] = field["allowed_values"][i]["value"]
                self.field_with_options[field["label"]] = current_dict

    def set_mandatory_properties(self):
        for field in self.test_run_details:
            if field["required"] and field["label"] in userConfig.testRunMandatory:
                current_dict_field = {}
                current_dict_field["field_id"] = int(field["id"])
                current_dict_field["field_name"] = field["label"]
                current_dict_field["field_value"] = ""
                if "allowed_values" in field.keys():
                    current_dict_field["field_value_name"] = ""
                self.test_run_structure["properties"].append(current_dict_field)

    def set_value(self, label, value):
        value_id = None
        index = self.get_property_by_name(label)
        if label in self.field_with_options.keys():
            for key in self.field_with_options[label]:
                if key == value:
                    value_id = self.field_with_options[label][key]
                    break
            self.test_run_structure["properties"][index]["field_value"] = value_id
            self.test_run_structure["properties"][index]["field_value_name"] = value
        else:
            self.test_run_structure["properties"][index]["field_value"] = value

    def get_test_run_structure(self):
        return self.test_run_structure
