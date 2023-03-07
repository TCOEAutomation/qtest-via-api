import pandas as pd
import os,datetime,sys
from pprint import pprint
import userConfig

class ExcelReader():
    def __init__(self):
        if userConfig.default_automation_report:
            xls = pd.ExcelFile("DataSheet.xlsx")
            self.df = xls.parse(xls.sheet_names[0], keep_default_na=False, na_values=userConfig.nan_values)
        else:
            xls = pd.ExcelFile(userConfig.excelFileName)
            self.df = xls.parse(xls.sheet_names[0], keep_default_na=False, na_values=userConfig.nan_values)
        self.test_data = self.df.to_dict(orient='index')

    def store_data(self):
        for item in self.test_data.values():
            if "Key" not in item.keys() and "Value" not in item.keys():
                print("[ERR] 'Key' and 'Value' Column Header should be in 'Config' Sheet")
                sys.exit(-1)

            if item["Key"] == "Tester Name":
                userConfig.username = item["Value"]
            if item["Key"] == "Authentication Header":
                userConfig.authHeader = item["Value"]
                if pd.isnull(item["Value"]):
                    print("[ERR] Authentication Header in excel is empty")
                    sys.exit(-1)
            if item["Key"] == "Project Prefix":
                project_name = F'{item["Value"]}_{userConfig.username}_{datetime.datetime.now().strftime("%b_%d_%Y_%H_%M_%S")}'
                userConfig.projectName =  project_name
            if item["Key"] == "Project ID":
                userConfig.projectId = item["Value"]
            if item["Key"] == "Module ID":
                userConfig.parentIdForModuleCreationForTestCases = item["Value"]
            if item["Key"] == "Test Execution ID":
                userConfig.parentIdForTestCycle = item["Value"]
                userConfig.targetReleaseIdForTestCycleCreationInTestExecution = item["Value"]
            if item["Key"] == "Planned Start Date":
                userConfig.plannedExecutionStartDate = f'{item["Value"]}T00:00:00+00:00'
            if item["Key"] == "Planned End Date":
                userConfig.plannedExecutionEndDate = f'{item["Value"]}T00:00:00+00:00'
            if item["Key"] == "Excel Filename":
                filename = item["Value"]
                if not pd.isnull(filename):
                    userConfig.excelFileName=os.path.join(os.getcwd(), filename)
if __name__ == '__main__':
    e = ExcelReader()
    e.store_data()
