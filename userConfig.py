import os,datetime
# Values to Change when using manual run test cases
default_automation_report = False
excelFileName="ManualTestCases.xlsx"
ignore_sheets=["Config", "REF", "Test Case Template Guidelines", "UPLOAD FIELD MAPPING & POLICY", "Revision History"]

# Values to Change when doing automated run test cases
# report_time_format="%Y-%m-%d %H:%M:%S"
report_time_format="%Y/%m/%d %H:%M:%S"
uploadInParentModule = False

excelAbsName=""
sofficePath="C:/Program Files/LibreOffice/program/soffice.exe"
username=""

#test runs will be created inside this
projectId=''
authHeader=''
projectName=f''

parentIdForModuleCreationForTestCases=None

parentTypeForTestCycle='test-cycle'
parentIdForTestCycle=None

uploadAttachments=True
attachmentFile=None

apiStructureConfig='structureGeneric'

targetReleaseIdForTestCycleCreationInTestExecution=None

plannedExecutionStartDate="2022-07-12T00:00:00+00:00"
plannedExecutionEndDate="2022-07-12T00:00:00+00:00"
actualExecutionStartDate="2022-07-12T00:00:00+00:00"
actualExecutionEndDate="2022-07-12T00:00:00+00:00"

attachmentList=[]

reuseExistingTestCases = False

allowed_columns = ["qTest TC ID", "Test Cycle ID", "Test Steps Description", "Test Steps Expected Result", "Test Steps Actual Result"]
testRunMandatory = ["Planned Start Date", "Planned End Date", "Status", "Execution Type"]
nan_values = ['-1.#IND', '1.#QNAN', '1.#IND', '-1.#QNAN', '#N/A N/A', '#N/A', 'n/a', 'NA', '', '#NA', 'NULL', 'null', 'NaN', '-NaN', 'nan', '-nan', '']

uploadTestSteps = False
timezone='+08:00'

retry_count = 5
retry_interval = 45 #seconds
test_execution_type="Functional"