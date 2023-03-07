sheet_obj=None

document=None
run=None
paragraph=None
stats_paragraph=None
current_sheet = ""

LastColumn="AH"
column_testCaseId="A"
column_testStepNumber="B"
column_testDescription="C"
column_expectedResult="F"
column_actualResult="G"
column_stepStatus="H"
column_screenshot="I"
column_startTime="J"
column_endTime="K"

name_testCaseId="Test Scenario Sr No"
name_testStepNumber="Test Step No"
name_testDesc="Test Case Desc"
name_testStepDesc="Test Step Desc"
name_ExpectedResult="Expected Result"
name_ActualResult="Actual Result"
name_Status="Status"
name_StartTime="StartTime"
name_EndTime="EndTime"


hexcolorcode_pass='90EE90'
hexcolorcode_fail='ff3232'
hexcolorcode_test_case_ref='A9A9A9'
hexcolorcode_test_step_ref='ccccff'


testCaseId=None
testStepDesc=None
testStatus=None

pass_ctr=0
fail_ctr=0


htmlFileName="htmlReport.html"
testCaseWrittenOnce=False
childAccordianId=None

testCaseCurrentNumber=0
testStepCurrentNumber=0


qTestDict={}
qTestSubDict={}

suiteCreated=False
attachmentUploaded=False

breaktestDescAfterChars=50
breakExpectedResultAfterChars=50
breakActualResultAfterChars=50

htmlFileName="htmlReport.html"
testCaseWrittenOnce=False
childAccordianId=None

junitFileName="junit.xml"

test_cases=[]
ts=[]
tcId=None
testDesc=None
currentTestCaseNumber=1
testCaseWiseReport=True
attachmentsToBePassedToQtest=[]

testCaseFields={}
testRunFields={}