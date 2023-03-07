import json,requests, systemConfig,importlib
# apiBaseStructures = importlib.import_module(f'{userConfig.apiStructureConfig}')
# import f'{userConfig.apiStructureConfig}' as  apiBaseStructures
import userConfig
import logging
from logging.handlers import RotatingFileHandler
import time
import sys
from datetime import datetime
from pprint import pprint
from test_case_structure_builder import TestCaseStructureBuilder
from test_run_structure_builder import TestRunStructureBuilder
from excel_reader import ExcelReader

handler = RotatingFileHandler(f'logs/qTestAPIs_{time.time()}.log', maxBytes=1048576, backupCount=10000)
logger = logging.getLogger('restInterface')
logger.setLevel(logging.DEBUG)
logger.addHandler(handler)

handlerMap = RotatingFileHandler(f'mapping/map_{time.time()}.log', maxBytes=1048576, backupCount=10000)
loggerMap = logging.getLogger('trace')
loggerMap.setLevel(logging.DEBUG)
loggerMap.addHandler(handlerMap)

# dictTestRunCreationStructure=json.loads(apiBaseStructures.testRunCreationStructure)
# dictCreateModuleStructure=json.loads(apiBaseStructures.createModuleStructure)
# dictCreateTestCycle=json.loads(apiBaseStructures.createTestCycleStructure)

def send_request_to_qtest(type, url, headers, payload):
    retries = 0
    while retries != userConfig.retry_count:
        try:
            response = requests.request(type, url, headers=headers, data=payload)
            logger.debug('\n\n**************** [Request] ****************\n')
            logger.debug(f'Url : {url}')
            logger.debug(f'Payload : {payload}')
            return response
        except Exception as e:
            print((f'[WARN] Connection with qTest Server failed. Will retry after {userConfig.retry_interval} seconds.'))
            time.sleep(userConfig.retry_interval)
            print(e)
            retries += 1
    print(f'[ERR] Retry Count {userConfig.retry_count} Exceeded.')

def createTestSuite(parentId):
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-suites"

    payload = json.dumps({
        "parentId": f'{parentId}',
        "name": f"{userConfig.projectName}"
    })

    response = send_request_to_qtest("POST", url, qtest_headers, payload)
    logger.debug('\n\n**************** [Response] ****************\n')
    print((f'[createTestSuite] : Status Code : {response.status_code}'))
    logger.debug(f'[createTestSuite] : Status Code : {response.status_code}')
    logger.debug({response.text})
    response=json.loads(response.text)
    return str(response['id'])


def getTestSteps(responseBody):
    responseBody=json.loads(responseBody)
    allTestSteps=responseBody['test_steps']
    testCaseId=responseBody['id']
    allStepIds=[]
    for testStep in allTestSteps:
        allStepIds.append(testStep['id'])
    return (testCaseId,allStepIds)


def getTestStepsWithId(responseBody):
    responseBody=json.loads(responseBody)
    testSteps=responseBody['test_steps']
    testCaseId=responseBody['id']
    allStepIds=[]
    dictTestStepIdDesc={}

    for testStep in testSteps:
        stepId=testStep['id']
        # dictTestStepIdDesc[stepId]['stepDesc']=testStep
        dictTestStepIdDesc[stepId]={}
        dictTestStepIdDesc[stepId]['status']='TBD'
        dictTestStepIdDesc[stepId]['description']=testStep['description']
        dictTestStepIdDesc[stepId]['expected']=testStep['expected']
        dictTestStepIdDesc[stepId]['actual']='TBD'

        allStepIds.append(testStep['id'])
    return (allStepIds,dictTestStepIdDesc)

def createTestRun(requestBody):
    url=f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-runs"
    #payload=json.dumps(requestBody)
    payload=json.dumps(requestBody,indent=4, sort_keys=True)

    payload = payload.replace("<", "")
    payload = payload.replace(">", "")
    response = send_request_to_qtest("POST", url, qtest_headers, payload)
    logger.debug('\n\n**************** [Response] ****************\n')
    print((f'[createTestRun] : Status Code : {response.status_code}'))
    logger.debug(f'[createTestRun] : Status Code : {response.status_code}')
    logger.debug({response.text})
    response=json.loads(response.text)
    try:
        loggerMap.info(f"qTest TR# : {response['id']}\n")
    except:
        loggerMap.info('qTest TR# : failed to create, please check logs folder\n')
    return str(response['id'])

def createModuleUnderTestDesign(requestBody):
    url=f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/modules"
    payload=json.dumps(requestBody,indent=4, sort_keys=True)

    response = send_request_to_qtest("POST", url, qtest_headers, payload)
    logger.debug('\n\n**************** [Response] ****************\n')

    print((f'[createModule] : Status Code : {response.status_code}'))
    logger.debug(f'[createModule] : Status Code : {response.status_code}')
    logger.debug({response.text})
    response=json.loads(response.text)
    return str(response['id'])

def createTestCycle(requestBody,parentId):
    url=f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-cycles?parentId={parentId}&parentType=test-cycle"
    payload=json.dumps(requestBody,indent=4, sort_keys=True)

    response = send_request_to_qtest("POST", url, qtest_headers, payload)
    logger.debug('\n\n**************** [Response] ****************\n')
    print((f'[createTestCycle] : Status Code : {response.status_code}'))
    logger.debug(f'[createTestCycle] : Status Code : {response.status_code}')
    logger.debug({response.text})
    response=json.loads(response.text)
    return str(response['id'])


def getReusableTestCaseAndSteps(reusableTestCaseId):
    """
        1. Read all TCs in a list
        2. Read from list and then remove that element
        3. Get Test Case, extract test steps and return those
    """
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-cases/{reusableTestCaseId}"
    payload = ""

    response = send_request_to_qtest("GET", url, qtest_headers, payload)
    responseText=json.loads(response.text)

    # print(response.text)
    logger.debug('[getUploadedTestCaseAndSteps]')
    logger.debug('\n\n**************** [Response] ****************\n')
    print(f'[getUploadedTestCaseAndSteps] Status-Code : {response.status_code}')
    logger.debug(f'[getUploadedTestCaseAndSteps] Status-Code : {response.status_code}')
    logger.debug(response.text)
    if str(response.status_code)=='200' or str(response.status_code)=='201':
        loggerMap.info(f"\nReport : {responseText['name']} \nqTest Test Case #: {responseText['id']}\nqTest TC Web URL : {responseText['web_url']}")
        return getTestStepsWithId(response.text)
    else:
        print('[ERROR] Upload Test Case failed, please check logs')
        logger.error('[ERROR] Upload Test Case failed, please check logs')

def getTestCaseFields():
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/settings/test-cases/fields"

    payload = {}
    response = send_request_to_qtest("GET", url, qtest_headers, payload)
    responseText=json.loads(response.text)
    logger.debug('[getTestCaseFields]')
    logger.debug('\n\n**************** [Response] ****************\n')
    print(f'[getTestCaseFields] Status-Code : {response.status_code}')
    logger.debug(f'[getTestCaseFields] Status-Code : {response.status_code}')

    if str(response.status_code)=='200' or str(response.status_code)=='201':
        systemConfig.testCaseFields = response.text
    else:
        print('[ERROR] Get Test Case Fields Failed, please check logs')
        logger.error('[ERROR] Get Test Case Fields Failed, please check logs')

def getTestRunFields():
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/settings/test-runs/fields"

    payload = {}
    response = send_request_to_qtest("GET", url, qtest_headers, payload)
    responseText=json.loads(response.text)
    logger.debug('[getTestCaseFields]')
    logger.debug('\n\n**************** [Response] ****************\n')
    print(f'[getTestCaseFields] Status-Code : {response.status_code}')
    logger.debug(f'[getTestCaseFields] Status-Code : {response.status_code}')

    if str(response.status_code)=='200' or str(response.status_code)=='201':
        systemConfig.testRunFields = response.text
        # return getTestSteps(response.text)
    else:
        print('[ERROR] Get Test Run Fields Failned, please check logs')
        logger.error('[ERROR] Get Test Run Fields Failned, please check logs')

def uploadTestCaseAndSteps(requestBody):
    getTestCaseFields()
    test_case_builder.set_payload(requestBody)
    test_case_builder.validate_excel_file()
    test_case_builder.set_test_property_by_test_cases_name(requestBody["name"])
    requestBody = test_case_builder.get_test_case_structure()
    payload=json.dumps(requestBody,indent=4, sort_keys=True)
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-cases"

    payload = payload.replace("<", "")
    payload = payload.replace(">", "")
    response = send_request_to_qtest("POST", url, qtest_headers, payload)
    responseText=json.loads(response.text.encode("utf-8"))

    logger.debug('[uploadTestCaseAndSteps]')
    logger.debug('\n\n**************** [Response] ****************\n')
    print(f'[uploadTestCaseAndSteps] Status-Code : {response.status_code}')
    logger.debug(f'[uploadTestCaseAndSteps] Status-Code : {response.status_code}')
    logger.debug((response.text).encode("utf-8"))
    loggerMap.info(f"\nReport : {requestBody['name']} \nqTest Test Case #: {responseText['id']}\nqTest TC Web URL : {responseText['web_url']}")

    if str(response.status_code)=='200' or str(response.status_code)=='201':
        #then proceed to get test steps
        return getTestSteps((response.text).encode("utf-8"))
    else:
        print('[ERROR] Upload Test Case failed, please check logs')
        logger.error('[ERROR] Upload Test Case failed, please check logs')

def doTestRunUsingSubmitTestLog(requestBody,testRunId):
    url = f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-runs/{testRunId}/test-logs"
    payload=json.dumps(requestBody,indent=4, sort_keys=True)

    logger.debug(f'Request [doTestRunUsingSubmitTestLog] Payload : {payload}')

    response = send_request_to_qtest("POST", url, qtest_headers, payload)

    logger.debug(f'[doTestRunUsingSubmitTestLog]')
    logger.debug('\n\n**************** [Response] ****************\n')

    print(f'[doTestRunUsingSubmitTestLog] Status Code : {response.status_code}')
    logger.debug(f'[doTestRunUsingSubmitTestLog] Status Code : {response.status_code}')
    logger.debug((response.text).encode("utf-8"))

def approve_test_case(test_case_id):
    url=f"https://globeph.qtestnet.com/api/v3/projects/{userConfig.projectId}/test-cases/{test_case_id}/approve"
    # url = "https://apitryout.qtestnet.com/api/v3/projects/123/test-cases/456/approve"

    payload={}

    logger.debug(f'Request [approve_test_case] Payload : {payload}')

    response = send_request_to_qtest("PUT", url, qtest_headers, payload)

    # print(response.text)
    # print(response.status_code)

    logger.debug(f'[approve_test_case]')
    logger.debug('\n\n**************** [Response] ****************\n')


    print(f'[approve_test_case] Status Code : {response.status_code}')
    logger.debug(f'[approve_test_case] Status Code : {response.status_code}')
    logger.debug((response.text).encode("utf-8"))


def create_test_case_and_test_runs():
    if not userConfig.uploadInParentModule:
        moduleId = createModuleUnderTestDesign(dictCreateModuleStructure)
    else:
        moduleId = userConfig.parentIdForModuleCreationForTestCases

    reuseTestCaseListIndex=-1

    currentTestCase = 1
    systemConfig.created_suite = {}
    for key in resultsInfoMainDict:
        #main login for pushing to qTsest goes here
        testCaseId=key
        print(f'************ [ {testCaseId} ] ************')
        logger.debug(f'************ [ {testCaseId} ] ************')
        testCaseDetails=resultsInfoMainDict[testCaseId]
        testCaseStatus=testCaseDetails['status']
        dictTestCaseStructure['name']=testCaseId
        dictTestCaseStructure['parent_id']=moduleId
        dictTestCaseStructure['description']=testCaseDetails['description']
        stepsList=testCaseDetails['steps']

        dictTestCaseStructure['test_steps']=[]
        testRunCreationStepsList=[]
        for eachStep in stepsList:

            dictTestStepStructure['description']=eachStep['description']
            dictTestStepStructure['expected']=eachStep['expected']
            dictTestStepStructure['plain_value_text']=eachStep['description']


            dictTestCaseStructure['test_steps'].append(dictTestStepStructure.copy())

            desc=eachStep['description']
            expected=eachStep['expected']
            #this will be used in test-run-creation
            testRunCreationStepsList.append({'description': desc, 'expected': expected})

        uploadedTestCaseId = None
        if not userConfig.default_automation_report:
            uploadedTestCaseId = test_case_builder.get_qtest_id(testCaseId)

        if uploadedTestCaseId is not None:
            userConfig.reuseExistingTestCases = True
            (stepIds,dictStepIdTestStep)=getReusableTestCaseAndSteps(uploadedTestCaseId)
        else:
            userConfig.reuseExistingTestCases = False
            (uploadedTestCaseId,stepIds)=uploadTestCaseAndSteps(dictTestCaseStructure)
            approve_test_case(uploadedTestCaseId)
        #At this moment, we have all the info for the first test case with all its steps, so lets run it

        #########################################################################################
        #Create test run inside test cycle
        #########################################################################################
        if userConfig.default_automation_report:
            if not systemConfig.suiteCreated:
                cycleId=createTestCycle(dictCreateTestCycle,userConfig.parentIdForTestCycle)
                # suiteId=createTestSuite(cycleId)
                systemConfig.suiteCreated=True
        else:
            cycle_id = test_case_builder.get_qtest_cycle_id(testCaseId)
            if cycle_id is None or cycleId == "":
                cycle_id = userConfig.parentIdForTestCycle

            if cycle_id not in systemConfig.created_suite.keys():
                cycleId=createTestCycle(dictCreateTestCycle, cycle_id)
                systemConfig.created_suite[cycle_id] = cycleId
            else:
                cycleId = systemConfig.created_suite[cycle_id]

        dictTestRunCreationStructure['parentId']=f'{cycleId}'
        dictTestRunCreationStructure['name']=testCaseId
        dictTestRunCreationStructure['test_case']['id']=uploadedTestCaseId
        dictTestRunCreationStructure['test_case']['name']=testCaseId
        dictTestRunCreationStructure['test_case']['test_steps']=testRunCreationStepsList
        dictTestRunCreationStructure['properties'] = []
        getTestRunFields()
        test_run_builder.set_payload(dictTestRunCreationStructure)
        test_run_builder.set_value("Planned Start Date", userConfig.plannedExecutionStartDate)
        test_run_builder.set_value("Planned End Date", userConfig.plannedExecutionEndDate)
        test_run_builder.set_value("Execution Type", userConfig.test_execution_type)

        if "pass" in str(testCaseStatus).lower():
            test_run_builder.set_value("Status", "Passed")
        else:
            test_run_builder.set_value("Status", "Failed")
        testRunCreationStructure = test_run_builder.get_test_run_structure()
        testRunId=createTestRun(testRunCreationStructure)

        import base64
        #########################################################################################
        #Submit Test log
        #########################################################################################
        dictTestRunStructure["attachments"] = []
        attachmentList = []
        if userConfig.uploadTestSteps:
            if len(attachmentList) == 0:
                dictTestRunStructure["attachments"]=[]
            else:
                print(f"Attachment list : {attachmentList}")
                for attachment in attachmentList:
                    attachmentDict={}
                    attachmentDict["name"]=f'{attachment[0]}.png'
                    attachmentDict["content_type"]="application/octet-stream"
                    attachmentFile=attachment[1]

                    file=open(attachmentFile,"rb")
                    fileContents=file.read()
                    fileContents=base64.encodebytes(fileContents)

                    fileContents=fileContents.decode("utf-8")
                    attachmentDict["data"]=fileContents

                    dictTestRunStructure["attachments"].append(attachmentDict)

        if userConfig.default_automation_report:
            attachmentDict = {}
            attachmentDict["name"]=f'{testCaseId}.docx'
            attachmentDict["content_type"]="application/octet-stream"
            attachmentFile=f"TestCaseWise\\TC_{currentTestCase}.docx"

            file=open(attachmentFile,"rb")
            fileContents=file.read()
            fileContents=base64.encodebytes(fileContents)

            fileContents=fileContents.decode("utf-8")
            attachmentDict["data"]=fileContents

            dictTestRunStructure["attachments"].append(attachmentDict)
        currentTestCase += 1

        dictTestRunStructure['submittedBy']=userConfig.username

        #same as test case name
        dictTestRunStructure['name']=testCaseId
        if not userConfig.default_automation_report:
            start_time = userConfig.plannedExecutionStartDate
            end_time = userConfig.plannedExecutionEndDate
        else:
            start_time = datetime.strptime(testCaseDetails['startTime'], '%Y/%m/%d %H:%M:%S')
            end_time = datetime.strptime(testCaseDetails['endTime'], '%Y/%m/%d %H:%M:%S')
            start_time = f"{start_time.strftime('%Y-%m-%dT%H:%M:%S')}{userConfig.timezone}"
            end_time = f"{end_time.strftime('%Y-%m-%dT%H:%M:%S')}{userConfig.timezone}"

        dictTestRunStructure['exe_start_date']=start_time
        dictTestRunStructure['exe_end_date']=end_time
        dictTestRunStructure['test_step_logs']=[]
        # for index, (key, val) in enumerate(dictStepIdTestStep.items()):
        # for index, val in enumerate(stepsList):
        indexForTestDescComparison=-1

        if userConfig.reuseExistingTestCases:
            for key in dictStepIdTestStep:
                indexForTestDescComparison+=1
                dictTestRunStepsStructure=apiBaseStructures.testRunStepsStructure
                dictTestRunStepsStructure['test_step_id']=key
                should_continue = True
                try:
                    tmpIndex = testRunCreationStepsList[indexForTestDescComparison]['description']
                except:
                    should_continue = False

                if should_continue and testRunCreationStepsList[indexForTestDescComparison]['description']==dictStepIdTestStep[key]['description']:
                    #update logic TBD - if step is same as in report, get status from report, else default is TC status
                    # testStepInfoFromJsonReport=testRunCreationStepsList[indexForTestDescComparison]
                    dictStepIdTestStep[key]['status']=stepsList[indexForTestDescComparison]['status']
                    dictStepIdTestStep[key]['actual']=stepsList[indexForTestDescComparison]['actual']

                    dictTestRunStepsStructure['status']['name']=dictStepIdTestStep[key]['status']
                    testStepStatus=dictTestRunStepsStructure['status']['name']
                    if "pass" in str(testStepStatus).lower():
                        dictTestRunStepsStructure['status']['id']='601'
                        dictTestRunStepsStructure['status']['name']='Passed'
                    else:
                        dictTestRunStepsStructure['status']['id']='602'
                        dictTestRunStepsStructure['status']['name']='Failed'

                    # dictTestRunStepsStructure['description']=dictStepIdTestStep[key]['description']
                    # dictTestRunStepsStructure['expected_result']=dictStepIdTestStep[key]['expected']
                    # dictTestRunStepsStructure['actual_result']=dictStepIdTestStep[key]['actual']

                else:
                    #ukate logic TBD - if step is same as in report, get status from report, else default is TC status
                    dictStepIdTestStep[key]['status']=testCaseStatus
                    if "pass" in str(dictStepIdTestStep[key]['status']).lower():
                        dictStepIdTestStep[key]['actual']='Same as expected'
                    else:
                        dictStepIdTestStep[key]['actual']='Different than expected (Check attachment for details)'
                    dictTestRunStepsStructure['status']['name']=dictStepIdTestStep[key]['status']
                    testStepStatus=dictTestRunStepsStructure['status']['name']
                    if "pass" in str(testStepStatus).lower():
                        dictTestRunStepsStructure['status']['id']='601'
                        dictTestRunStepsStructure['status']['name']='Passed'
                    else:
                        dictTestRunStepsStructure['status']['id']='602'
                        dictTestRunStepsStructure['status']['name']='Failed'
                dictTestRunStepsStructure['description']=dictStepIdTestStep[key]['description']
                dictTestRunStepsStructure['expected_result']=dictStepIdTestStep[key]['expected']
                dictTestRunStepsStructure['actual_result']=dictStepIdTestStep[key]['actual']

                if "pass" in str(testCaseStatus).lower():
                    dictTestRunStructure['status']['id']='601'
                    dictTestRunStructure['status']['name']='Passed'
                else:
                    dictTestRunStructure['status']['id']='602'
                    dictTestRunStructure['status']['name']='Failed'
                dictTestRunStructure['test_step_logs'].append(dictTestRunStepsStructure.copy())

        else:
            for index, val in enumerate(stepsList):
                dictTestRunStepsStructure=apiBaseStructures.testRunStepsStructure
                dictTestRunStepsStructure['test_step_id']=stepIds[index]
                dictTestRunStepsStructure['status']['name']=stepsList[index]['status']
                # dictTestRunStepsStructure['status']['color']="#000000"
                # if "pass" in str(stepsList[index]['status']).lower():
                #     dictTestRunStepsStructure['status']['color']="#0cdda8"
                testStepStatus=dictTestRunStepsStructure['status']['name']
                if "pass" in str(testStepStatus).lower():
                    dictTestRunStepsStructure['status']['id']='601'
                    dictTestRunStepsStructure['status']['name']='Passed'
                else:
                    dictTestRunStepsStructure['status']['id']='602'
                    dictTestRunStepsStructure['status']['name']='Failed'

                dictTestRunStepsStructure['description']=stepsList[index]['description']
                dictTestRunStepsStructure['expected_result']=stepsList[index]['expected']
                dictTestRunStepsStructure['actual_result']=stepsList[index]['actual']

                if "pass" in str(testCaseStatus).lower():
                    dictTestRunStructure['status']['id']='601'
                    dictTestRunStructure['status']['name']='Passed'
                else:
                    dictTestRunStructure['status']['id']='602'
                    dictTestRunStructure['status']['name']='Failed'
                dictTestRunStructure['test_step_logs'].append(dictTestRunStepsStructure.copy())

        doTestRunUsingSubmitTestLog(dictTestRunStructure,testRunId)

        # print(dictTestCaseStructure)
        # return dictTestCaseStructure

def main():
    global apiBaseStructures,dictTestCaseStructure, dictTestStepStructure, dictTestRunStructure, dictTestRunStepsStructure, dictTestRunCreationStructure, dictCreateModuleStructure, dictCreateTestCycle, test_case_builder, test_run_builder, resultsInfoMainDict, qtest_headers

    excel_reader = ExcelReader()
    excel_reader.store_data()
    qtest_headers = {
        'Authorization': f'{userConfig.authHeader}',
        'Content-Type': 'application/json'
    }
    test_case_builder = TestCaseStructureBuilder()
    test_run_builder = TestRunStructureBuilder()
    resultsInfoMainDict=json.load(open('results.json'))
    apiBaseStructures = importlib.import_module(f'{userConfig.apiStructureConfig}')
    dictTestCaseStructure=apiBaseStructures.testCaseStructure
    dictTestStepStructure=apiBaseStructures.testStepStructure
    dictTestRunStructure=apiBaseStructures.testRunStructure
    dictTestRunStepsStructure=apiBaseStructures.testRunStepsStructure
    dictTestRunCreationStructure=apiBaseStructures.testRunCreationStructure
    dictCreateModuleStructure=apiBaseStructures.createModuleStructure
    dictCreateTestCycle=apiBaseStructures.createTestCycleStructure
    create_test_case_and_test_runs()
