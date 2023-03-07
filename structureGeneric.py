import userConfig

testRunCreationStructure = {
    "parentId": f"{userConfig.parentIdForTestCycle}",
    "parentType": f'{userConfig.parentTypeForTestCycle}',
    "automation": "Yes",
    "name": "17 Jan Test Run via API",
    "properties":
    [

    ],
    "test_case":
    {
        "id": None,
        "name": "",
        "test_steps":
        [
        ]
    }
}
createTestCycleStructure={
  "name": f'{userConfig.projectName}',
  "description": f'{userConfig.projectName}',
  "target_release_id": f'{userConfig.targetReleaseIdForTestCycleCreationInTestExecution}',
  "target_build_id": 0,
  "test-cycles": [
    None
  ]
}


createModuleStructure={
  "name": f'{userConfig.projectName}',
  "parent_id": f'{userConfig.parentIdForModuleCreationForTestCases}',
  "description": f'{userConfig.projectName}',
  "shared": False,
  "children": [
    None
  ],
  "recursive": False
}


testRunStepsStructure= {
    "test_step_id": 0,
    "status": {
        "id": 0
    }
}

testRunStructure={
    "submittedBy": "",
    "id": 1,

    "exe_start_date": f"{userConfig.actualExecutionStartDate}",
    "exe_end_date": f"{userConfig.actualExecutionEndDate}",
    "note": "Note",
    "attachments": [
    ],
    "name": "test run name",
    "planned_exe_time": 0,
    "actual_exe_time": 0,
    "build_number": "10",
    "build_url": "http://localhost:8080/jenkins",
    "properties": [
    ],
    "status": {
        "id": 601,
        "name": "Passed",
        "is_default": False,
        "color": "#0cdda8",
        "active": True
    },
    "result_number": 0,
    "test_step_logs": [],
    "defects": []
}

testCaseStructure = {
    "id": 1,
    "name": "",
    "order": 1,
    "properties": [],
    "test_steps": [],
    "parent_id": 0,
    "description": "",
    "precondition": "",
    "agent_ids": []
}

testStepStructure={
    "description": "",
    "expected": "",
    "order": 1,
    "attachments": [],
    "plain_value_text": ""
}