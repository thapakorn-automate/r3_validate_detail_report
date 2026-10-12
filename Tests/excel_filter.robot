
*** Settings ***
Resource    ../Resources/excel_filter_keyword.robot


*** Variables ***
${ROUND_NAME}    Round12


*** Test Cases ***

Combine Detail Files Against Master
    [Documentation]    Combine and validate Master/Detail Excel files.
    Run Excel Filter Automation    ${ROUND_NAME}
    Validate Exported Excel Exists    ${ROUND_NAME}
    Validate Excel Comparison Results