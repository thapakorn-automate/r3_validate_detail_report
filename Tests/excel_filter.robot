
*** Settings ***
Resource    ../Resources/excel_filter_keyword.robot
Library    Collections
Suite Setup    Get Round From Keyboard


*** Test Cases ***

# Combine Detail Files Against Master
#     [Documentation]    Combine and validate Master/Detail Excel files.
#     Run Excel Filter Automation    ${ROUND}
#     Validate Exported Excel Exists    ${ROUND}
#     Validate Excel Comparison Results

TC01 Validation round for Migration
    Validate migration round
TC02 Validation total files Input
    Validate total files Input
TC03 Validation Path for Input and Output
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Path for Input and Output
TC04 Validation Files name of Master
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Files name of Master
TC05 Validation Files name of Detail
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Files name of Detail
TC06 Validation Exported Excel Exists
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Exported Excel Exists
TC07 Validation Excel Comparison Results
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Excel Comparison Results
TC08 Validation Field Names With Not Match Result
    [Setup]    Prepare Export For Round    ${ROUND}
    Validate Field Name Counts By Master And Detail
