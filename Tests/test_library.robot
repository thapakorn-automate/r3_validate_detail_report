*** Settings ***
Library    Libraries.excel_filter_library.ExcelFilterLibrary

*** Test Cases ***
Check Library Keyword Registration
    ${count}=    Get Input File Count
    Log To Console    Input file count: ${count}