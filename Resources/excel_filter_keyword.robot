*** Settings ***
Library    Libraries.excel_filter_library.ExcelFilterLibrary    WITH NAME    ExcelFilter


*** Variables ***
${ROUND}    Round14


*** Keywords ***
Run Excel Filter Automation
    [Arguments]    ${round_name}=${ROUND}

    ${input_count}=    ExcelFilter.Get Input File Count    ${round_name}
    Should Be True    ${input_count} > 0    No Excel files found for ${round_name}

    ${output_file}=    ExcelFilter.Process All Excel Files    ${round_name}

    ExcelFilter.Verify Output Files Exist    ${round_name}

    RETURN    ${output_file}


Validate Exported Excel Exists
    [Arguments]    ${round_name}=${ROUND}
    ExcelFilter.Verify Output Files Exist    ${round_name}


Validate Excel Comparison Results
    ExcelFilter.Validate Excel Comparison Results