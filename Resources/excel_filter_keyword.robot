*** Settings ***
Library    Libraries.excel_filter_library.ExcelFilterLibrary    WITH NAME    ExcelFilter
Library    Dialogs
Library    String


*** Variables ***
${ROUND}    Round12


*** Keywords ***
Get Round From Keyboard
    ${round_name}=    Get Value From User    Enter Round name (for example, Round12):
    ${round_name}=    Strip String    ${round_name}
    Should Not Be Empty    ${round_name}    Round name must not be empty.
    ExcelFilter.Validate Round Input Folder    ${round_name}
    Set Suite Variable    \${ROUND}    ${round_name}
    Log To Console    Selected Round: ${round_name}

Run Excel Filter Automation
    [Arguments]    ${round_name}=${ROUND}

    Validate total files Input    ${round_name}
    ${output_file}=    ExcelFilter.Process All Excel Files    ${round_name}
    Validate migration round    ${round_name}
    Validate Path for Input and Output
    Validate Files name of Master
    Validate Files name of Detail
    Validate Exported Excel Exists    ${round_name}

    RETURN    ${output_file}

Validate migration round
    [Arguments]    ${round_name}=${ROUND}
    ${input_dir}=    ExcelFilter.Validate Round Input Folder    ${round_name}
    Log To Console    ${EMPTY}
    Log To Console    Selected Round: ${round_name}
    Log To Console    Input folder: ${input_dir}

Validate total files Input
    [Arguments]    ${round_name}=${ROUND}

    ${input_count}=    ExcelFilter.Get Input File Count    ${round_name}
    Should Be True    ${input_count} > 0    No Excel files found for ${round_name}
    Log To Console    ${EMPTY}
    Log To Console    Input file count: ${input_count} Files

Prepare Export For Round
    [Arguments]    ${round_name}=${ROUND}
    ExcelFilter.Process All Excel Files    ${round_name}

Validate Files name of Master
    ${summary}=    ExcelFilter.Get Last Export Summary
    Log To Console    ${EMPTY}
    Log To Console    Master count (${summary}[master_name]): ${summary}[master_count]

Validate Files name of Detail
    ${summary}=    ExcelFilter.Get Last Export Summary
    Log To Console    ${EMPTY}
    FOR    ${detail}    IN    @{summary}[detail_counts]
        Log To Console    Detail count (${detail}[name]): ${detail}[count]
    END
    Log To Console    Detail file count: ${summary}[detail_file_count]

Validate Field Name Counts By Master And Detail
    ${file_counts}=    ExcelFilter.Get Last Field Name Counts
    Log To Console    ${EMPTY}
    Log To Console    Unique Field names with Result = Not match:
    FOR    ${file_count}    IN    @{file_counts}
        Log To Console    ${file_count}[type] (${file_count}[name]): ${file_count}[count] Field names
    END

Validate Path for Input and Output
    ${summary}=    ExcelFilter.Get Last Export Summary
    Log To Console    ${EMPTY}
    Log To Console    Input folder: ${summary}[input_dir]
    Log To Console    Output file: ${summary}[output_file]

Validate Exported Excel Exists
    [Arguments]    ${round_name}=${ROUND}
    ExcelFilter.Verify Output Files Exist    ${round_name}
    Log To Console    ${EMPTY} 
    Log To Console    Exported Excel file exists

Validate Excel Comparison Results
    ExcelFilter.Validate Excel Comparison Results
    log To Console    ${EMPTY}
    Log To Console    Excel comparison results validated
