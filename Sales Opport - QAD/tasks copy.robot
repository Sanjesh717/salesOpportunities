*** Settings ***
Documentation       Sales Opportinities Bot.

Library             RPA.Robocloud.Secrets
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Excel.Files
Library             Collections
Library             XML
Library             OperatingSystem
Library             RPA.Desktop.OperatingSystem
Library             String


*** Tasks ***
Sales Opportinities Bot.
    ${configlist}=    Read Config file
    Open Workbook    ${configlist}[0]    # read input excel file
    Read Worksheet    ${configlist}[1]    # read input excel sheet
    ${inputexceltable}=    Read Worksheet As Table    header=${True}
    open intranet website    ${configlist}
    Log IN
    Sleep    5s
    Process each transaction    ${inputexceltable}


*** Keywords ***
Process each transaction
    [Arguments]    ${inputexceltable}
    FOR    ${inputexcelrow}    IN    @{inputexceltable}
        Log    ${inputexcelrow}
        ${input data list}=    Get Data from Input Excel File    ${inputexcelrow}
        IF    "${input data list}[0]" != "None"
            Filter the data    ${input data list}
            ${resultexist}=    Does Page Contain Element    xpath=//tr[@class="k-master-row k-state-selected"]
            IF    ${resultexist} == ${True}
                IF    "${input data list}[1]" == "Existing"
                    Edit the data    ${input data list}
                ELSE
                    Log    business exception
                END
            ELSE
                Create the data    ${input data list}
            END
        ELSE
            Log    cell is empty
        END
    END

open intranet website
    [Arguments]    ${configlist}
    Open Available Browser    ${configlist}[8]    maximized=${True}

Log IN
    ${secret}=    Get Secret    QAD
    Input Text    id:username    ${secret}[username]
    Sleep    2s
    Input Password    id:password    ${secret}[password]
    Sleep    2s
    Submit Form
    sleep    10s
    RPA.Browser.Selenium.Click Element    xpath=//span[@class="k-dropdown-wrap k-state-default"]
    Sleep    2s
    RPA.Browser.Selenium.Click Element    xpath=//li[@data-offset-index="27"]
    Sleep    2s
    RPA.Browser.Selenium.Click Element    xpath=//span[@class="fa fa-search"]
    Sleep    2s
    Input Text    //*[@id="webshellMenu_kAutoCompleteMenuSearch"]    Sales Opportunities
    Wait Until Element Is Visible    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li/a/div/div/span[2]
    RPA.Browser.Selenium.Click Element    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li/a/div/div/span[2]

Read Config file
    ${config}=    Parse Xml    config.xml
    ${config excelpath}=    Get Element Text    ${config}[0]
    ${configexcelSheetname}=    Get Element Text    ${config}[1]
    ${system name}=    Get Username
    ${input excelpath}=    Replace String    ${config excelpath}    name    ${system name}
    Open Workbook    ${input excelpath}
    Read Worksheet    ${configexcelSheetname}
    ${table}=    Read Worksheet As Table    header=${True}
    ${configlist}=    Create List
    FOR    ${row}    IN    @{table}
        Log    ${row}
        ${inputexcelpath}=    Set Variable    ${row}[Input Excel file path]
        ${inputexcelsheetname}=    Set Variable    ${row}[Input Excel Sheet Name]
        ${BE Exception mail ID}=    Set Variable    ${row}[Business Exception Mail ID]
        ${BE Exception Subject}=    Set Variable    ${row}[Business Exception Subject]
        ${BE Exception body}=    Set Variable    ${row}[Business Exception Mail Body]
        ${SE Exception mail ID}=    Set Variable    ${row}[System Exception Mail ID]
        ${SE Exception Subject}=    Set Variable    ${row}[System Exception Subject]
        ${SE Exception body}=    Set Variable    ${row}[System Exception Mail Body]
        ${QAD url}=    Set Variable    ${row}[QAD URL]
        Append To List
        ...    ${configlist}
        ...    ${inputexcelpath}
        ...    ${inputexcelsheetname}
        ...    ${BE Exception mail ID}
        ...    ${BE Exception Subject}
        ...    ${BE Exception body}
        ...    ${SE Exception mail ID}
        ...    ${SE Exception Subject}
        ...    ${SE Exception body}
        ...    ${QAD url}
    END
    RETURN    ${configlist}

Get Data from Input Excel File
    [Arguments]    ${inputexcelrow}
    ${input data list}=    Create List
    ${Opportunity Name}=    Set Variable    ${inputexcelrow}[Opportunity Name]
    ${Creation Type}=    Set Variable    ${inputexcelrow}[Creation Type]
    ${Change Stage}=    Set Variable    ${inputexcelrow}[Change Stage]
    ${TYPE}=    Set Variable    ${inputexcelrow}[TYPE]
    ${External Account Mgr}=    Set Variable    ${inputexcelrow}[External Account Mgr]
    ${Address}=    Set Variable    ${inputexcelrow}[Address]
    ${Postal Code}=    Set Variable    ${inputexcelrow}[Postal Code]
    ${City}=    Set Variable    ${inputexcelrow}[City]
    ${State}=    Set Variable    ${inputexcelrow}[State]
    ${Country}=    Set Variable    ${inputexcelrow}[Country]
    ${Email}=    Set Variable    ${inputexcelrow}[Email]
    ${Website}=    Set Variable    ${inputexcelrow}[Website]
    ${Region}=    Set Variable    ${inputexcelrow}[Region]
    Append To List
    ...    ${input data list}
    ...    ${Opportunity Name}
    ...    ${Creation Type}
    ...    ${Change Stage}
    ...    ${TYPE}
    ...    ${External Account Mgr}
    ...    ${Address}
    ...    ${Postal Code}
    ...    ${City}
    ...    ${State}
    ...    ${Country}
    ...    ${Email}
    ...    ${Website}
    ...    ${Region}
    RETURN    ${input data list}

Edit the data
    [Arguments]    ${input data list}
    Click Element    xpath=//a[@id="ToolBtnUpdate"]
    Sleep    2s
    Scroll Element Into View    xpath=//button[@id="SalesOpportunitys_ChangeStage"]
    Sleep    2s
    Click Element    xpath=//button[@id="SalesOpportunitys_ChangeStage"]
    Sleep    7s
    Select Frame    toRuleThemAll
    Click Element    xpath=//span[@class="k-input"]
    Sleep    3s
    #Click element    xpath=//ul[@class="k-list k-reset"]/li[text()='Identify Sponsor']
    #Click Element    xpath=//ul[@class="k-list k-reset"]/li[text()[contains(.,'Identify Sponsor')]]
    TRY
        Click Element    xpath=//ul[@class="k-list k-reset"]/li[text()[contains(.,'${input data list}[2]')]]
        Sleep    2s
        Click Element    //*[@id="ToolBtnSave"]
        Sleep    7s
        ${subelement}=    Does Page Contain Element    xpath=//div[@id="qModalDialog"]
        IF    ${subelement} == ${True}
            Click Element    xpath=//div[@id="qModalDialog"]//button[@id="qModalDialogConfirm"]
            Sleep    5s
        ELSE
            ${exceptsavebutton}=    Set Variable    No Subelement is present
        END
        Unselect Frame
        Sleep    2s
        Click Element    xpath=//div[@id="qViewFormPane"]//span[@class="fa fa-minus"]
        Sleep    4s
    EXCEPT
        ${exceptionMessage}=    Set Variable    Change stage cannot be found
        sleep    2s
        Click Element    xpath=//button[@id="kendoWindowChildCloseButton"]
        sleep    10s
        Unselect Frame
        ${subcloseelement}=    Does Page Contain Element
        ...    xpath=//div[@uib-modal-window="modal-window"]//div[@id="qModalDialog"]
        #${subcloseelement}=    Does Page Contain Element    xpath=//div[@class="modal-content"]
        IF    ${subcloseelement} == ${True}
            Click Element    xpath=//div[@id="qModalDialog"]//button[@id="qModalDialogConfirm"]
            Sleep    3s
            Click Element    //div[@id="qViewFormPane"]//span[@class="fa fa-minus"]
        ELSE
            ${exceptleavebutton}=    Set Variable    unable to close the element
            Sleep    3s
        END
    END

Create the data
    [Arguments]    ${input data list}
    IF    "${input data list}[0]" != "None"
        IF    "${input data list}[3]" != "None"
            IF    "${input data list}[4]" != "None"
                IF    "${input data list}[5]" != "None"
                    IF    "${input data list}[6]" != "None"
                        IF    "${input data list}[7]" != "None"
                            IF    "${input data list}[9]" != "None"
                                IF    "${input data list}[12]" != "None"
                                    Click Element    xpath=//span[@class="fa fa-search"]
                                    Sleep    2s
                                    Input Text    //*[@id="webshellMenu_kAutoCompleteMenuSearch"]    CRM Accounts
                                    Sleep    5s
                                    Click Element
                                    ...    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li/a/div/div/span[2]
                                    Sleep    5s
                                    Click Element    xpath=//a[@id="ToolBtnNew"]
                                    Sleep    2s
                                    Click Element    //*[@id="RecordTypeAutoField1_dropFocus"]/span
                                    sleep    2s
                                    Click Element
                                    ...    xpath=//ul[@id="RecordTypeAutoField1_listbox"]//li[text()='${input data list}[3]']
                                    Sleep    2s
                                    Input Text    //*[@id="NameAutoField"]    ${input data list}[0]
                                    Sleep    2s
                                    Input Text    //*[@id="OwnerAutoField"]    ${input data list}[4]
                                    Sleep    2s
                                    Scroll Element Into View    //*[@id="Address1AutoField"]
                                    sleep    2s
                                    Input Text    //*[@id="Address1AutoField"]    ${input data list}[5]
                                    Sleep    2s
                                    Input Text    //*[@id="PostCodeAutoField1"]    ${input data list}[6]
                                    Sleep    2s
                                    Input Text    //*[@id="CityAutoField"]    ${input data list}[7]
                                    Sleep    2s
                                    Input Text    //*[@id="StateAutoField"]    ${input data list}[8]
                                    Sleep    2s
                                    Input Text    //*[@id="CountryAutoField"]    ${input data list}[9]
                                    Sleep    2s
                                    Scroll Element Into View    //*[@id="EmailAutoField"]
                                    sleep    2s
                                    Input Text    //*[@id="EmailAutoField"]    ${input data list}[10]
                                    Sleep    2s
                                    Input Text    //*[@id="WebsiteAutoField"]    ${input data list}[11]
                                    Sleep    2s
                                    Scroll Element Into View    //*[@id="AccountRegionAutoField"]
                                    sleep    2s
                                    Input Text    //*[@id="AccountRegionAutoField"]    ${input data list}[12]
                                    Sleep    2s
                                    Click Element    //*[@id="ToolBtnSave"]
                                    Sleep    3s
                                    Click Element    //*[@id="btnViewFormPane"]/span
                                ELSE
                                    Log    Field missing
                                END
                            ELSE
                                Log    Field missing
                            END
                        ELSE
                            Log    Field missing
                        END
                    ELSE
                        Log    Field missing
                    END
                ELSE
                    Log    Field missing
                END
            ELSE
                Log    Field missing
            END
        ELSE
            Log    Field missing
        END
    ELSE
        Log    Field missing
    END
    #RETURN    ${element}

Filter the data
    [Arguments]    ${input data list}
    Click Element    xpath=//button[@id="btnSearchAdvance"]
    Sleep    2s
    Click Element    xpath=//span[@id="scWrap_scFieldList_1"]
    Sleep    10s
    Click Element    xpath=//ul[@id="scFieldList_{{sc.seq}}_listbox"]//li[@data-offset-index="29"]
    Sleep    2s
    Click Element    xpath=//span[@id="scWrap_scCondList_1"]
    Sleep    2s
    Click Element    //*[@id="scCondList_{{sc.seq}}_listbox"]/li[1]
    Sleep    2s
    Click Element    xpath=//input[@id="scString1_1"]
    Sleep    2s
    Input Text    xpath=//input[@id="scString1_1"]    ${input data list}[0]
    Sleep    2s
    Click Element    xpath=//button[@id="btnSaveSearchCond"]
    sleep    4s
