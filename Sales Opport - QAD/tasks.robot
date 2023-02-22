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
Library             RPA.Outlook.Application

Suite Teardown      RPA.Outlook.Application.Quit Application
Task Setup          RPA.Outlook.Application.Open Application


*** Tasks ***
Sales Opportinities Bot.
    TRY
        ${xmllist}=    Read xml config file
        TRY
            ${configlist}=    Read Config file    ${xmllist}
            TRY
                Open Workbook    ${configlist}[0]    # read input excel file
                Read Worksheet    ${configlist}[1]    # read input excel sheet
                ${inputexceltable}=    Read Worksheet As Table    header=${True}
                TRY
                    open intranet website    ${configlist}
                    Log IN
                    Sleep    5s
                    Process each transaction    ${inputexceltable}    ${configlist}
                    Sleep    4s
                    Close All Browsers
                EXCEPT
                    System exception mail    ${configlist}
                END
            EXCEPT
                Input File Exception    ${configlist}
            END
        EXCEPT
            Config exception    ${xmllist}
        END
    EXCEPT
        Log    Unable to Read Config xml file.
    END


*** Keywords ***
Process each transaction
    [Arguments]    ${inputexceltable}    ${configlist}
    TRY
        FOR    ${inputexcelrow}    IN    @{inputexceltable}
            Log    ${inputexcelrow}
            ${input data list}=    Get Data from Input Excel File    ${inputexcelrow}
            IF    "${input data list}[0]" != "None"
                Filter the data    ${input data list}
                ${resultexist}=    Does Page Contain Element    xpath=//tr[@class="k-master-row k-state-selected"]
                IF    ${resultexist} == ${True}
                    IF    "${input data list}[1]" == "Existing"
                        Edit the data    ${configlist}    ${input data list}
                    ELSE
                        Log    business exception
                        Business exception mail    ${configlist}    ${input data list}
                    END
                ELSE
                    IF    "${input data list}[1]" == "Existing"
                        Business Exception subject if new marked as Existing    ${configlist}    ${input data list}
                    ELSE
                        Click Element    xpath=//a[@id="ToolBtnNew"]
                        sleep    8s
                        Click Element
                        ...    xpath=//button[@id="AccountNameAutoField_lookup"]//span[@class="fa fa-search"]
                        sleep    2s
                        Select Frame    lookUpModalIframe
                        sleep    3s
                        Input Text
                        ...    //div[@class="qBrowseSearchToolbar"]//input[@placeholder="Account Name starts with"]
                        ...    ${input data list}[0]
                        Sleep    3s
                        Press Keys
                        ...    //div[@class="qBrowseSearchToolbar"]//input[@placeholder="Account Name starts with"]
                        ...    ENTER
                        Sleep    7s
                        ${rowpresent}=    Does Page Contain Element
                        ...    xpath=//tr[@class="k-master-row k-state-selected"]
                        IF    ${rowpresent} == ${True}
                            Sleep    4s
                            Unselect Frame
                            Click Element    xpath=//button[@ng-click="ok()"]
                            Sleep    7s
                            Click Element    xpath=//button[@id="ToolBtnSave"]
                            sleep    4s
                            Click Element    xpath=//button[@id="btnViewFormPane"]//span[@class="fa fa-minus"]
                        ELSE
                            Unselect Frame
                            Click Element    xpath=//button[@id="lookUpCancelBtn"]
                            sleep    7s
                            Create the data    ${input data list}
                        END
                    END

                    #Create the data    ${input data list}
                END
            ELSE
                Log    cell is empty
            END
        END
    EXCEPT
        Input File Exception    ${configlist}
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

Read xml config file
    ${config}=    Parse Xml    config.xml
    ${xmllist}=    Create List
    ${config excelpath}=    Get Element Text    ${config}[0]
    ${configexcelSheetname}=    Get Element Text    ${config}[1]
    ${configexcptionmailid}=    Get Element Text    ${config}[2]
    ${configexcptionmailsubject}=    Get Element Text    ${config}[3]
    ${configexcptionmailbody}=    Get Element Text    ${config}[4]
    ${system name}=    Get Username
    ${input excelpath}=    Replace String    ${config excelpath}    name    ${system name}
    Append To List
    ...    ${xmllist}
    ...    ${input excelpath}
    ...    ${configexcelSheetname}
    ...    ${configexcptionmailid}
    ...    ${configexcptionmailsubject}
    ...    ${configexcptionmailbody}
    RETURN    ${xmllist}

Read Config file
    [Arguments]    ${xmllist}
    Open Workbook    ${xmllist}[0]
    Read Worksheet    ${xmllist}[1]
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
        ${inputfileexception}=    Set Variable    ${row}[InputFileExcption]
        ${inputfileexceptionsub}=    Set Variable    ${row}[Input Excel missing Subject]
        ${businesexception}=    Set Variable    ${row}[Business Exception subject if new]
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
        ...    ${inputfileexception}
        ...    ${inputfileexceptionsub}
        ...    ${businesexception}
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
    [Arguments]    ${configlist}    ${input data list}
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
        #Business exception mail    ${configlist}    ${input data list}
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
                                    Sleep    5s
                                    Click Element    //span[@class="fa fa-search"]
                                    sleep    5s
                                    Input Text
                                    ...    //*[@id="webshellMenu_kAutoCompleteMenuSearch"]
                                    ...    Sales Opportunities
                                    Wait Until Element Is Visible
                                    ...    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li/a/div/div/span[2]
                                    RPA.Browser.Selenium.Click Element
                                    ...    //*[@id="webshellMenu_kAutoCompleteMenuSearch_listbox"]/li/a/div/div/span[2]
                                    Sleep    5s
                                    Click Element    xpath=//a[@id="ToolBtnNew"]
                                    sleep    5s
                                    Input Text    //*[@id="AccountNameAutoField"]    ${input data list}[0]
                                    Press Keys    //*[@id="AccountNameAutoField"]    ENTER
                                    Sleep    4s
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

Business exception mail
    [Arguments]    ${configlist}    ${input data list}
    ${subject}=    Replace String    ${configlist}[3]    <name>    ${input data list}[0]
    ${mailbody}=    Replace String    ${configlist}[4]    <name>    ${input data list}[0]
    Send Message    recipients=${configlist}[2]
    ...    subject=${subject}
    ...    body=${mailbody}

System exception mail
    [Arguments]    ${configlist}
    Send Message    recipients=${configlist}[5]
    ...    subject=${configlist}[6]
    ...    body=${configlist}[7]

Config exception
    [Arguments]    ${xmllist}
    Send Message    recipients=${xmllist}[2]
    ...    subject=${xmllist}[3]
    ...    body=${xmllist}[4]

Input File Exception
    [Arguments]    ${configlist}
    Send Message    recipients=${configlist}[2]
    ...    subject=${configlist}[10]
    ...    body=${configlist}[9]

Business Exception subject if new marked as Existing
    [Arguments]    ${configlist}    ${input data list}
    ${subject}=    Replace String    ${configlist}[3]    <name>    ${input data list}[0]
    ${mailbody}=    Replace String    ${configlist}[11]    <name>    ${input data list}[0]
    Send Message    recipients=${configlist}[2]
    ...    subject=${subject}
    ...    body=${mailbody}
