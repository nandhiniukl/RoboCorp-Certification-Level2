*** Settings ***
Documentation     Insert the sales data for the week and export it as a PDF.
Library           RPA.Browser.Selenium    
Library           RPA.Excel.Files
Library           RPA.HTTP
Library           RPA.PDF
Library           RPA.Tables
Library           RPA.Desktop
Library           OperatingSystem
Library           RPA.Archive
Library           RPA.Dialogs
Library           RPA.Robocorp.Vault 

*** Variables ***
${img_folder}     ${CURDIR}${/}image_files
${pdf_folder}     ${CURDIR}${/}pdf_files
${receipt_dir}    ${CURDIR}${/}receipt


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Fill the form using the data from the Excel file
    
*** Keywords ***
Take a screenshot of the robot
 
    Set Local Variable      ${img_robot}        //*[@id="robot-preview-image"]
    Wait Until Element Is Visible   ${img_robot}   
    #get the order ID   
    ${orderid}=                     Get Text            //*[@id="receipt"]/p[1]
    # Create the File Name
    Set Local Variable              ${fully_qualified_img_filename}    ${img_folder}${/}${orderid}.PNG
    Sleep   1sec
    Log To Console                  Capturing Screenshot to ${fully_qualified_img_filename}
    Capture Element Screenshot      ${img_robot}    ${fully_qualified_img_filename} 
    [Return]      ${fully_qualified_img_filename}
    
Store the receipt as a PDF file
    [Arguments]        ${ORDER_NUMBER}
    Wait Until Element Is Visible   //*[@id="receipt"]
    Log To Console                  Printing ${ORDER_NUMBER}
    ${order_receipt_html}=          Get Element Attribute   //*[@id="receipt"]  outerHTML
    Set Local Variable              ${fully_qualified_pdf_filename}    ${pdf_folder}${/}${ORDER_NUMBER}.pdf
    Html To Pdf                     content=${order_receipt_html}   output_path=${fully_qualified_pdf_filename}
    [Return]    ${fully_qualified_pdf_filename}

Embed the robot screenshot to the receipt PDF file
    [Arguments]     ${IMG_FILE}     ${PDF_FILE}
    Log To Console                  Printing Embedding image ${IMG_FILE} in pdf file ${PDF_FILE}   
    @{myfiles}=       Create List     ${IMG_FILE}:x=0,y=0
    Open PDF  ${PDF_FILE}
    Add Files To PDF    ${myfiles}    ${PDF_FILE}     ${True}
    #Close PDF           ${PDF_FILE}

Open the intranet website
    ${secret}=    Get Secret    credentials
    GetInputFromUser
    Open Available Browser    ${secret}[url]
    Wait Until Page Contains Element    css:.btn-dark   
   
Preview the robot
    # Define local variables for the UI elements
    Set Local Variable              ${btn_preview}      //*[@id="preview"]
    Set Local Variable              ${img_preview}      //*[@id="robot-preview-image"]
    Click Button                    ${btn_preview}
    Wait Until Element Is Visible   ${img_preview}

Submit the order 
    Click button                    //*[@id="order"]
    Page Should Contain Element      //*[@id="receipt"]

Close the popup
    Wait Until Element Is Visible    css:.btn-dark   10s
    Click Button    css:.btn-dark 


Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Close the popup
    Wait Until Element Is Visible   head    
    Select From List By Value    head  ${sales_rep}[Head]
    Click Element  id:id-body-${sales_rep}[Body]  
    Input Text   xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input  ${sales_rep}[Legs]
    Input Text    id:address    ${sales_rep}[Address]
    Wait Until Keyword Succeeds     10x     2s    Preview the robot
    Wait Until Keyword Succeeds     10x     2s    Submit The Order   
    Wait Until Element Is Visible  receipt
    ${imagepath}=  Take a screenshot of the robot
    ${fully_qualified_pdf_filename}=   Store the receipt as a PDF file     ${sales_rep}[Order number]
    Embed the robot screenshot to the receipt PDF file     ${imagepath}         ${fully_qualified_pdf_filename}
    ${sales_results_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf  pagenum=${sales_rep}[Order number]
    Click Button    id:order-another

Collect search query from user
    Add text input    search    label=Paste the url to download report
    ${response}=    Run dialog
    [Return]    ${response.search} 
GetInputFromUser
     ${url}=        Collect search query from user
     Download      ${url}
Fill the form using the data from the Excel file 
   
    #Download     https://robotsparebinindustries.com/orders.csv
    ${sales_reps}=    Read table from CSV  orders.csv
   
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END
    ${zip_file_name} =    Set Variable    ${OUTPUT_DIR}${/}all_receipts.zip
    Archive Folder With Zip    ${pdf_folder}    ${zip_file_name}
    Remove File    orders.csv