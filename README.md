# MSController
A small C# library for interacting with Outlook and Excel.  
Download the DLL and XML from [releases](https://github.com/DStewart1997/MSController/releases), put the two files in your \bin\debug\ folder, add a reference to the dll and you're good to go.


## Quick examples - ExcelHandler

    ExcelHandler excelHandler = new ExcelHandler();
    excelHandler.open(FILEPATH);
    string data = excelHandler.getCell("A",1);  // Gets value from cell A1
    string dataLast = excelHandler.getLastRowCell("A");  // Gets the value from the last occupied row in column A
    excelHandler.close();
    


## Quick examples - Outlook Handler

    OutlookHandler outlookHandler = new OutlookHandler();
    // recipeint and attachmentPath can either be strings or List<string>s - attachmentPath is optional
    outlookHandler.sendMail(subject, body, recipient, attachmentPath);  
    
-------------------------------------------
    
#### Future changes
ExcelHandler
- Finish the getLastColumnCell method.
- Will likely swap round the getLastRowCell and getLastColumnCell methods as they make more sense with the names reversed.
- Implement a writeLastColumnCell method.
- Allow the workbook to be seleced.

OutlookHandler
- Nothing for now.
