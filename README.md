# PaySlipGenerator
Coding test - to generate payslips
===================================

Description
-----------------------------------------------------------------------------------------------------------------------------------
This is a web application with no authentication or authorization. The only task of this application is to take employee information as input and generate paylips after calculation income tax and super.

Since different states can have different tax slabs. User can select State and income tax will be deducted accordingly.



-----------------------------------------------------------------------------------------------------------------------------------
Input and Output
-----------------------------------------------------------------------------------------------------------------------------------
The applications accepts data in excel sheet format(.xslx).

Input excel sheet has following columns, order of columns can be anything. 
Header names should be same as listed below.

=> FirstName - string/text
=> LastName - string/text
=> AnnualSalary - decimal greater than 0
=> SuperRate - decimal value between 0 - 50
=> PayPeriod - string/text


Output excel sheet format is fixed and contains 6 columns:

=> Name - string/text 'FirstName LastName'
=> GrossIncome
=> IncomeTax
=> NetIncome
=> Super
=> PayPeriod

Sample Input and Output excel sheet is present in "/TestSample" folder.



-----------------------------------------------------------------------------------------------------------------------------------
Validations/Rules
-----------------------------------------------------------------------------------------------------------------------------------
Input data

=> AnnualSalary should be a decimal value greater than 0

=> SuperRate should be a decimal value without '%' symbol and value should be between 0 - 50

=> Blank spaces are trimmed while reading information from excel sheet.
example: following three text are considered as same 'Randy'
'  Randy'
'  Randy  '
'Randy   '



-----------------------------------------------------------------------------------------------------------------------------------
Error Logging
-----------------------------------------------------------------------------------------------------------------------------------
- There is a custom error logger to handle exceptions and errors.
- All unhandled exceptions and errors are logged in 'logs.txt' file in '/PaySlipGenerator/logs' folder.
- Logger can be configured to log exceptions in file or display in console. But for now it will only display errors in console



-----------------------------------------------------------------------------------------------------------------------------------
Visual Studio Solution Structure
-----------------------------------------------------------------------------------------------------------------------------------
Solution consists of 5 projects:

=> CommonLogger: Provides static logger object to log errors and exceptions in a file.
NLog Nuget component is used to write logs. Behavior of NLog can be configured by modifying NLog.config file.


=> PaySlipEngine
=> PaySlipFactory: These two projects implements factory pattern. There is separate Engine for each state, engine handles different tax rules for each state.
Base Engine provides common methods to calculate gross income, net income and super.


=> PaySlipGenerator: It is a web project. There is a helper class that process the input excel worksheet and generates output excel worksheet.
Helper create the engine object using factory, depending on the state selected by user.

Producer/Consumer multithreading pattern is used. Multiple threads read records from excel sheet and add to a thread-safe collection. 
Separate thread reads items from collections and process them to generate payslip.


=> PaySlipTest: This project contains unit and integration test cases.
Unit test cases - Single record is processed directly by Engine object.
Integration test cases - Records are sent as excel workeet to Helper.
Negative/Exception test cases are present that fails to pass the validation/Rules defined in prior sections.



-----------------------------------------------------------------------------------------------------------------------------------
Constraints/Assumptions
-----------------------------------------------------------------------------------------------------------------------------------
=> Applications accept excel sheet in .xslx format only.
=> Sheet should contain 5 columns with header names defined in 'Input and Output' section. Order of columns can be anything.
=> Please refer 'Validation/Rules' section as well.



-----------------------------------------------------------------------------------------------------------------------------------
How to run the application
-----------------------------------------------------------------------------------------------------------------------------------
=> Download solution from gitHub.
=> Project is created using VisualStudio 2017 free community edition.
=> Build the solution and set 'PaySlipGenerator' as startup project.
=> Press 'Ctrl + F5' buttons or 'Start without Debugging' option to run the application.
=> User will land up on the Home/Default page. 
=> Select and upload a valid excel sheet(only .xslx) format and click 'Generate File' button.
=> Output excel worksheet having payslips informations gets downloaded.
=> Please check NetworkService permissions and browser settings to allow downloads. Although application will work with default settings and permissions.
=> In case of any errors, please check '/PaySlipGenerator/logs/logs.txt' file.'
=> Sample excel sheet is present in '/TestSample' folder.
