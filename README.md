### How to make web service calls from spreadsheet
A demo showing how to connect excel to a web service 
#### What is the problem 
A spreadsheet is a great tool for data analysis and visualization. They are nimble and fast with impressive support for charting and tabular data. However, spreadsheets or visual basic has limitations as a programming tool. For data transformations that involve a sequence of complex operations, its better done by a more powerful language like Python on a Linux server. Spreadsheets also have several out of the box powerful utilities like filtering, sorting and formatting capabilities. 
#### What is the solution 
If the power of Python can be integrated with the visual capabilities of a spreadsheet, powerful applications can be developed very quickly. Here we are going to show how to wrap a Python script as a web service over FastAPI and use it with a spreadsheet. 
#### Example 
The following example will implement a sequence of operation on the server side: 
- Read prices of 5 ETFs from a csv file 
- Pivot the data
- Calculate returns from prices 
- Round decimals to two 

The front-end will be a Microsoft Excel spreadsheet with a “Fetch” button. Clicking the button will: 
- Clear the table 
- Call a web service 
- Parse the JSON response 
- Reload the table 

The chart to the right will take the table as input 

**Note:**
*A plugin may need to be activated to allow excel to make web service calls.*  All code including the prices file and Excel spreadsheet will be made available in GitHub. 

#### Explaining the application 
Create an Excel macro-enabled workbook (*xlsm) and add a form control button as shown below. 

![excel_webservice1](https://github.com/user-attachments/assets/04231298-13bf-4931-aa1c-c270eb68c400)


Check the [Sample spreadsheet](./excel_webservice.xlsm) here which is also uploaded to git 

