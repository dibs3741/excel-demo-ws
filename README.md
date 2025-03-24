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
##### The Front End 
<hr>

Create an Excel macro-enabled workbook (*xlsm) and add a form control button as shown below. 

![excel_webservice1](https://github.com/user-attachments/assets/04231298-13bf-4931-aa1c-c270eb68c400)

It will auto generate a function with a default name and link the button to it. accept the default name 
```
Sub Button1_Click()
End Sub
```
##### The Server Side  
<hr>

The python end point 
```
from fastapi import FastAPI, Depends
from typing import Dict, List, Optional
import pandas as pd
from src.models import PricesDemo


app = FastAPI()

@app.get("/demo/excelapi/v1", response_model=List[PricesDemo])
def demo_excelapi_v1():
    df6 = (pd.read_csv('./db/demo_prices.csv', header=0)
        # establish the date format
        .pipe(lambda df: df.assign(asofdate = pd.to_datetime(df.asofdate)))
        .pipe(lambda df: df.assign(asofdate = df.asofdate.dt.strftime('%Y-%m-%d')))
        #
        # pivot the frame so columns are etf's
        .pipe(lambda df: df.pivot(index='asofdate', columns='ticker', values='price'))
        .pipe(lambda df: pd.DataFrame(df.to_records()))
        #
        # assign the index and sort the data in sequence
        .pipe(lambda df: df.set_index('asofdate'))
        .pipe(lambda df: df.sort_values(by='asofdate'))
        #
        # convert prices to returns
        .pipe(lambda df: df.pct_change())
        #
        # round the result to 4 dec pt
        .pipe(lambda df: df.round(4)[1:])
        .pipe(lambda df: df.reset_index())
        )
    return df6.to_dict(orient='records')
```
Check the [Sample spreadsheet](./excel_webservice.xlsm) here which is also uploaded to git 

