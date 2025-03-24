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
    #console.log(df6)
    return df6.to_dict(orient='records')                        
