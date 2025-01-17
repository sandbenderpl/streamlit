import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import pandas as pd
from openpyxl.styles import Alignment
import numpy as np
import pandas as pd
import pyodbc
import datetime
import psycopg2
from dateutil.relativedelta import relativedelta
conn = psycopg2.connect(
    host="195.167.155.89",     
    database="fuksiarz_dwh",  
    user="bi_user",  
    password="hyBJdTwclvG0P31fMIn4_bi_usr",
    port="17318"
)
cursor = conn.cursor()

data1= datetime.date(2025, 1, 16)
data2= datetime.date(2025, 1, 16)
t0,t1,t2=[],[],[]
cursor.execute("""select transaction_type, sum(transaction_amount) from balance_transaction_journal where 
                transaction_date::date>='"""+str(data1)+"""' and transaction_date::date<='"""+str(data2)+"""'
                group by transaction_type
               union 
               select 10000+transaction_type, sum(transaction_amount) from offers__balance_transaction_journal where 
                transaction_date::date>='"""+str(data1)+"""' and transaction_date::date<='"""+str(data2)+"""'
                group by transaction_type
               union
               select 1001, sum(stake) from bet_slip where slip_id in (select reference_id  from balance_transaction_journal where 
                transaction_date::date>='"""+str(data1)+"""' and transaction_date::date<='"""+str(data2)+"""' and transaction_type=1 and transaction_amount=0)
               union
               select 1003, sum(stake) from bet_slip where slip_id in (select reference_id  from balance_transaction_journal where 
                transaction_date::date>='"""+str(data1)+"""' and transaction_date::date<='"""+str(data2)+"""' and transaction_type=3 and transaction_amount=0)
               union
               select 1013, sum(stake) from bet_slip where slip_id in (select reference_id  from balance_transaction_journal where 
                transaction_date::date>='"""+str(data1)+"""' and transaction_date::date<='"""+str(data2)+"""' and transaction_type=13)
               """)
rows = cursor.fetchall()
for i in range(len(rows)):
    t0.append(i+1)
    t1.append(rows[i][0])
    t2.append(rows[i][1])
df=pd.DataFrame({
    "kat":t1,
    "stake":t2
})


a1,a2=[],[]
x11=(df[df.kat==1][['stake']].sum()+df[df.kat==1001][['stake']].sum()+df[df.kat==10001][['stake']].sum()-df[df.kat==3][['stake']].sum()-
     df[df.kat==1003][['stake']].sum()-df[df.kat==10003][['stake']].sum()-df[df.kat==1013][['stake']].sum()-df[df.kat==5][['stake']].sum()+df[df.kat==3945][['stake']].sum()).sum()
a1.append("betamount")
a2.append(round(x11,2))
x12=(df[df.kat==2][['stake']].sum()+df[df.kat==11][['stake']].sum()+df[df.kat==10002][['stake']].sum()-df[df.kat==10004][['stake']].sum()-df[df.kat==4][['stake']].sum()-df[df.kat==1013][['stake']].sum()
     +df[df.kat==3940][['stake']].sum()).sum()
a1.append("wiamount")
a2.append(round(x12,2))
x13=(df[df.kat==187][['stake']].sum()-df[df.kat==186][['stake']].sum()+df[df.kat==1505][['stake']].sum()+df[df.kat==504][['stake']].sum()).sum()
a1.append("bonus")
a2.append(round(x13,2))
x14=(x11*0.88- df[df.kat==1001][['stake']].sum()-df[df.kat==10001][['stake']].sum()+df[df.kat==1003][['stake']].sum()+df[df.kat==10003][['stake']].sum()-x12-x13-
     df[df.kat==10004][['stake']].sum()+df[df.kat==10002][['stake']].sum()+df[df.kat==13][['stake']].sum()-df[df.kat==10506][['stake']].sum()).sum()
a1.append("NGR")
a2.append(round(x14,2))

wyniki=pd.DataFrame({
    "nazwa":a1,
    "stake":a2
}
)
wyniki