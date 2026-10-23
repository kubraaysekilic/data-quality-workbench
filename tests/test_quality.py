import io
import pandas as pd
import pytest
from openpyxl import load_workbook
from quality import *

def test_preserves_zero_and_na_literal():
    df=load_table(b'id,value\n001,NA\n002,\n','x.csv')
    assert df.iloc[0].tolist()==['001','NA']
    assert len(inspect(df))==1

def test_duplicate_and_invalid():
    df=pd.DataFrame({'email':['bad','a@b.com','a@b.com'],'date':['2026-02-30','2026-01-01','2026-01-01']})
    report=inspect(df,['email'],['date'])
    assert set(report['Sorun'])=={'E-posta biçimi','Tarih biçimi','Yinelenen satır'}
    assert report.iloc[-1]['Satır']==4

def test_selected_corrections_do_not_mutate_source():
    df=pd.DataFrame({'a':[' X ','X',''],'date':['23.10.2026']*3})
    result=transform(df,trim=True,drop_duplicates=True,date_columns=['date'],date_format='%d.%m.%Y',normalize_dates=True)
    assert len(result)==2 and result.iloc[0]['date']=='2026-10-23'
    assert df.iloc[0]['a']==' X '

def test_fill_only_selected_columns():
    df=pd.DataFrame({'a':[' '],'b':['']})
    assert transform(df,fill_columns=['a'],fill_value='unknown').iloc[0].tolist()==['unknown','']

def test_invalid_dates_stay_unchanged():
    df=pd.DataFrame({'d':['not a date']})
    assert transform(df,date_columns=['d'],normalize_dates=True).equals(df)

def test_csv_and_excel_formula_safety():
    df=pd.DataFrame({'=header':['=1+1','@command','001']})
    assert "'=1+1" in csv_bytes(df).decode('utf-8-sig')
    book=load_workbook(io.BytesIO(xlsx_bytes(df)),data_only=False)
    assert book.active['A1'].data_type=='s' and book.active['A2'].data_type=='s'
    assert book.active['A2'].value=='=1+1'

def test_excel_roundtrip():
    df=pd.DataFrame({'id':['001','002'],'mail':['a@b.com','']})
    assert load_table(xlsx_bytes(df),'sample.xlsx').equals(df)

@pytest.mark.parametrize('raw',[b'a,a\n1,2',b'a,b\n1,2,3',b',b\n1,2'])
def test_rejects_ambiguous_headers_and_shape(raw):
    with pytest.raises(ValueError):load_table(raw,'x.csv')

def test_exact_dates_and_calendar_validation():
    assert parse_date('2024-02-29','%Y-%m-%d')=='2024-02-29'
    for value in ['2026-02-29','2026-1-01']:
        with pytest.raises(ValueError):parse_date(value,'%Y-%m-%d')
