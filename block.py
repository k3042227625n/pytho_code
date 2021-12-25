print(2)

import pandas as pd
import openpyxl
from glob import glob
import datetime
from natsort import natsorted
openpyxl.reader.excel.warnings.simplefilter('ignore')


# １店舗分の週売上
filepaths1 = glob('block/*.xlsx') #ファイルパスを取得
filepath1 = filepaths1[0]

def extract(filepath1):
    # 1.Excelを読み込む
    _df1 = pd.read_excel(filepath1)
    # 2.columnを取り出す
    columns = ['S','木','金','土','日','月','火','水','週累計']
    # 3.メインのデータを抽出
    df1 = _df1.iloc[3:9, [1,2,3,4,5,6,7,8,9]]
    # 4.column名を再定義
    df1.columns = columns
#     5.meta情報(mainではない付加情報)の追加
    df1['店名'] = _df1.iloc[1, 1]
    return df1

# 全店舗分の週売上
df1 = pd.DataFrame()
for filepath1 in natsorted(filepaths1):
    _df1 = extract(filepath1)
    df1 = pd.concat([df1, _df1], ignore_index=True)

table_piv = df1.pivot(index = "店名", columns = "S", values = "週累計").fillna(0)
table_piv["売上昨対"] = round(table_piv["売上昨対"], 3)*100
table_piv["客数昨対"] = round(table_piv["客数昨対"], 3)*100
table_piv.loc["ブロック"] = table_piv.sum(axis=0)

table_piv.loc["ブロック"][0] = round(table_piv.loc["ブロック"][1] / table_piv.loc["ブロック"][4], 3)*100
table_piv.loc["ブロック"][3] = round(table_piv.loc["ブロック"][2] / table_piv.loc["ブロック"][5], 3)*100
table_piv["実績"]= table_piv["実績"].astype(int)
table_piv["客数"]= table_piv["客数"].astype(int)
table_piv["昨年実績"]= table_piv["昨年実績"].astype(int)
table_piv["昨年客数"]= table_piv["昨年客数"].astype(int)
table_piv = table_piv.reset_index()
table_piv


# 1つ目のライバル店売上
filepaths = glob('block_data/block_data11/*.csv')
filepath = filepaths[0]

def extract(filepath):
    _df = pd.read_csv(filepath, encoding = "shift-jis")
    return _df

# 全店舗分のライバル店売上
df = extract(filepath)
df = pd.DataFrame()
for filepath in filepaths:
    _df = extract(filepath)
    df = pd.concat([df, _df], ignore_index= False)

# 自店とライバル店に分けて上下で結合
df_l = df.iloc[:, [0,2,3,4,6,9,11,13]].rename(columns={'自店':'B', '日別売上（自店）':'日別実績', '月間売上（自店）':'月度累計', '計画達成率（自店）':'計画', '昨年売上（自店）':'昨年累計', '昨年累計比（自店）':'累計昨対'})
df_r = df.iloc[:, [1,2,3,5,7,10,12,14]].rename(columns={'ライバル店':'B', '日別売上（ラ店）':'日別実績', '月間売上（ラ店）':'月度累計', '計画達成率（ラ店）':'計画', '昨年売上（ラ店）':'昨年累計', '昨年累計比（ラ店）':'累計昨対'})
df_11 = df_l.append(df_r, ignore_index= False)

# 7行前のデータとの差分
df_11["月度累計"] = df_11["月度累計"].diff(7)
df_11["昨年累計"] = df_11["昨年累計"].diff(7)

# 週実績、週昨対、累計昨対、計画
_total_res = []
_stores = df_11["B"].unique()

for _store in _stores:
    _df_11 = df_11[df_11["B"] == _store]
    
    wednesday = _df_11["年月日"].unique()
    week_sales = _df_11[_df_11["年月日"] == wednesday[27]]["月度累計"].sum()
    last_sales = _df_11[_df_11["年月日"] == wednesday[27]]["昨年累計"].sum()
    week_YoY = week_sales/last_sales
    total_YoY = _df_11[_df_11["年月日"] == wednesday[27]]["累計昨対"].sum()
    plan = _df_11[_df_11["年月日"] == wednesday[27]]["計画"].sum()    
    _total_res.append(dict(B=_store,週実績=round(week_sales), 昨年実績=round(last_sales), 週昨対=round(week_YoY, 3)*100 ,累計昨対=total_YoY, 計画=plan))

    
_df11 = pd.DataFrame(_total_res, columns=["B", "週実績", "昨年実績", "週昨対", "累計昨対", "計画"])
_df11 = _df11.drop(_df11.index[[2]])
_df11 = _df11.reindex(index=[0,3,4,5,1]).reset_index(drop=True)

_df11["既存店平均差"] = _df11["計画"] - 105.0
_df11 = _df11.set_index('B')

_df11.loc["B計"] = _df11.sum(axis=0)
_df11 = _df11.reset_index()

com = pd.concat([_df11, table_piv], axis=1)
com["週実績"][5] = com["売上昨対"][5]