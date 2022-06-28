# cmd>pip install pandas
# cmd>pip install openxrl
# cmd>pip install pyinstaller

from datetime import date
import pandas as pd
import warnings
import fnmatch
import os
import lib as me

chk_file_path = "./주문내역/"  # 주문 내역 파일을 저장할 디렉터리
shop_order_file_name = "주문내역*"
pg_file_name = dict()
pg_file_name["*_0*.xlsx"]=me.inicisCardLoad
pg_file_name["*_16*.xlsx"]=me.inicisTransLoad
pg_file_name["*_17*.xlsx"]=me.inicisGasangLoad
pg_file_name["TradeCard*.xls"]=me.kcpCardLoad
pg_file_name["TradeAcnt*.xls"]=me.kcpTransLoad
pg_file_name["신용*.xlsx"]=me.tossCardLoad
pg_file_name["계좌*.xlsx"]=me.tossTransLoad
pg_file_name["가상*.xlsx"]=me.tossGasangLoad

# === chk_file_path 디렉터리에서 주문 엑셀 파일 전부 로드 > shop_order_list , pg_order_list ===
shop_order_list = list()
pg_order_list = list()
for filename in os.listdir(chk_file_path):
    if fnmatch.fnmatch(filename,shop_order_file_name):
        excelFile = chk_file_path+filename
        shop_order_list = me.shopOrderLoad(excelFile)
    else:
        for k in pg_file_name:
            if fnmatch.fnmatch(filename,k):
                pg_data = pg_file_name[k](chk_file_path+filename)
                pg_order_list.append(pg_data)
                
pg_order_list = pd.concat(pg_order_list)
#print(pg_order_list)

# === PG 중복 결제 건 확인 ===
pg_dup = pg_order_list[pg_order_list.duplicated(['상점ID','구매자','결제금액','주문거래상태','주문번호'],keep='last')]  # PG 중복 결제 건 확인
print(pg_dup)

# === 주문번호를 기준으로 주문금액합산 ===
shop_order_list['쇼핑몰정산금액'] = shop_order_list['쇼핑몰정산금액'].astype('int')
shop_order_list = shop_order_list.groupby(by=['가맹점ID','회원ID','주문자명','주문일시','결제방법','주문번호'],as_index=False).sum('쇼핑몰정산금액')
shop_order_list= shop_order_list.reindex(columns=['가맹점ID','회원ID','주문자명','주문일시','결제방법','쇼핑몰정산금액','주문번호'])
#print(shop_order_list)

pg_order_list['PG정산금액'] = pg_order_list['PG정산금액'].astype('float')
pg_order_list['PG정산금액'] = pg_order_list['PG정산금액'].astype('int')
pg_order_list = pg_order_list.groupby(by=['상점ID','지불수단','주문번호','구매자'],as_index=False).sum('PG정산금액')
pg_order_list = pg_order_list.reindex(columns=['주문번호','PG정산금액','상점ID','지불수단','구매자'])
#print(pg_order_list)

# === 쇼핑몰 주문서와 PG 주문서 비교 ===
tot_order_list = list()
odd_order_list = list()
tot_order_list = pd.merge(shop_order_list, pg_order_list, how='outer', on='주문번호')   # 합집합
tot_order_list['비교결과'] = tot_order_list.apply(lambda x:'완료' if x['쇼핑몰정산금액'] == x['PG정산금액'] else '확인필요' ,axis=1)
odd_order_list = tot_order_list[ (tot_order_list['비교결과'] == '확인필요') & ( (tot_order_list['PG정산금액'] > 0) | (tot_order_list['쇼핑몰정산금액'] > 0)  ) ]
#print(odd_order_list)

with pd.ExcelWriter('./결과파일/'+date.today().isoformat()+"-"+'PG대조결과'+'.xlsx') as writer:
    odd_order_list.to_excel(writer, sheet_name='확인주문건' ,index=False)
    tot_order_list.to_excel(writer, sheet_name='전체데이터' ,index=False)
    shop_order_list.to_excel(writer, sheet_name='주문데이터' ,index=False)
    pg_order_list.to_excel(writer, sheet_name='PG데이터' ,index=False)

print("--------------------------------------------------------")
print()
print()

os.system("pause")


