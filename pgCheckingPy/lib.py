import pandas as pd
import os
import warnings
import fnmatch

def shopOrderLoad(excelFile): 
    try:
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str)
            orderData = pd.DataFrame(orderData,
                            columns=['가맹점ID','회원ID','주문자명','주문일시','주문상태','결제방법','포인트결제','배송비','실결제금액','총주문금액','주문번호'])
    except:
        print("엑셀파일이 존재 하지 않습니다")

    orderData = orderData[ orderData['주문상태'] != '입금대기' ]
    orderData = orderData[ orderData['결제방법'] != '기본금' ]
    orderData = orderData[ ~( (orderData['주문상태'] == '취소완료') & (orderData['결제방법'].str.contains('가상계좌') ) ) ]
    
    orderData['실결제금액'] = orderData['실결제금액'].astype('int')
    orderData['총주문금액'] = orderData['총주문금액'].astype('int')
    orderData['배송비'] = orderData['배송비'].astype('int')
    orderData['포인트결제'] = orderData['포인트결제'].astype('int')
    orderData['쇼핑몰정산금액'] = orderData.apply(lambda x:0 if x['주문상태'] in ['취소완료','환불완료','반품완료'] else x['실결제금액'] ,axis=1)

    return orderData

#------------------------------------------------------------------
def inicisCardLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str,header=3)
            orderData = pd.DataFrame(orderData, columns=['상점ID','카드계열','주문번호','구매자','승인일자','취소일자','신용카드금액 (원)','거래상태'] )
    except:
        print("엑셀파일이 존재 하지 않습니다")

    orderData.rename( columns={ '카드계열':'지불수단','신용카드금액 (원)':'결제금액','거래상태':'주문거래상태' } ,inplace = True)

    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    #orderData['정산금액'] = orderData.apply( lambda x:x['결제금액']*-1 if ( x['주문거래상태'] in ['매입전취소','매입후취소'] ) and ( len( orderData[orderData['주문번호']==x['주문번호']]) > 1 ) else x['결제금액'] ,axis=1)
    orderData['PG정산금액'] = orderData.apply(
        lambda x : x['결제금액']*-1
        if ( x['주문거래상태'] in ['매입전취소','매입후취소'] ) and ( len( orderData[orderData['주문번호']==x['주문번호']]) > 1 )
        else (
            0
            if x['주문거래상태'] in ['매입전취소','매입후취소']
            else x['결제금액'] ) ,axis=1)

    return orderData

def inicisTransLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str,header=3)
            orderData = pd.DataFrame(orderData, columns=['상점ID','지불수단','주문번호','구매자명','승인일자','취소일자','이체금액','거래상태'])
    except:
        print("엑셀파일이 존재 하지 않습니다")
        
    orderData.rename( columns={'구매자명':'구매자','이체금액':'결제금액','거래상태':'주문거래상태' }, inplace = True)
    
    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData.apply(lambda x:x['결제금액']*-1 if x['주문거래상태'] in ['취소'] else x['결제금액'] ,axis=1)
    return orderData


def inicisGasangLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str,header=3)
            orderData = pd.DataFrame(orderData, columns=['상점ID','지불수단','주문번호','구매자','승인일자','취소일자','입금금액','입금처리상태'])
    except:
        print("엑셀파일이 존재 하지 않습니다")
        
    orderData.rename( columns={ '입금금액':'결제금액','입금처리상태':'주문거래상태' },inplace = True )
    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    #orderData['PG정산금액'] = orderData.apply(lambda x:x['결제금액']*-1 if x['주문거래상태'] not in ['입금(매칭)'] else x['결제금액'] ,axis=1)
    orderData['PG정산금액'] = orderData.apply(
        lambda x : x['결제금액']*-1
        if ( x['주문거래상태'] not in ['입금(매칭)'] ) and ( len( orderData[orderData['주문번호']==x['주문번호']]) > 1 )
        else (
            0
            if x['주문거래상태'] not in ['입금(매칭)']
            else x['결제금액'] ) ,axis=1)

    return orderData

#--------------------------------------------------------
def kcpCardLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 지불수단, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, dtype = str)
            orderData = pd.DataFrame(orderData,columns=['사이트명','카드종류','주문번호','주문자','승인일자','취소일자','거래금액','최종결제상태','취소가능금액'])
            orderData = orderData[orderData['주문번호'].notna()]
    except:
        print("엑셀파일이 존재 하지 않습니다")

    orderData.rename( columns={'사이트명': '상점ID','카드종류':'지불수단' ,'주문자':'구매자','거래금액':'결제금액','최종결제상태':'주문거래상태','취소가능금액':'PG정산금액'},inplace = True )
    '''
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData.apply(lambda x:0
                                if x['주문거래상태'] in ['취소'] else ( x['취소가능금액'] if x['주문거래상태'] in ['부분취소'] else x['결제금액'] ) ,axis=1)
    '''
    return orderData


def kcpTransLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, dtype = str, header=2)
            orderData = pd.DataFrame(orderData,columns=['사이트명','은행명','주문번호','주문자','승인일자','취소일자','거래금액','거래상태','취소가능금액'])
            orderData = orderData[orderData['주문번호'].notna()]
    except:
        print("엑셀파일이 존재 하지 않습니다")

    orderData.rename( columns={'사이트명':'상점ID', '은행명':'지불수단', '주문자':'구매자','거래금액':'결제금액','거래상태':'주문거래상태','취소가능금액':'PG정산금액'},inplace = True )
    '''
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData.apply(lambda x:x['취소가능금액']
                                if x['주문거래상태'] in ['취소정완료'] else x['결제금액'] ,axis=1)
    '''
    return orderData

#--------------------------------------------------------
def tossCardLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str)
            orderData = pd.DataFrame(orderData, columns=['상점아이디(MID)','결제기관','주문번호','구매자명','결제·취소일시','결제·취소액','결제상태'] )
    except:
        print("엑셀파일이 존재 하지 않습니다")
    
    orderData.rename( columns={ '상점아이디(MID)':'상점ID','결제기관':'지불수단','구매자명':'구매자','결제·취소일시':'승인일자','결제·취소액':'결제금액','결제상태':'주문거래상태' } ,inplace = True)
    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData['결제금액']
    return orderData

def tossTransLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str)
            orderData = pd.DataFrame(orderData, columns=['상점아이디(MID)','은행','주문번호','구매자명','결제·취소일시','결제·취소액','결제상태'] )
    except:
        print("엑셀파일이 존재 하지 않습니다")
    
    orderData.rename( columns={ '상점아이디(MID)':'상점ID','은행':'지불수단','구매자명':'구매자','결제·취소일시':'승인일자','결제·취소액':'결제금액','결제상태':'주문거래상태' } ,inplace = True)
    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData['결제금액']
    return orderData


def tossGasangLoad(excelFile):
    try:
        with warnings.catch_warnings(record=True):
            # 상점ID, 주문번호, 구매자, 승인일자, 취소일자, 결제금액, 주문거래상태
            warnings.simplefilter("always")
            orderData = pd.read_excel(excelFile, engine="openpyxl",dtype = str)
            orderData = pd.DataFrame(orderData, columns=['상점아이디(MID)','은행','주문번호','구매자명','결제·취소일시','입금·취소액','결제상태'] )
    except:
        print("엑셀파일이 존재 하지 않습니다")
    
    orderData.rename( columns={ '상점아이디(MID)':'상점ID','은행':'지불수단','구매자명':'구매자','결제·취소일시':'승인일자','입금·취소액':'결제금액','결제상태':'주문거래상태' } ,inplace = True)
    orderData['결제금액'] = orderData['결제금액'].astype('float')
    orderData['결제금액'] = orderData['결제금액'].astype('int')
    orderData['PG정산금액'] = orderData['결제금액']
    return orderData




