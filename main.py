import tkinter as tk
import pandas as pd
import os
import xlwings as xw
import phonenumbers
from tkinter import filedialog
from tkinter import *
from constants import *

# 엑셀 파일 이름으로 부터 `market` value 결정하는 함수
def get_market_from_filename(filename):
    if '주문서확인처리_CJ택배송장' in filename:
        return 'sabangnet'
    elif 'DeliveryList' in filename:
        return 'coupang'
    elif '배송등록엑셀' in filename:
        return 'saiso'
    elif '주문내역-상품준비중' in filename:
        return 'toss'

# 대형품목 엑셀 파일을 읽어서 `big_item_list`에 저장하는 함수
def get_big_item_list(file_path=None):
    global big_item_list
    global label_big_item_list
    if file_path is None:
        file_path = filedialog.askopenfilename()
    try:
        big_item_list = pd.read_excel(file_path)
    except:
        popup = tk.Tk()
        popup.title("Error")
        popup.geometry("400x100")
        popup.resizable(False, False)
        label = tk.Label(popup, text="Please upload the file in the correct format!")
        label.pack()
        popup.mainloop()
        
    label_big_item_list.config(text="File Path: " + file_path)

# 소형품목 엑셀 파일을 읽어서 `small_item_list`에 저장하는 함수
def get_small_item_list(file_path=None):
    global small_item_list
    global label_small_item_list
    if file_path is None:
        file_path = filedialog.askopenfilename()
    try:
        small_item_list = pd.read_excel(file_path)
    except:
        popup = tk.Tk()
        popup.title("Error")
        popup.geometry("400x100")
        popup.resizable(False, False)
        label = tk.Label(popup, text="Please upload the file in the correct format!")
        label.pack()
        popup.mainloop()
        
    label_small_item_list.config(text="File Path: " + file_path)

# 배송 목록 엑셀 파일을 읽어서 각 파일의 경로를 `file_paths`에 저장하는 함수
def get_delivery_list():
    global file_paths
    global label_delivery_list
    file_paths = filedialog.askopenfilenames()

    label_delivery_list.config(text=f"Uploaded {len(file_paths)} files")

# 사방넷의 상품명과 다른 업체의 상품명을 매칭하는 함수
def match_sabangnet_item_names(item_names):
    if market.get() == 'sabangnet':
        return item_names

    global small_item_list
    global big_item_list

    sabangnet_item_list = small_item_list.iloc[:, 0].values # 사방넷 소형품목 상품명 리스트
    sabangnet_big_item_list = big_item_list.iloc[:, 0].values # 사방넷 대형품목 상품명 리스트
    
    current_small_item_list = small_item_list[SMALL_ITEM_LIST_COL[market.get()]].values # 현재 업체의 소형품목 상품명 리스트
    current_big_item_list = big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values # 현재 업체의 대형품목 상품명 리스트

    results = []
    for name in item_names:
        if name in current_small_item_list: # 현재 업체의 소형품목 상품명 리스트에 있는 경우
            idx = current_small_item_list.tolist().index(name)
            results.append(sabangnet_item_list[idx])
        elif all(char.isdigit() for char in name) and int(name) in current_small_item_list: # 현재 업체의 소형품목 상품명이 숫자로만 이루어진 경우 (품목 코드인 경우)
            idx = current_small_item_list.tolist().index(int(name))
            results.append(sabangnet_item_list[idx])
        elif name in current_big_item_list: # 현재 업체의 대형품목 상품명 리스트에 있는 경우
            idx = current_big_item_list.tolist().index(name)
            results.append(sabangnet_big_item_list[idx])
        elif all(char.isdigit() for char in name) and int(name) in current_big_item_list: # 현재 업체의 대형품목 상품명이 숫자로만 이루어진 경우 (품목 코드인 경우)
            idx = current_big_item_list.tolist().index(int(name))
            results.append(sabangnet_big_item_list[idx])
        else:
            print(name)
            pass
    assert len(results) == len(item_names)
    return results

# 송장을 생성하는 함수
def _generate_invoice(df):

    order_dates = df[ORDER_DATE_COL[market.get()]].values # 주문일자
    order_dates = pd.to_datetime(order_dates)
    order_dates = order_dates.strftime('%Y-%m-%d').values # 주문일자의 형식을 'YYYY-MM-DD'로 변경

    # 접수처 설정
    if market.get() == 'sabangnet': # 사방넷의 경우 '접수처' 열에서 값을 가져옴
        order_receptions = df[ORDER_RECEPTION_COL[market.get()]].values
    else:
        order_receptions = [KOREAN_RECEPTION_DICT[market.get()] for _ in range(len(df))] # 다른 업체의 경우 '접수처'를 직접 설정

    receiver_names = df[RECEIVER_NAME_COL[market.get()]].values # 받는분 이름
    receiver_addrs = df[RECEIVER_ADDR_COL[market.get()]].values # 받는분 주소
    receiver_phones = df[RECEIVER_PHONE_COL[market.get()]].values # 받는분 전화번호

    if RECEIVER_PHONE_COL2[market.get()] is not None: # 받는분 전화번호2가 있는 경우
        receiver_phones2 = df[RECEIVER_PHONE_COL2[market.get()]].values
    else:
        receiver_phones2 = [None for _ in range(len(df))] # 받는분 전화번호2가 없는 경우

    receiver_phones_formatted = []

    for phone_num in receiver_phones:
        try:
            phone = phonenumbers.parse(phone_num, "KR")
            phone = phonenumbers.format_number(phone, phonenumbers.PhoneNumberFormat.NATIONAL) # 전화번호를 국내 형식으로 변경 (ex. 050-1234-5678)
            if phone[:4] == '050-' and phone[4] != '0': # 전화번호가 050-12345-6789로 형식이 정해진 경우 0501-2345-6789로 변경
                phone = phone[:3] + phone[4] + '-' + phone[5:]
        except:
            phone = phone_num

        receiver_phones_formatted.append(phone)

    assert len(receiver_phones) == len(receiver_phones_formatted)
    receiver_phones = receiver_phones_formatted

    receiver_phones2_formatted = []
    for phone_num in receiver_phones2:
        try:
            phone = phonenumbers.parse(phone_num, "KR")
            phone = phonenumbers.format_number(phone, phonenumbers.PhoneNumberFormat.NATIONAL) # 전화번호2를 국내 형식으로 변경 (ex. 050-1234-5678)
            if phone[:4] == '050-' and phone[4] != '0': # 전화번호가 050-12345-6789로 형식이 정해진 경우 0501-2345-6789로 변경
                phone = phone[:3] + phone[4] + '-' + phone[5:]
        except:
            phone = phone_num

        receiver_phones2_formatted.append(phone)

    assert len(receiver_phones2) == len(receiver_phones2_formatted)
    receiver_phones2 = receiver_phones2_formatted

    deliver_items = df[DELIVERY_ITEM_LIST_COL[market.get()]].values # 상품명
    deliver_items = match_sabangnet_item_names(deliver_items) # 상품명을 사방넷의 상품명과 매칭
    order_quantities = df[ORDER_QUANTITY_COL[market.get()]].values # 수량
    customer_names = df[CUSTOMER_NAME_COL[market.get()]].values # 주문자명
    delivery_msgs = df[DELIVERY_MSG_COL[market.get()]].values # 배송메세지

    df_output = pd.DataFrame(columns=['주문일자', '접수처', '받는분', '받는분전화번호1', '받는분전화번호2', '받는분 주소',  '상품명', '수량', '배송메세지', '송장번호', '주문자명', '박스수량', '박스타입', '사방넷주문번호', '', '고객요청메세지(CJ)'])
    df_output['주문일자'] = order_dates
    df_output['접수처'] = order_receptions
    df_output['받는분'] = receiver_names
    df_output['받는분전화번호1'] = receiver_phones
    df_output['받는분전화번호2'] = receiver_phones2
    df_output['받는분 주소'] = receiver_addrs
    df_output['상품명'] = deliver_items
    df_output['수량'] = order_quantities
    df_output['주문자명'] = customer_names
    df_output['배송메세지'] = delivery_msgs

    if market.get() == 'sabangnet':
        df_output['사방넷주문번호'] = df[SABANGNET_ORDER_NUM_COL].values
    else:
        df_output['사방넷주문번호'] = [None for _ in range(len(df))]

    # remove index
    df_output.reset_index(drop=True, inplace=True)

    return df_output

# 전체 송장을 소형품목과 대형품목으로 나누어 송장을 생성하는 함수
def generate_invoice():
    global small_item_list
    global big_item_list
    global output_filename
    global market
    global delivery_list
    global file_paths

    entire_small_df = pd.DataFrame() # 소형품목 송장
    entire_other_df = pd.DataFrame() # 대형품목 송장

    for file_path in file_paths: # 각 파일에 대해 송장 생성
        error_df = pd.DataFrame() # 송장 생성 중 에러가 발생한 행을 저장하는 데이터프레임
        market.set(get_market_from_filename(file_path))
        print(os.path.basename(file_path), market.get())
        try:
            with xw.App(visible=False) as app:
                delivery_list = xw.Book(file_path, password=PASSWORD).sheets[0].used_range.options(pd.DataFrame, index=False).value
        except Exception as e:
            print(e)
            popup = tk.Tk()
            popup.title("Error")
            popup.geometry("400x100")
            popup.resizable(False, False)
            label = tk.Label(popup, text="Please upload the file in the correct format!")
            label.pack()
            popup.mainloop()

        if 'Unnamed: 0' in delivery_list.columns or None in delivery_list.columns:
            with xw.App(visible=False) as app:
                # read the first sheet starting from the second row
                delivery_list = xw.Book(file_path, password=PASSWORD).sheets[0].range('A2').expand().options(pd.DataFrame, index=False).value

        print(len(delivery_list))
        output_filename = os.path.basename(file_path)

        _small_item_list = small_item_list[SMALL_ITEM_LIST_COL[market.get()]].dropna().values # "소형품목" 엑셀 파일에서의 각 업체의 소형품목 상품명 리스트
        receiver_name_col = RECEIVER_NAME_COL[market.get()] # 받는분 이름 리스트
        receiver_addr_col = RECEIVER_ADDR_COL[market.get()] # 받는분 주소 리스트
        delivery_item_list_col = DELIVERY_ITEM_LIST_COL[market.get()] # 상품명 리스트
        order_quantity_col = ORDER_QUANTITY_COL[market.get()] # 수량 리스트
        phone_col = RECEIVER_PHONE_COL[market.get()] # 받는분 전화번호 리스트
        delivery_list[order_quantity_col] = delivery_list[order_quantity_col].astype(int)
        delivery_list[phone_col] = delivery_list[phone_col].astype(str)

        aggregated_delivery_list = delivery_list.groupby([receiver_name_col, receiver_addr_col]).agg({order_quantity_col: 'sum'}) # 받는분 이름과 받는분 주소로 그룹화하여 수량을 합침
        multiple_order_roster = aggregated_delivery_list[aggregated_delivery_list[order_quantity_col] > 1] # 상품을 2개 이상 주문한 경우
        need_to_be_examined = aggregated_delivery_list[~aggregated_delivery_list.index.isin(multiple_order_roster.index)] # 상품을 1개만 주문한 경우
        results = []

        for index, row in delivery_list.iterrows(): # 모든 주문 내역에 대해서 1kg에 해당하는 지 검사하는 반복문
            if (row[receiver_name_col], row[receiver_addr_col]) in need_to_be_examined.index:
                if row[delivery_item_list_col] in _small_item_list or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in _small_item_list): # 1kg에 해당하는 경우
                    results.append(True)
                else:
                    if row[delivery_item_list_col] in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values): # 2kg에 해당하는 경우
                        results.append(False)
                    else: # 상품명이 '소형품목', '대형품목' 엑셀파일 모두에 존재하지 않는 경우 (에러)
                        error_df = pd.concat([error_df, row.to_frame().T], ignore_index=True)
                        delivery_list.drop(index, inplace=True)
            else:
                if row[delivery_item_list_col] in _small_item_list or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in _small_item_list): # 2개 이상 주문한 경우는 2kg로 처리
                    results.append(False)
                elif row[delivery_item_list_col] in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values): # 대형 품목도 여러 개 주문한 경우는 2kg로 처리
                    results.append(False)
                else: # 상품명이 '소형품목', '대형품목' 엑셀파일 모두에 존재하지 않는 경우 (에러)
                    error_df = pd.concat([error_df, row.to_frame().T], ignore_index=True)
                    delivery_list.drop(index, inplace=True)
                    
        assert len(results) == len(delivery_list)
        
        print(len(delivery_list))
        if len(error_df) > 0: # 에러가 발생한 행이 하나라도 존재하는 경우 이를 엑셀 파일로 저장
            error_df.to_excel(f"missing_items_in_{os.path.basename(file_path).replace('xlsx', 'xls')}", index=False, engine='openpyxl')

        delivery_list['small'] = results
        
        small_df = delivery_list[delivery_list['small'] == True]
        small_df = small_df.drop(columns=['small']) # 1kg에 해당하는 주문 내역만 추출

        if len(small_df) > 0: # 1kg 송장 생성
            small_df = _generate_invoice(small_df)
            entire_small_df = pd.concat([entire_small_df, small_df], ignore_index=True)
        

        other_df = delivery_list[delivery_list['small'] == False]
        other_df = other_df.drop(columns=['small']) # 2kg에 해당하는 주문 내역만 추출

        if len(other_df) > 0: # 2kg 송장 생성
            other_df = _generate_invoice(other_df)
            entire_other_df = pd.concat([entire_other_df, other_df], ignore_index=True)

        
    if len(entire_small_df) > 0:
        entire_small_df.to_excel(f'1kg.xls', index=False, engine='openpyxl') # 1kg 송장 엑셀 파일 저장
    if len(entire_other_df) > 0:
        entire_other_df.to_excel(f'2kg.xls', index=False, engine='openpyxl') # 2kg 송장 엑셀 파일 저장

    # add pop up message
    popup = tk.Tk()
    popup.title("Invoice Generation")
    popup.geometry("400x100")
    popup.resizable(False, False)
    label = tk.Label(popup, text="Invoice has been generated successfully!")
    label.pack()
    popup.mainloop()

def main():
    root = tk.Tk()
    root.title("Invoice Generator")
    root.geometry("800x800")
    root.resizable(True, True)
    menu_bar = tk.Menu(root)
    root.config(menu=menu_bar)
    global label_small_item_list
    global label_delivery_list
    global market
    global label_big_item_list

    # 파일 업로드 메뉴
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Upload 1kg Item List", command=get_small_item_list)
    file_menu.add_command(label="Upload 2kg Item List", command=get_big_item_list)
    file_menu.add_command(label="Upload Delivery List", command=get_delivery_list)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)
    menu_bar.add_cascade(label="File", menu=file_menu)

    label_small_item_list = tk.Label(root, text="1kg File Path: ")
    label_small_item_list.pack()

    label_big_item_list = tk.Label(root, text="2kg File Path: ")
    label_big_item_list.pack()

    label_delivery_list = tk.Label(root, text="Upload Delivery Lists")
    label_delivery_list.pack()

    market = tk.StringVar()

    button_generate_invoice = tk.Button(root, text="Generate Invoice", command=generate_invoice)
    button_generate_invoice.pack()

    get_small_item_list(file_path=SMALL_ITEM_FILEPATH)
    get_big_item_list(file_path=BIG_ITEM_FILEPATH)

    root.mainloop()

if __name__ == '__main__':
    main()