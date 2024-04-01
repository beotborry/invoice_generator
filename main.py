import tkinter as tk
import pandas as pd
import os
import xlwings as xw
from tkinter import filedialog
from tkinter import *
from constants import *


def get_market_from_filename(filename):
    if '주문서확인처리_CJ택배송장' in filename:
        return 'sabangnet'
    elif 'DeliveryList' in filename:
        return 'coupang'
    elif '배송등록엑셀' in filename:
        return 'saiso'
    elif '주문내역-상품준비중' in filename:
        return 'toss'


def get_big_item_list(file_path=None):
    global big_item_list
    global label_big_item_list
    if file_path is None:
        file_path = filedialog.askopenfilename()
    try:
        big_item_list = pd.read_excel(file_path)
    except:
        # add pop up message
        popup = tk.Tk()
        popup.title("Error")
        popup.geometry("400x100")
        popup.resizable(False, False)
        label = tk.Label(popup, text="Please upload the file in the correct format!")
        label.pack()
        popup.mainloop()
        
    # update the label with the file path
    label_big_item_list.config(text="File Path: " + file_path)

# write the code for getting file path and reading the file
def get_small_item_list(file_path=None):
    global small_item_list
    global label_small_item_list
    if file_path is None:
        file_path = filedialog.askopenfilename()
    try:
        small_item_list = pd.read_excel(file_path)
    except:
        # add pop up message
        popup = tk.Tk()
        popup.title("Error")
        popup.geometry("400x100")
        popup.resizable(False, False)
        label = tk.Label(popup, text="Please upload the file in the correct format!")
        label.pack()
        popup.mainloop()
        
    # update the label with the file path
    label_small_item_list.config(text="File Path: " + file_path)

def get_delivery_list():
    global file_paths
    global label_delivery_list
    file_paths = filedialog.askopenfilenames()

    # update the label with the file path
    label_delivery_list.config(text=f"Uploaded {len(file_paths)} files")

def match_sabangnet_item_names(item_names):
    if market.get() == 'sabangnet':
        return item_names

    global small_item_list
    global big_item_list

    sabangnet_item_list = small_item_list.iloc[:, 0].values
    sabangnet_big_item_list = big_item_list.iloc[:, 0].values
    
    current_small_item_list = small_item_list[SMALL_ITEM_LIST_COL[market.get()]].values
    current_big_item_list = big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values

    results = []
    for name in item_names:
        if name in current_small_item_list:
            idx = current_small_item_list.tolist().index(name)
            results.append(sabangnet_item_list[idx])
        elif all(char.isdigit() for char in name) and int(name) in current_small_item_list:
            idx = current_small_item_list.tolist().index(int(name))
            results.append(sabangnet_item_list[idx])
        elif name in current_big_item_list:
            idx = current_big_item_list.tolist().index(name)
            results.append(sabangnet_big_item_list[idx])
        elif all(char.isdigit() for char in name) and int(name) in current_big_item_list:
            idx = current_big_item_list.tolist().index(int(name))
            results.append(sabangnet_big_item_list[idx])
        else:
            print(name)
            pass
    assert len(results) == len(item_names)
    return results

def _generate_invoice(df):

    order_dates = df[ORDER_DATE_COL[market.get()]].values   
    # change the type of order_dates to datetime
    order_dates = pd.to_datetime(order_dates)
    order_dates = order_dates.strftime('%Y-%m-%d').values

    if market.get() == 'sabangnet':
        order_receptions = df[ORDER_RECEPTION_COL[market.get()]].values
    else:
        order_receptions = [market.get() for _ in range(len(df))]

    receiver_names = df[RECEIVER_NAME_COL[market.get()]].values
    receiver_addrs = df[RECEIVER_ADDR_COL[market.get()]].values
    receiver_phones = df[RECEIVER_PHONE_COL[market.get()]].values
    # make phone number format to be 010-1234-5678
    receiver_phones = [f"{phone.replace('-','')[:3]}-{phone.replace('-','')[3:7]}-{phone.replace('-','')[7:]}" if len(phone.replace('-', '')) == 11 else f"{phone.replace('-','')[:4]}-{phone.replace('-','')[4:8]}-{phone.replace('-','')[8:]}" for phone in receiver_phones]

    deliver_items = df[DELIVERY_ITEM_LIST_COL[market.get()]].values
    deliver_items = match_sabangnet_item_names(deliver_items)
    order_quantities = df[ORDER_QUANTITY_COL[market.get()]].values
    customer_names = df[CUSTOMER_NAME_COL[market.get()]].values
    delivery_msgs = df[DELIVERY_MSG_COL[market.get()]].values

    df_output = pd.DataFrame(columns=['주문일자', '접수처', '받는분', '받는분전화번호1', '받는분전화번호2', '받는분 주소',  '상품명', '수량', '배송메세지', '송장번호', '주문자명'])
    df_output['주문일자'] = order_dates
    df_output['접수처'] = order_receptions
    df_output['받는분'] = receiver_names
    df_output['받는분전화번호1'] = receiver_phones
    df_output['받는분 주소'] = receiver_addrs
    df_output['상품명'] = deliver_items
    df_output['수량'] = order_quantities
    df_output['주문자명'] = customer_names
    df_output['배송메세지'] = delivery_msgs

    # remove index
    df_output.reset_index(drop=True, inplace=True)

    return df_output

def generate_invoice():
    global small_item_list
    global big_item_list
    global output_filename
    global market
    global delivery_list
    global file_paths

    entire_small_df = pd.DataFrame()
    entire_other_df = pd.DataFrame()

    for file_path in file_paths:
        error_df = pd.DataFrame()
        market.set(get_market_from_filename(file_path))
        print(os.path.basename(file_path), market.get())
        try:
            with xw.App(visible=False) as app:
                delivery_list = xw.Book(file_path, password=PASSWORD).sheets[0].used_range.options(pd.DataFrame, index=False).value

            # delivery_list = pd.read_excel(file_path)
        except Exception as e:
            print(e)
            # add pop up message
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

        _small_item_list = small_item_list[SMALL_ITEM_LIST_COL[market.get()]].dropna().values
        receiver_name_col = RECEIVER_NAME_COL[market.get()]
        receiver_addr_col = RECEIVER_ADDR_COL[market.get()]
        delivery_item_list_col = DELIVERY_ITEM_LIST_COL[market.get()]
        order_quantity_col = ORDER_QUANTITY_COL[market.get()]
        phone_col = RECEIVER_PHONE_COL[market.get()]
        # write the code for generating invoice
        delivery_list[order_quantity_col] = delivery_list[order_quantity_col].astype(int)
        delivery_list[phone_col] = delivery_list[phone_col].astype(str)

        aggregated_delivery_list = delivery_list.groupby([receiver_name_col, receiver_addr_col]).agg({order_quantity_col: 'sum'})
        multiple_order_roster = aggregated_delivery_list[aggregated_delivery_list[order_quantity_col] > 1]
        need_to_be_examined = aggregated_delivery_list[~aggregated_delivery_list.index.isin(multiple_order_roster.index)]
        results = []

        # get '상품명' from the delivery_list whose ['받는분', '받는분 주소'] is in the need_to_be_examined
        for index, row in delivery_list.iterrows():
            if (row[receiver_name_col], row[receiver_addr_col]) in need_to_be_examined.index:
                if row[delivery_item_list_col] in _small_item_list or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in _small_item_list):
                    results.append(True)
                else:
                    if row[delivery_item_list_col] in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values):
                        results.append(False)
                    else:
                        error_df = pd.concat([error_df, row.to_frame().T], ignore_index=True)
                        # remove the row from the delivery_list
                        delivery_list.drop(index, inplace=True)
            else:
                if row[delivery_item_list_col] in _small_item_list or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in _small_item_list):
                    results.append(False)
                elif row[delivery_item_list_col] in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values or (all(char.isdigit() for char in row[delivery_item_list_col]) and int(row[delivery_item_list_col]) in big_item_list[SMALL_ITEM_LIST_COL[market.get()]].values):
                    results.append(False)
                else:
                    error_df = pd.concat([error_df, row.to_frame().T], ignore_index=True)
                    # remove the row from the delivery_list
                    delivery_list.drop(index, inplace=True)
                    
        assert len(results) == len(delivery_list)
        
        print(len(delivery_list))
        if len(error_df) > 0:
            error_df.to_excel(f"missing_items_in_{os.path.basename(file_path).replace('xlsx', 'xls')}", index=False, engine='openpyxl')

        delivery_list['small'] = results
        
        small_df = delivery_list[delivery_list['small'] == True]
        small_df = small_df.drop(columns=['small'])

        if len(small_df) > 0:
            small_df = _generate_invoice(small_df)
            entire_small_df = pd.concat([entire_small_df, small_df], ignore_index=True)
        

        other_df = delivery_list[delivery_list['small'] == False]
        other_df = other_df.drop(columns=['small'])

        if len(other_df) > 0:
            other_df = _generate_invoice(other_df)
            entire_other_df = pd.concat([entire_other_df, other_df], ignore_index=True)

        
    if len(entire_small_df) > 0:
        entire_small_df.to_excel(f'1kg.xls', index=False, engine='openpyxl')
    if len(entire_other_df) > 0:
        entire_other_df.to_excel(f'2kg.xls', index=False, engine='openpyxl')

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

    # File menu
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