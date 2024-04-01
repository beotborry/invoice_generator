import os
print(os.getcwd())
SMALL_ITEM_FILEPATH = os.path.join(os.getcwd(), "소형품목.xlsx")
BIG_ITEM_FILEPATH = os.path.join(os.getcwd(), "대형품목.xlsx")

SMALL_ITEM_LIST_COL = {
    'sabangnet': "사방넷 상품명",
    'coupang': "쿠팡 품목코드",
    'toss': "토스 상품명",
    'saiso': "사이소 상품코드"
}

RECEIVER_NAME_COL = {
    'sabangnet': "받는분",
    'coupang': "수취인이름",
    'toss': '수령인',
    'saiso': '수취인',
    'nh': '수취인명',
    'auction': '수령인명'
}

RECEIVER_ADDR_COL = {
    'sabangnet': "받는분 주소",
    'coupang': "수취인 주소",
    'toss': '주소',
    'saiso': '주소',
    'nh': '수취인주소',
    'auction': '주소'
}

DELIVERY_ITEM_LIST_COL = {
    'sabangnet': "상품명", 
    'coupang': "업체상품코드",
    'toss': "옵션",
    'saiso': "상품코드",
    'nh': '단품명',
    'auction': '상품명'
}

ORDER_QUANTITY_COL = {
    'sabangnet': "수량",
    'coupang': "구매수(수량)",
    'toss': '수량',
    'saiso': '수량',
    'nh': '주문수량',
    'auction': '수량'
}

ORDER_DATE_COL = {
    'sabangnet': "주문일자",
    'coupang': "주문일",
    'toss': "주문일자",
    'saiso': "주문일자",
    'nh': "주문일자",
    'auction': "주문일자(결제확인전)"
}

ORDER_RECEPTION_COL = {
    'sabangnet': "접수처",
}

RECEIVER_PHONE_COL = {
    'sabangnet': "받는분전화번호2",
    'coupang': "수취인전화번호",
    'toss': '수령인전화번호',
    'saiso': '수취인 연락처 1',
    'nh': '수취인휴대폰번호',
    'auction': '수령인전화번호'
    
}    

CUSTOMER_NAME_COL = {
    'sabangnet': "주문자명",
    'coupang': "구매자",
    'toss': '주문자명',
    'saiso': '주문자',
    'nh': '주문자명',
    'auction': '구매자명'
}

DELIVERY_MSG_COL = {
    'sabangnet': "배송메세지",
    'coupang': "배송메세지",
    'toss': '요청사항',
    'saiso': '택배사 전달사항',
    'nh': '배송요청내용',
    'auction': '배송시 요구사항'
}

PASSWORD = "0000"