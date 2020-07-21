import os
import traceback
import urllib
from datetime import date
from urllib import request

import requests
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook

columns = ["상품상태", "카테고리ID", "상품명", "판매가", "재고수량",
           "A/S 안내내용", "A/S 전화번호", "대표 이미지 파일명", "추가 이미지 파일명", "상품 상세정보",
           "판매자 상품코드", "판매자 바코드", "제조사", "브랜드", "제조일자",
           "유효일자", "부가세", "미성년자 구매", "구매평 노출여부", "원산지 코드",
           "수입사", "복수원산지 여부", "원산지 직접입력", "배송방법", "배송비 유형",
           "기본배송비", "배송비 결제방식", "조건부무료-상품판매가합계", "수량별부과-수량", "반품배송비",
           "교환배송비", "지역별 차등배송비 정보", "별도설치비", "판매자 특이사항", "즉시할인 값",
           "즉시할인 단위", "복수구매할인 조건 값", "복수구매할인 조건 단위", "복수구매할인 값", "복수구매할인 단위",
           "상품구매시 포인트 지급 값", "상품구매시 포인트 지급 단위", "텍스트리뷰 작성시 지급 포인트",
           "포토/동영상 리뷰 작성시 지급 포인트", "한달사용\n 텍스트리뷰 작성시 지급 포인트",
           "한달사용\n 포토 / 동영상리뷰 작성시 지급 포인트", "톡톡친구 / 스토어찜고객\n 리뷰 작성시 지급 포인트", "무이자 할부 개월", "사은품", "옵션형태",
           "옵션명", "옵션값", "옵션가", "옵션 재고수량", "추가상품명",
           "추가상품값", "추가상품가", "추가상품 재고수량", "상품정보제공고시 품명", "상품정보제공고시 모델명",
           "상품정보제공고시 인증허가사항", "상품정보제공고시 제조자", "스토어찜회원 전용여부", "문화비 소득공제", "ISBN", "독립출판"]

row_sample = ["신상품", "수정", "수정", "수정", "5",
              "그린나래는 세일행사중인 해외명품을 고객님께 구매대행 서비스를 제공하는 업체입니다.\n주문하신 제품의 하자가 있는 경우 교환,반품은 가능합니다만,\n구입초기의 하자가 아닌 착용 및 세탁 "
              "후 발생한 하자의 경우 고객님께서 직접 해당 브랜드 매장을 통해 수선을 하셔야 합니다.\n또한 아울렛 상품의 경우 제거가능한 얼룩, 박음질/소재 불만족 등과 같은 미세하자를 사유로 한 "
              "반품접수는 어렵다는 점 참고하시기 바랍니다.", "010-4783-1346", "", "", "TEST",
              "", "", "", "", "",
              "", "과세상품", "Y", "Y", "03",
              "", "N", "", "택배, 소포, 등기", "유료",
              "30000", "선결제", "", "", "80000",
              "100000", "", "", "", "",
              "", "", "", "", "",
              "", "", "",
              "", "",
              "", "", "", "", "",
              "", "", "", "", "",
              "", "", "", "", "",
              "", "", "N", "", "", "", ""
              ]

# 수정
#######################################################################################################
information = "shoes_womens"  # 띄어쓰기 X
url = "https://www.selfridges.com/KR/en/cat/shoes/womens/"
saving_photo = True  # 사진 저장 할 때는 True 안하려면 False
#######################################################################################################


today = date.today().isoformat()
wb = Workbook()
ws = wb.active
ws.append(columns)

try:
    os.makedirs(f"data/{today}/{information}")
except Exception:
    pass

total = requests.get(url)
a=BeautifulSoup(total.content, "html.parser").find_all("div", { "class": "plp-listing-load-status"})
print(a)
number_of_products = int(re.findall("\d+", BeautifulSoup(total.content, "html.parser").find_all("div", {
    "class": "plp-listing-load-status"})[0].get_text())[1])
number_of_pages = int((number_of_products - 1) / 60) + 1
# 페이지 수
for page in range(1, number_of_pages + 1):
    if "?" in url:
        result = requests.get(url + "&pn=" + str(page))
    else:
        result = requests.get(url + "?pn=" + str(page))
    list_src = result.content

    soup = BeautifulSoup(list_src, "html.parser")

    products = soup.find_all("div", {"data-js-action": "listing-item"})
    product_links = []
    photo_links = []
    for idx, item in enumerate(products):
        print(str(page) + "," + str(idx) + "th product")
        try:
            a = item.find_all("a")[0]
            textbox = item.find_all("a", {"class": "c-prod-card__cta-box-link-mask"})[0]
            title = textbox.find("span",{"class":"c-prod-card__cta-box-description"}).get_text()[:-1]
            brand_name = textbox.find("h5").get_text()
            price_txt = item.find("span", {"class": "c-prod-card__cta-box-price"}).get_text()
            price = price_txt[2:-3].replace(",", "")
            product_row = row_sample.copy()
            product_detail_page = requests.get("https://www.selfridges.com" + a.attrs['href'])
            # product_detail_page = requests.get(
            #     "https://www.selfridges.com/KR/en/cat/zadigvoltaire-sourca-v-neck-cashmere-jumper_R00007777/?previewAttribute=FAUVE")
            product_detail_src = product_detail_page.content

            soup_detail_page = BeautifulSoup(product_detail_src, "html.parser")
            product_detail_soup = soup_detail_page.find_all("article", {"id": "content1"})[0].find("div",
                                                                                              {"class": "c-tabs__copy"})
            product_detail = str(product_detail_soup).replace("<ul>","").replace("</ul>","")

            category_section = soup_detail_page.find_all("section", {"class": "c-breadcrumb"})[0]
            category = list(
                map(lambda x: x.get_text(),
                    category_section.find_all("span", {"itemprop": "name"})))
            category_string = ",".join(category)

            filter_section = soup_detail_page.find_all("section", {"data-action": "filter"})[0]
            option_types = ""
            size_string = ""
            try:
                size_options = list(
                    map(lambda x: x.find("span").get_text(),
                        filter_section.find_all("div", {"data-select-name": "Size"})[0].find_all("span", {
                            "class": "c-select__dropdown-item"})))
                option_types = "size"
                size_string = ",".join(size_options)
            except IndexError as e:
                pass

            product_links.append(a.attrs['href'])
            try:
                photo_link = a.find_all('img')[0].attrs['src'] + "&$PDP_M_ZOOM$"
                photo_links.append(photo_link)
            except KeyError:
                photo_link = a.find_all('img')[0].attrs['data-src'] + "&$PDP_M_ZOOM$"
                photo_links.append(photo_link)

            photo_link2 = photo_link.replace("ALT10", "ALT01")
            photo_link3 = photo_link.replace("ALT10", "ALT02")
            photo_link4 = photo_link.replace("ALT10", "ALT03")

            if saving_photo is True:
                urllib.request.urlretrieve("http:" + photo_link,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_1.jpg')
                urllib.request.urlretrieve("http:" + photo_link2,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_2.jpg')
                urllib.request.urlretrieve("http:" + photo_link3,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_3.jpg')
                urllib.request.urlretrieve("http:" + photo_link4,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_4.jpg')

            # 상품명
            product_row[2] = f"{brand_name}{title}"
            # 가격
            product_row[3] = str(int(price))
            # 대표이미지, 추가이미지
            product_row[7] = f"{information}_{page}_{idx}_1.jpg"
            product_row[
                8] = f"{information}_{page}_{idx}_2.jpg, {information}_{page}_{idx}_3.jpg, {information}_{page}_{idx}_4.jpg"
            # 제조사, 브랜드
            product_row[12] = brand_name
            product_row[13] = brand_name
            # 본문
            product_row[9] = f"<div style=\"text-align: center; font-size:16px\">" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMDA2MjVfMjc2/MDAxNTkzMDk0MzI3Nzc5.OwUQzOjgBoxEHhuP7Xaz2ex8NWmJKgWngDxQbpe9FJAg.LqNhExQ8scg7FA99opGkWeF-MWbrCTERv1_64fetnd8g.PNG.c_maru05/22.%EC%85%80%ED%94%84%EB%A6%AC%EC%A7%802.png\" />" \
                             f"{str(product_detail)}" \
                             f"<img src=\"http:{photo_link}\" />" \
                             f"<img src=\"http:{photo_link2}\" />" \
                             f"<img src=\"http:{photo_link3}\" />" \
                             f"<img src=\"http:{photo_link4}\" />" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMDA0MjNfMTkw/MDAxNTg3NjM1NDE2MjYx.5YU97JBTTScA3RGNXkCzACcQDKDoLuFVKOgryL_qucMg.cNYjXOFMHBCfBUDfeuQdtxJsye2yfPlNiY-yY4UHuucg.PNG.c_maru05/%EC%85%80%ED%94%84%EB%A6%AC%EC%A7%80_%EC%82%AC%EC%9D%B4%EC%A6%88%ED%91%9C2.png\" />" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMDA2MjVfMjE0/MDAxNTkzMDk0MzI0MzQw.s4A3jERbU2RaHHD0ThxwLxlCPZFY9R6Gzb1fzAMPHA0g.BhuK6w8Ch0wiCpBJtXzzQIiL3LMgkdyV1hIOWz33rjMg.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_1.jpg\" />" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMDA2MjVfMjkg/MDAxNTkzMDk0MzI1MTEx.GDen6g0QSmGDg7IN0auZOxbeDUNfSvtbL2Rt-MoIxBYg.EFbfSi6yLxEtsLKYOQmvaeHrzWEgcxyw_JsbUN1s8oMg.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_2.jpg\" />" \
                             f"</div> "
            # category
            product_row[66] = category_string
            # 옵션값
            if size_string != "":
                product_row[49] = "단독형"
                product_row[50] = option_types
                product_row[51] = size_string
            ws.append(product_row)
        except Exception as e:
            traceback.print_exc()
            pass

wb.save(f"data/{today}/{information}/{information}.xlsx")