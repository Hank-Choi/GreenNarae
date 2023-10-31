import os
import random
import traceback
import urllib
from datetime import date
from urllib import request

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

columns = ["판매자 상품코드", "카테고리코드", "상품명", "상품상태", "판매가", "부가세", "재고수량", "옵션형태", "옵션명", "옵션값", "옵션가",
           "옵션 재고수량", "직접입력 옵션", "추가상품명", "추가상품값", "추가상품가", "추가상품 재고수량", "대표이미지", "추가이미지", "상세설명", "브랜드",
           "제조사", "제조일자", "유효일자", "원산지코드", "수입사", "복수원산지여부", "원산지 직접입력", "미성년자 구매", "배송비 템플릿코드", "배송방법",
           "택배사코드", "배송비유형", "기본배송비", "배송비 결제방식", "조건부무료-상품판매가 합계", "수량별부과-수량", "구간별-2구간수량", "구간별-3구간수량",
           "구간별-3구간배송비", "구간별-추가배송비", "반품배송비", "교환배송비", "지역별 차등 배송비", "별도설치비", "상품정보제공고시", "템플릿코드",
           "상품정보제공고시 품명", "상품정보제공고시 모델명", "상품정보제공고시 인증허가사항", "상품정보제공고시 제조자", "A/S 템플릿코드", "A/S 전화번호",
           "A/S 안내	판매자특이사항", "PC 즉시할인 값", "PC 즉시할인 단위", "모바일 즉시할인 값", "모바일 즉시할인 단위", "복수구매할인 조건 값",
           "복수구매할인 조건 단위", "복수구매할인 값", "복수구매할인 단위", "상품구매시 포인트 지급 값", "상품구매시 포인트 지급 단위",
           "텍스트리뷰 작성시 지급 포인트", "포토/동영상 리뷰 작성시 지급 포인트", "한달사용 텍스트리뷰 작성시 지급 포인트",
           "한달사용 포토/동영상리뷰 작성시 지급 포인트", "톡톡친구/스토어찜고객 리뷰 작성시 지급 포인트", "무이자 할부 개월", "사은품", "판매자바코드",
           "구매평 노출여부", "구매평 비노출사유", "스토어찜회원 전용여부", "ISBN", "ISSN", "독립출판", "출간일", "출판사", "글작가", "그림작가",
           "번역자명", "문화비 소득공제"]

row_sample = ["", "수정", "상품명(자동)", "신상품", "판매가(자동)", "", "5", "", "", "", "", "",
              "", "", "", "", "", "대표이미지(자동)", "추가이미지(자동)", "상세설명(자동)", "브랜드(자동)", "제조사(자동)", "", "",
              "0201035", "영국백화점", "", "", "", "2073733", "",
              "", "", "", "", "", "", "", "",
              "", "", "", "", "", "", "1888668",
              "", "", "", "", "2242440", "", "",
              "", "", "", "",
              "", "", "", "", "", "", "",
              "", "", "", "", "", "", "", "",
              "", "", "", "", "", "", "", "", "", "", "", ""]

# 수정
#######################################################################################################
information = "seletti"  # 띄어쓰기 X
url = "https://www.selfridges.com/KR/en/cat/home-tech/home/seletti/"

number_of_products = 724

saving_photo = True  # 사진 저장 할 때는 True 안하려면 False
#######################################################################################################

today = date.today().isoformat()
wb = load_workbook('template.xlsx')
ws = wb.active

try:
    os.makedirs(f"data/{today}/{information}")
except Exception:
    pass

number_of_pages = int((number_of_products - 1) / 60) + 1
# 페이지 수
for page in range(1, number_of_pages + 1):
    driver = webdriver.Chrome()
    driver.implicitly_wait(time_to_wait=random.random() * 5 + 5)
    if "?" in url:
        driver.get(url + "&pn=" + str(page))
    else:
        driver.get(url + "?pn=" + str(page))

    products = driver.find_elements(By.XPATH, '//*[@data-js-action="listing-item"]')
    product_links = []
    photo_links = []

    item: WebElement
    for idx, item in enumerate(products):
        driver_detail = webdriver.Chrome()
        driver_detail.implicitly_wait(time_to_wait=random.random() * 5 + 3)
        print(str(page) + "," + str(idx) + "th product")
        try:
            a = item.find_element(By.TAG_NAME, "a")
            textbox = item.find_element(By.CLASS_NAME, 'c-prod-card__cta-box-link-mask')
            title = textbox.find_element(By.CLASS_NAME, 'c-prod-card__cta-box-description').text
            brand_name = textbox.find_element(By.TAG_NAME, "h2").text.split("\n")[0]
            price_txt = item.find_element(By.CLASS_NAME, 'c-prod-card__cta-box-price').text
            price = price_txt[1:-3].replace(",", "")
            product_row = row_sample.copy()
            driver_detail.get(a.get_attribute('href'))
            try:
                driver_detail.find_element(By.CLASS_NAME, 'c-cookie-banner__accept-all').click()
            except Exception:
                pass
            category = list(
                map(lambda x: x.get_attribute('innerHTML'),
                    driver_detail.find_elements(By.CLASS_NAME, 'src__Link-sc-rejbql-2')))
            category_string = ",".join(category)

            try:
                size_options = list(
                    map(lambda x: x.text,
                        driver_detail.find_elements(By.CLASS_NAME, "SizeGrid__SizeContainer-sc-ckqyrx-3")))
                option_types = "size"
                size_string = ",".join(size_options)
                product_row[7] = "단독형"
                product_row[8] = option_types
                product_row[9] = size_string
            except Exception:
                pass

            product_detail = ""
            try:
                driver_detail.find_element(By.XPATH,'//*[text()="Product details"]').click()
                product_detail = str(
                    driver_detail.find_element(By.CLASS_NAME, "ProductDetails__Content-sc-rurg54-1").get_attribute(
                        'innerHTML')).replace("<ul>", "").replace("</ul>", "").replace("<br>", "")
            except Exception:
                pass

            product_links.append(a.get_attribute('href'))
            try:
                photo_link = a.find_element(By.TAG_NAME, 'img').get_attribute('src') + "&$PDP_M_ZOOM$"
                photo_links.append(photo_link)
            except TypeError:
                photo_link = "https:" + a.find_element(By.TAG_NAME, 'img').get_attribute('data-src') + "&$PDP_M_ZOOM$"
                photo_links.append(photo_link)

            photo_link2 = photo_link.replace("ALT10", "ALT01")
            photo_link3 = photo_link.replace("ALT10", "ALT02")
            photo_link4 = photo_link.replace("ALT10", "ALT03")

            if saving_photo is True:
                urllib.request.urlretrieve(photo_link,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_1.jpg')
                urllib.request.urlretrieve(photo_link2,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_2.jpg')
                urllib.request.urlretrieve(photo_link3,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_3.jpg')
                urllib.request.urlretrieve(photo_link4,
                                           f'data/{today}/{information}/{information}_{page}_{idx}_4.jpg')

            # 상품명
            product_row[2] = f"{brand_name} {title}".replace("\n", "")

            product_row[4] = str(int(price))
            # 대표이미지, 추가이미지
            product_row[17] = f"{information}_{page}_{idx}_1.jpg"
            product_row[
                18] = f"{information}_{page}_{idx}_2.jpg\n{information}_{page}_{idx}_3.jpg\n{information}_{page}_{idx}_4.jpg"
            # 제조사, 브랜드
            product_row[20] = brand_name
            product_row[21] = brand_name
            # 본문
            product_row[19] = f"<div style=\"text-align: center; font-size:16px\">" \
                              f"<img src=\"https://postfiles.pstatic.net/MjAyMjAxMDFfMTE0/MDAxNjQxMDE1NDM4MTg4.NmS6nb43C0VQXim3oGWX3KFtjk8j4nYnH5sDIxet650g._G6NDb6W7JMLYpCQ0_Do9m1TnzxvQRmGt4gjrDNbrdkg.PNG.c_maru05/23.EU.2%EC%A3%BC.%EA%B4%80%EB%B6%80%EA%B0%80%EC%84%B8%ED%8F%AC%ED%95%A8_(1).png?type=w966\" />" \
                              f"{product_detail}" \
                              f"<img src=\"{photo_link}\" />" \
                              f"<img src=\"{photo_link2}\" />" \
                              f"<img src=\"{photo_link3}\" />" \
                              f"<img src=\"{photo_link4}\" />" \
                              f"<img src=\"https://blogfiles.pstatic.net/MjAyMjAxMDFfMjky/MDAxNjQxMDE0MTA3OTEx.V4BRICGohcJ3GEV3FeyL43NdXyA9WvY7zBxAxq6v9dsg.NsM_paA_JHQuCT6fvCUaENvRjHoHgABCCFFe7MPQe-4g.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_1.jpg\" />" \
                              f"<img src=\"https://blogfiles.pstatic.net/MjAyMjAxMDFfMjQ3/MDAxNjQxMDE0MTA3OTEy.uEGfU9fGAVlMW8ElD1c_9uXlmqmfoiTaPuuQHtyawvQg.vYLMVVbvNFFrkb50Dac950zsyCuOYI31dIbPGegGii8g.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_2.jpg\" />" \
                              f"</div> "
            # category
            product_row[1] = category_string
            ws.append(product_row)
        except Exception as e:
            traceback.print_exc()
            pass
        wb.save(f"data/{today}/{information}/{information}.xlsx")
        driver_detail.close()
    driver.close()
