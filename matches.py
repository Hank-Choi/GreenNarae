import os
import re
import traceback
from ast import literal_eval
from datetime import date
from io import BytesIO

from PIL import Image

import requests
from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook
import ssl

ssl._create_default_https_context = ssl._create_unverified_context


def hex_to_str(string):
    return bytes.decode(literal_eval(f"b'{string}'"), 'utf-8')


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
           "상품정보제공고시 인증허가사항", "상품정보제공고시 제조자", "스토어찜회원 전용여부", "문화비 소득공제", "ISBN", "ISSN", "독립출판"]

row_sample = ["신상품", "수정", "수정", "수정", "5",
              "그린나래는 해외 명품제품에 대한 고객님의 직구를 대신해드리는 구매대행업체입니다. 주문하신 제품에 심각한 하자가 있는 경우 교환,반품을 도와드립니다만"
              "봉제오류, 잔기스 등 사소한 하자의 경우는 저희가 책임을 지지 않는 범위이니 참고하시고 신중한 구입을 해주시기를 부탁드립니다.", "010-4783-1346", "", "", "TEST",
              "", "", "", "", "",
              "", "과세상품", "Y", "Y", "03",
              "", "N", "", "택배, 소포, 등기", "유료",
              "40000", "선결제", "", "", "180000",
              "180000", "", "", "", "",
              "", "", "", "", "",
              "", "", "",
              "", "",
              "", "", "", "", "",
              "", "", "", "", "",
              "", "", "", "", "",
              "", "", "N", "", "", "", "", ""
              ]

# 수정
#######################################################################################################
information = "match_candle"  # 띄어쓰기  X
url = "https://www.matchesfashion.com/en-kr/womens/shop/clothing/dresses"
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

headers = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate, br',
    'cache-control': 'max-age=0',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
    'Content-Type': 'application/json',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36',
    'cookie': 'SESSION_TID=KXG5W33WEJCT8-59VCXB; plpLayoutMobile=2; plpLayoutTablet=2; plpLayoutDesktop=3; plpLayoutLargeDesktop=4; JSESSIONID=sc~C7DEEDE6AE123AA61F21C22CF400E753; country=KOR; indicativeCurrency=KRW; sizeTaxonomy=""; gender=womens; loggedIn=false; saleRegion=APAC; defaultSizeTaxonomy=WOMENSSHOESKOSEARCH; AWSALBAPP-1=_remove_; AWSALBAPP-2=_remove_; AWSALBAPP-3=_remove_; _pxhd=RmVTwqONWbVqv5SQ1KTKVK0BXk4dA3biVZ04ZBE3oHDXfhgrqEcNYgOE4XNEgd6Z8a6wW8a7FRW69EXqtmhL0g==:l80PUkWcpSUupEFaa/PzBJNqMHtL/i4NO/VgOAA3-YRGW6Ri2D-CSEPxpIoiKGK8tPHLsKyVTKj3dfAyOQpr2-VXCWMnaZmzBcubiecdsgw=; language=ko; ab-user-id=17; _gcl_au=1.1.1777112756.1640094454; _cs_c=0; _fbp=fb.1.1640094457236.1767171833; _pin_unauth=dWlkPVptSmtaREl6Wm1VdE5UQXpOaTAwT0RNNUxXSmlORFF0T0RRNU1XSmhabUkxT0RWaA; _gid=GA1.2.1639333752.1640094459; __tmbid=sg-1640094459-2f5180dea1cd4e27a43190f84dd34432; iadvize-2161-vuid=d5e747f3058db2e665da6c0c3f7056c461c1dafb34730; sailthru_visitor=1a5f4152-1fd0-45b3-9ec6-1ff5c8df5b3f; _pxvid=8a6bb4f7-6264-11ec-a538-6c4679435972; pxcts=9058e7c0-6264-11ec-83e7-f36ee24817b6; rskxRunCookie=0; rCookie=s8t1cxqtodo17lpb7tyuokxg5w950; signed-up-for-updates=true; billingCurrency=EUR; _cs_id=41653424-df29-a34a-822e-1bea448b83a0.1640094456.2.1640097168.1640097152.1.1674258456845; _cs_s=2.0.0.1640098968727; _uetsid=8ea6c6c0626411ec91d9a779852f9824; _uetvid=8ea712c0626411ecad8aa94375ab5c09; AWSALBAPP-0=AAAAAAAAAABCjmX/i9b3YPaMqbz/JILdVRryP59LzwWtAfYAx9rxAEc0ofQ/vjLCCJzlHdAn8NPIl3U9HBB8fcaVI9PQhL3lkUvcEnMsDWRTl7X4e6JSBQY4lWbmIanYImFL59/O10ymNAw=; sailthru_pageviews=1; _ga_K7BPDXYMDW=GS1.1.1640097167.2.1.1640097169.58; lastRskxRun=1640097170808; mfSearchActive=true; _ga=GA1.2.475749444.1640094459; _dd_s=rum=0&expire=1640098113030; _pxff_rf=1; _pxff_fp=1; _px2=eyJ1IjoiZGY5ODUyYzAtNjI2YS0xMWVjLWIxNGUtZTU5MjQwODU2NWRlIiwidiI6IjhhNmJiNGY3LTYyNjQtMTFlYy1hNTM4LTZjNDY3OTQzNTk3MiIsInQiOjE1NjE1MDcyMDAwMDAsImgiOiIzNTlkZWEwNTBmMjg5OGFiZDgyMjQ0MTlkOGEyMzk4YmIwY2FiNDk2N2I5MzBlN2U3ZmZmMGU0NjdkNDlhOTM0In0=; _pxde=6e5ecafb1cae7773c58d217fd4bac1cb6a30a3f6b475c83d43700fb65cc19d01:eyJ0aW1lc3RhbXAiOjE2NDAwOTcyMjc0MjUsImZfa2IiOjAsImlwY19pZCI6W119'
}
total = requests.get(url, headers=headers)
headers['cookie'] = headers['cookie']

bs = BeautifulSoup(total.content, "html.parser")
number_text = bs.find("p", {"data-testid": "FilterStatus-search-results"}).get_text()
number_of_products = int(re.findall("\d+", number_text)[0])
number_of_pages = int((number_of_products - 1) / 60) + 1

# 페이지 수
for page in range(2, number_of_pages + 1):
    if "?" in url:
        result = requests.get(f"{url}&page={str(page)}&noOfRecordsPerPage=240", headers=headers)
    else:
        result = requests.get(f"{url}?page={str(page)}&noOfRecordsPerPage=240", headers=headers)
    list_src = str(result.content)
    print(list_src)

    soup = BeautifulSoup(list_src, "html.parser")

    products = soup.find_all("li", {"class": "lister__item"})
    product_links = []
    if len(products) == 0:
        break
    for idx, item in enumerate(products):
        print(str(page) + "," + str(idx) + "th product")
        try:
            photo_links = []
            a = item.find_all("a")[0]
            brand_name = hex_to_str(item.find_all("div", {"class": "lister__item__title"})[0].get_text())
            title = hex_to_str(item.find_all("div", {"class": "lister__item__details"})[0].get_text())
            sizes = list(map(lambda y: y.get_text(), list(filter(lambda x: not hasattr(x.attrs, 'class'),
                                                                 item.find_all("ul", {"class": "sizes"})[0].find_all(
                                                                     "li")))))
            product_row = row_sample.copy()
            product_detail_page = requests.get("https://www.matchesfashion.com" + a.attrs['href'], headers=headers)
            product_detail_src = product_detail_page.content

            soup_detail_page = BeautifulSoup(product_detail_src, "html.parser")
            product_detail_soup = soup_detail_page.find_all("div", {"class": "pdp-grid-accordion"})[0]
            try:
                product_description = \
                product_detail_soup.find_all("div", {"data-testid": "ProductsCarousel-description"})[0].find_all(
                    "span")[0].get_text()
            except IndexError:
                product_description = ""
            product_code = soup_detail_page.find_all("p", {"data-testid": "ProductCode-code"})[0].find_all("strong")[
                0].get_text()
            product_detail_detail = str(
                soup_detail_page.find_all("div", {"data-testid": "ProductsCarousel-detail-bullets"})[0].find_all("ul")[
                    0])
            product_detail_size = str(
                soup_detail_page.find_all("div", {"data-testid": "ProductsCarousel-size-and-fit-bullets"})[0].find_all(
                    "ul")[0])
            product_detail = f"<h3>상품 설명</h3><p>{product_description}</p><h3>상세 정보</h3>{product_detail_detail}<h3>사이즈&핏</h3>{product_detail_size}"
            # price_txt = "".join([
            #     bs_object
            #     for bs_object
            #     in soup_detail_page.find_all("div", {"class": "product-price"})[0]
            #     if isinstance(bs_object, NavigableString)
            # ])
            #
            # price_won_txt, price_euro_txt = price_txt.split('/')
            price_won = soup_detail_page.find_all("span", {"data-testid": "ProductPrice-indicativ-price"})[
                0].get_text().replace(",", "").replace('₩', "").replace("\xa0", "").replace("/", "")
            price_euro = soup_detail_page.find_all("span", {"data-testid": "ProductPrice-billing-price"})[
                0].get_text().replace("\xa0", "").replace(",", "").replace('€', "").replace('\n', "")

            # category_section = soup_detail_page.find_all("ul", {"class": "pdp-viewall"})[0].find_all("li")[1]
            # category = list(map(lambda x: x.get_text(), category_section.find_all("a")))
            # category_string = ",".join(category)

            option_types = ""
            size_string = ""
            try:
                option_types = "size"
                size_string = ",".join(sizes)
            except IndexError as e:
                pass

            product_links.append(a.attrs['href'])

            photos = []
            photo_index = 0
            while True:
                photo_index = photo_index + 1
                if saving_photo is True:
                    img = requests.get(
                        f"https://assetsprx.matchesfashion.com/img/product/{product_code}_{photo_index}_zoom.jpg",
                        headers=headers
                    ).content
                    if not img.startswith(b'RIFF'):
                        break
                    pilimg = Image.open(BytesIO(img))
                    pilimg.convert('RGB').save(
                        f'data/{today}/{information}/{information}_{page}_{idx}_{photo_index}.jpg', 'jpeg')
                elif not os.path.isfile(f"data/{today}/{information}/{information}_{page}_{idx}_{photo_index}.jpg"):
                    break
                photo_links.append(
                    f"https://assetsprx.matchesfashion.com/img/product/{product_code}_{photo_index}_zoom.jpg")
                photos.append(f'{information}_{page}_{idx}_{photo_index}.jpg')

            image_tags = "".join(list(map(lambda x: f"<img src=\"{x}\" />", photo_links)))

            # 상품명
            product_row[2] = f"{brand_name} {title}"
            # 가격
            product_row[3] = str(int(price_won))
            # 대표이미지, 추가이미지
            product_row[7] = f"{information}_{page}_{idx}_1.jpg"
            product_row[8] = ", ".join(photos[1:])
            # 제조사, 브랜드
            product_row[12] = brand_name
            product_row[13] = brand_name
            # 본문
            product_row[9] = f"<div style=\"text-align: center; font-size:16px\">" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMjAxMDFfMTE0/MDAxNjQxMDE1NDM4MTg4.NmS6nb43C0VQXim3oGWX3KFtjk8j4nYnH5sDIxet650g._G6NDb6W7JMLYpCQ0_Do9m1TnzxvQRmGt4gjrDNbrdkg.PNG.c_maru05/23.EU.2%EC%A3%BC.%EA%B4%80%EB%B6%80%EA%B0%80%EC%84%B8%ED%8F%AC%ED%95%A8_(1).png\" />" \
                             f"{str(product_detail)}" \
                             f"{image_tags}" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMjAxMDFfMjky/MDAxNjQxMDE0MTA3OTEx.V4BRICGohcJ3GEV3FeyL43NdXyA9WvY7zBxAxq6v9dsg.NsM_paA_JHQuCT6fvCUaENvRjHoHgABCCFFe7MPQe-4g.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_1.jpg\" />" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMjAxMDFfMjQ3/MDAxNjQxMDE0MTA3OTEy.uEGfU9fGAVlMW8ElD1c_9uXlmqmfoiTaPuuQHtyawvQg.vYLMVVbvNFFrkb50Dac950zsyCuOYI31dIbPGegGii8g.JPEG.c_maru05/%ED%95%B4%EC%99%B8%EB%B0%B0%EC%86%A1%EC%A2%85%ED%95%A9%EC%95%88%EB%82%B4_200623_2.jpg\" />" \
                             f"</div> "
            # category
            # product_row[66] = category_string
            # euro
            product_row[67] = str(int(price_euro))
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
