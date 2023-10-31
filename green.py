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
              "홈나래는 영국 명품백화점의 제품을 구매대행해드리는 업체입니다.\n주문하신 제품의 하자가 있는 경우 교환,반품은 가능합니다만,구입초기의 하자가 아닌\n제품수령후 발생한 하자의 "
              "경우에는 반품비용(왕복배송비)을 부담하셔야 가능함을 알려드립니다.", "010-4783-1346", "", "", "TEST",
              "", "", "", "", "",
              "", "과세상품", "Y", "Y", "03",
              "", "N", "", "택배, 소포, 등기", "유료",
              "20000", "선결제", "", "", "40000",
              "60000", "", "", "", "",
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
information = "glassware" # 띄어쓰기 X
url = "https://www.selfridges.com/KR/en/cat/home-tech/home/dining/glassware/"
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
    'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'cache-control': 'max-age=0',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88',
    'cookie': '__cf_bm=nYXS0r0__DkyBw_1pSJ06QUtnZgL5k4LOu8AeuoDS24-1695380906-0-AUthrsxhN4NjTw+WqxB7l2xekll75g+qx44SOufk+EsoX31ITnST4J71/DjaUtIbbMjsXBj052xSbdpNE+9m887P3uGmm+t/H4NTU75ytM43; mmapi.p.bid=%22prodfracgeu03%22; mmapi.p.srv=%22prodfracgeu03%22; utag_chan={"channel":"","channel_set":"","channel_converted":false,"awc":""}; SF_COUNTRY_LANG=KR_en; WC_SESSION_ESTABLISHED=true; WC_PERSISTENT=6huu9KwLCc6dWEkvOpDaLaUqv%2Fs8BRGiIuBYpuziMho%3D%3B2023-09-22+11%3A08%3A27.973_1695380907906-929839_10052_3271048016%2C-1%2CGBP%2Cf6JFok5eCnAnENE6GlpA7dBrxUWh2cpDqL%2FZEbxq9dnIor2pmaTU5DYMI869eeafUAFTHs7KoK40qvpR3cagUg%3D%3D_10052; WC_AUTHENTICATION_3271048016=3271048016%2C3KqtZ0ma7IQsoeREZUwSxqL9452PXcElSzs8cntG6Zs%3D; WC_ACTIVEPOINTER=-1%2C10052; WC_USERACTIVITY_3271048016=3271048016%2C10052%2Cnull%2Cnull%2C1695380907976%2Cnull%2Cnull%2Cnull%2Cnull%2Cnull%2C1877362032%2Cver_1695380907918%2C9LFpfp1Fi4KVGybG21BrECCl8oCU6La%2FnkGqN1OWvHz2VP43sCyiJgs1kjCC9jNKGD4kUY8j%2B08aLfkZ7ON7NQGpbRMxZA6k%2FsIILFKL9XWSySn9UfrxVNf9cLMtJojDWhVv77sBMzemXS%2FOzU1pw2DBLIYCH3VLdHZrGn%2BZ4%2FmSwuSFaWuJRCTjf%2F7CMOzgDgU2MiGZqZf2UIBiFbSEaFWSofdSt0pUSiZ%2FXe1hTW78jzMgW4LUEtEyYqhvAxcLEhDIXcEPyN2%2Fqtv5YlMv408OhVMIuyYmaZqxL%2FJl17M%3D; _cs_mk=0.2660284284591745_1695380910445; _scid=08002203-6801-4b11-8d22-5f8bf4a3e0fb; _cs_c=0; _gid=GA1.2.1195159920.1695380911; _gcl_au=1.1.227384627.1695380911; _pin_unauth=dWlkPU1HSXhOV0k0TVRVdE1qVTBNaTAwWm1RNExXSTNObUV0WkRrMlpqY3hNRGd5TjJabA; _tt_enable_cookie=1; _ttp=lkd73uqZQ7cvwNg9BugHP8-mH83; __tmbid=sg-1695380910-a77d599b3c4e4f48885fadceee5cc637; _fbp=fb.1.1695380912199.99731237; _sctr=1%7C1695308400000; cf_clearance=HZNkrMGFtWwPB0khw_WSbN42HzyUkj1tjUhT3aaQhuA-1695380994-0-1-50b96a33.f033212b.7b835d3-0.2.1695380994; mmapi.p.pd=%22_ldLeH91pVOXy69iSDSEBGldVRaS40sime0vivYkRU4%3D%7CDAAAAApDH4sIAAAAAAAEAGNh2FwT63xH_MM1Bua0okRGIQZGJwYOz5syjAy7tO1dYnff9ph-ZYZdDJBmAIL_UMDA5pJZlJpcwnhHnBEkDgat2xkYfm5iYgDRUMDoCgCiko7wYQAAAA%3D%3D%22; JSESSIONID=00007OfkZIBOpOnhvVZ16BO1Ryn:-1; utag_main=v_id:018abc93a7ad000dee082911d1350506f006306700bd0$_sn:1$_ss:0$_pn:6%3Bexp-session$_st:1695383148431$ses_id:1695380907949%3Bexp-session$dc_visit:1$dc_event:6%3Bexp-session$dc_region:ap-northeast-1%3Bexp-session; _scid_r=08002203-6801-4b11-8d22-5f8bf4a3e0fb; _uetsid=5d3534a0593811ee973c813686b0b8ca; _uetvid=5d355180593811ee8002af79db1302e4; _cs_id=a6adc1c4-82b5-a574-caca-979984752aaf.1695380910.1.1695381348.1695380910.1.1729544910926; _cs_s=6.5.0.1695383148681; _ga=GA1.1.1777655497.1695380911; _ga_R05V82D63H=GS1.1.1695380911.1.1.1695381349.56.0.0'
}
total = requests.get(url, headers=headers)
print(total.content)
a = BeautifulSoup(total.content, "html.parser").find_all("div", {"class": "plp-listing-load-status"})
number_of_products = int(re.findall("\d+", BeautifulSoup(total.content, "html.parser")
                                    .find_all("div", {"class": "plp-listing-load-status"})[0].get_text())[1])

number_of_pages = int((number_of_products - 1) / 60) + 1
# 페이지 수
for page in range(1, number_of_pages + 1):
    if "?" in url:
        result = requests.get(url + "&pn=" + str(page), headers=headers)
    else:
        result = requests.get(url + "?pn=" + str(page), headers=headers)
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
            title = textbox.find("span", {"class": "c-prod-card__cta-box-description"}).get_text()[:-1]
            brand_name = textbox.find("h5").get_text()
            price_txt = item.find("span", {"class": "c-prod-card__cta-box-price"}).get_text()
            price = price_txt[2:-4].replace(",", "")
            product_row = row_sample.copy()
            product_detail_page = requests.get("https://www.selfridges.com" + a.attrs['href'], headers=headers)
            # product_detail_page = requests.get(
            #     "https://www.selfridges.com/KR/en/cat/zadigvoltaire-sourca-v-neck-cashmere-jumper_R00007777/?previewAttribute=FAUVE")
            product_detail_src = product_detail_page.content

            soup_detail_page = BeautifulSoup(product_detail_src, "html.parser")
            product_detail_soup = soup_detail_page.find_all("article", {"id": "content1"})[0].find("div",
                                                                                                   {
                                                                                                       "class": "c-tabs__copy"})
            product_detail = str(product_detail_soup).replace("<ul>", "").replace("</ul>", "")

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
            product_row[8] = f"{information}_{page}_{idx}_2.jpg, {information}_{page}_{idx}_3.jpg, {information}_{page}_{idx}_4.jpg"
            # 제조사, 브랜드
            product_row[12] = brand_name
            product_row[13] = brand_name

            # 본문
            product_row[9] = f"<div style=\"text-align: center; font-size:16px\">" \
                             f"<img src=\"https://blogfiles.pstatic.net/MjAyMDExMTZfMjc4/MDAxNjA1NDg1MzU1NDYw.9WNaxEFOEZjCTOrlwwyrHo3zD5SVX2kyJT1KhnMiKTIg.hydsAsmKPWx3Oz7JtoWV3R-RtgAN2oGnO_GWG2b1JAsg.PNG.c_maru05/19.EU.1%EC%A3%BC.%EA%B4%80%EB%B6%80%EA%B0%80%EC%84%B8%ED%8F%AC%ED%95%A8.png\" />" \
                             f"{str(product_detail)}" \
                             f"<img src=\"http:{photo_link}\" />" \
                             f"<img src=\"http:{photo_link2}\" />" \
                             f"<img src=\"http:{photo_link3}\" />" \
                             f"<img src=\"http:{photo_link4}\" />" \
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