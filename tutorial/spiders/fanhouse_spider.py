from pathlib import Path
from urllib import request
from io import BytesIO
from urllib.request import urlopen

import scrapy
import xlsxwriter
import io
from PIL import Image
import logging
import json


class FanhouseSpider(scrapy.Spider):
    name = "fanhouse"
    start_urls = [
        "https://fanhouse.waca.ec/en/product/all/newest",
    ]
    page = 1
    request_body = None

    def create_req_body(self, token, id_number_ancestor):
        self.page += 1
        self.request_body = {
            "_token": token,
            "page": self.page,
            "pagePre": 20,
            "sort": "newest",
            "componentId": id_number_ancestor,
            "method": ["content"],
            "mode": "other",
            "type": "product",
            "value": "all",
        }

    def parse(self, response):
        # workbook = xlsxwriter.Workbook("fanhouse.xlsx")
        # worksheet = workbook.add_worksheet()

        # bold = workbook.add_format({"bold": True})
        # text_wrap = workbook.add_format({"text_wrap": True})
        # worksheet.write("A1", "No.", bold)
        # worksheet.write("B1", "Title", bold)
        # worksheet.write("C1", "Link", bold)
        # worksheet.write("D1", "Price", bold)
        # worksheet.write("E1", "Image", bold)
        first_link = ""
        token = None
        id_number_ancestor = None
        # row = 1
        # col = 0
        for index, item in enumerate(response.css("ul.pt_grid_list li")):
            image_url = item.css("span::attr(style)").re(
                r"background-image: url\((.*?)\)"
            )[0]
            title = item.css("span::attr(title)").get()
            link = item.css("a::attr(href)").get()
            if index == 0:
                first_link = link
            json_response = {
                "image_url": image_url,
                "title": title,
                "link": link,
                "price": item.css(".js_origin_price::text").get(),
                "sale_price": item.css(".js_pt_sale::text").get(),
            }

            yield json_response

            # image_url = item.css("span::attr(style)").re(
            #     r"background-image: url\((.*?)\)"
            # )[0]
            # image_data = BytesIO(urlopen(image_url).read())
            # image = Image.open(image_data)
            # maxImageWidth = 0
            # imageWidth, imageHeight = image.size
            # worksheet.write_number(row, col, row)
            # worksheet.write_string(
            #     row, col + 1, item.css("span::attr(title)").get(), text_wrap
            # )
            # worksheet.write_url(
            #     row, col + 2, item.css("a::attr(href)").get(), text_wrap
            # )
            # worksheet.write_string(
            #     row, col + 3, item.css(".js_origin_price::text").get()
            # )
            # cell_position = "E{}".format(row + 1)
            # worksheet.insert_image(
            #     cell_position,
            #     image_url,
            #     {
            #         "image_data": image_data,
            #         "x_scale": 0.7,
            #         "y_scale": 0.7,
            #         "object_position": 1,
            #     },
            # )
            # if imageWidth > maxImageWidth:
            #     maxImageWidth = imageWidth
            # worksheet.set_row(row, imageHeight / 1.8)
            # row += 1
        # worksheet.set_column("E:E", maxImageWidth / 8)
        # worksheet.set_column("B:B", 40)
        # worksheet.set_column("C:C", 30)
        # worksheet.set_column("D:D", 15)
        # workbook.close()
        first_item = response.xpath("//a[@href=$val]", val=first_link).get()
        if first_item is not None:
            item_selector = response.xpath("//a[@href=$val]", val=first_link)[0]
            id_number_ancestor = item_selector.xpath("ancestor::*[@id]/@id").get()
            logging.info(id_number_ancestor)

        get_token = response.xpath(
            "//script[contains(text(),'fetchthemeproduct')]"
        ).get()
        if get_token is not None:
            token_selector = response.xpath(
                "//script[contains(text(),'fetchthemeproduct')]"
            )[0]

            token_regex = token_selector.re(r"csrf_token.*u0022(.*?)\\u0022}")
            if token_regex is not None:
                token = token_regex[0]
                logging.info(token)

        if token is not None and id_number_ancestor is not None:
            self.create_req_body(token, id_number_ancestor)
            yield scrapy.Request(
                "https://fanhouse.waca.ec/en/fetchthemeproducthtml",
                method="POST",
                body=json.dumps(self.request_body),
                callback=self.fetchMoreProduct,
                headers={"Content-type": "application/json; charset=UTF-8"},
                cb_kwargs=dict(token=token, id_number_ancestor=id_number_ancestor),
            )

    def fetchMoreProduct(self, response, token, id_number_ancestor):
        logging.info(response.status_code)
        json_response = response.json()
        raw_product_html = json_response["productHtml"]
        hasMorePage = json_response["hasMorePage"]
        product_html = raw_product_html.replace('\\"', "\\")
        product_html = product_html.replace("\\n", "")
        for item in product_html.css("li"):
            image_url = item.css("span::attr(style)").re(
                r"background-image: url\((.*?)\)"
            )[0]
            title = item.css("span::attr(title)").get()
            link = item.css("a::attr(href)").get()
            json_response = {
                "image_url": image_url,
                "title": title,
                "link": link,
                "price": item.css(".js_origin_price::text").get(),
                "sale_price": item.css(".js_pt_sale::text").get(),
            }

            yield json_response

        if hasMorePage:
            self.create_req_body(token, id_number_ancestor)
            yield scrapy.Request(
                "https://fanhouse.waca.ec/en/fetchthemeproducthtml",
                method="POST",
                body=json.dumps(self.request_body),
                callback=self.fetchMoreProduct,
                headers={"Content-type": "application/json; charset=UTF-8"},
                cb_kwargs=dict(token=token, id_number_ancestor=id_number_ancestor),
            )
