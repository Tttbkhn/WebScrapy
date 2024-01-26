from pathlib import Path
from urllib import request
from io import BytesIO
from urllib.request import urlopen

import scrapy
import xlsxwriter
import io
from PIL import Image


class FanhouseSpider(scrapy.Spider):
    name = "fanhouse"
    start_urls = [
        "https://fanhouse.waca.ec/en/product/all/newest",
    ]

    def parse(self, response):
        workbook = xlsxwriter.Workbook("fanhouse.xlsx")
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": 1})
        worksheet.write("A1", "No.", bold)
        worksheet.write("B1", "Title", bold)
        worksheet.write("C1", "Link", bold)
        worksheet.write("D1", "Price", bold)
        worksheet.write("E1", "Image", bold)
        worksheet.write("F1", "Image2", bold)

        row = 1
        col = 0
        for item in response.css("ul.pt_grid_list li"):
            # yield {
            #     "title": item.css("span::attr(title)").get(),
            #     "price": item.css(".js_origin_price::text").get(),
            #     "link": item.css("a::attr(href)").get(),
            #     "image": item.css("span::attr(style)").re(
            #         r"background-image: url\((.*?)\)"
            #     )[0],
            # }
            image_url = item.css("span::attr(style)").re(
                r"background-image: url\((.*?)\)"
            )[0]
            image_data = BytesIO(urlopen(image_url).read())
            image = Image.open(image_data)
            maxImageWidth = 0
            imageWidth, imageHeight = image.size
            worksheet.write_number(row, col, row)
            worksheet.write_string(row, col + 1, item.css("span::attr(title)").get())
            worksheet.write_url(row, col + 2, item.css("a::attr(href)").get())
            worksheet.write_string(
                row, col + 3, item.css(".js_origin_price::text").get()
            )
            worksheet.write_string(row, col + 4, image_url)
            cell_position = "F{}".format(row + 1)
            worksheet.insert_image(
                cell_position,
                image_url,
                {
                    "image_data": image_data,
                    "x_scale": 0.5,
                    "y_scale": 0.5,
                    "object_position": 1,
                },
            )
            if imageWidth > maxImageWidth:
                maxImageWidth = imageWidth
            worksheet.set_row(row, imageHeight / 2)
            row += 1
        worksheet.set_column("F:F", maxImageWidth / 2)
        worksheet.autofit()
        workbook.close()
