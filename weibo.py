import datetime
import random
import re
import time

import requests
import xlwt
from lxml import etree


class WeiBoText(object):

    def __init__(self, keywords):
        self.url = "https://s.weibo.com/weibo?q=%s"%keywords
        self.keywords = keywords
        self.start_request()

    def getHeaders(self):
        # user_agent列表
        user_agent_list = [
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.71 Safari/537.1 LBBROWSER',
            'Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; QQDownload 732; .NET4.0C; .NET4.0E)',
            'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.84 Safari/535.11 SE 2.X MetaSr 1.0',
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.3.4000 Chrome/30.0.1599.101 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.122 UBrowser/4.0.3214.0 Safari/537.36'
        ]
        headers = {
            'User-Agent': random.choice(user_agent_list),
        }
        return headers

    # 发起搜索请求
    def start_request(self):
        try:
            # 发起请求获取返回响应信息
            response = requests.get(url=self.url, headers=self.getHeaders())
            # 设置编码
            response.encoding = "utf-8"
            content = etree.HTML(response.text)
            card_wrap_list = content.xpath('//*[@id="pl_feedlist_index"]/div[1]/div[@action-type="feed_list_item"]')
            if len(card_wrap_list) > 0:
                self.save_data(card_wrap_list)
            else:
                print("没有抓取到数据,请重新输入搜索关键字！")
        except Exception as e:
            print("搜索请求异常："+e)

    # 处理数据
    def save_data(self, data_list):
        data_lis = []
        for item in data_list:
            data = {}
            href = item.xpath('div[@class="card"]/div[@class="card-feed"]/div[@class="avator"]/a/@href')
            if len(href) > 0:
                href = href[0]
            else:
                continue
            # 博主id
            data['bz_id'] = re.findall("//weibo.com/(.*?)\?refer_flag=1001030103_", href)[0]
            # 发布内容
            content_text = item.xpath('div[@class="card"]/div[@class="card-feed"]/div[@class="content"]/p[@class="txt"]/text()')
            data['content_text'] = ''.join(content_text).replace('\n', '').replace('\r', '').replace(' ', '')
            # 发布时间
            date_time = item.xpath('div[@class="card"]/div[@class="card-feed"]/div[@class="content"]/p[@class="from"]/a/text()')
            if len(date_time) > 0:
                date_time = date_time[0].replace('\n', '').replace('\r', '').replace(' ', '')
                data['date_time'] = self.create_time(date_time)
            else:
                continue
            # 转发数
            data['forward_num'] = item.xpath('div[@class="card"]/div[@class="card-act"]/ul/li[2]/a/text()')[0].replace('\n', '').replace('\r', '').replace(' 转发 ', '')
            if len(data['forward_num'].strip()) == 0:
                data['forward_num'] = '0'
            # 评论数
            data['conment_num'] = item.xpath('div[@class="card"]/div[@class="card-act"]/ul/li[3]/a/text()')[0].replace('\n', '').replace('\r', '').replace('评论 ', '')
            if len(data['conment_num'].strip()) == 0:
                data['conment_num'] = '0'
            # 点赞数
            give_num = item.xpath('div[@class="card"]/div[@class="card-act"]/ul/li[4]/a/em/text()')
            if len(give_num) > 0:
                data['give_num'] = give_num[0].strip()
            else:
                data['give_num'] = '0'
            data_lis.append(data)
        self.write_excel(data_lis)

    # 发布时间处理
    def create_time(self, data):
        date_time = ""
        if data.find("今天") == 0:
            date_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
            date_time = date_time + " " + data.replace("今天", "")
        elif data.find("分钟前") > 0:
            m = data[0:data.find('分钟前')]
            date_time = (datetime.datetime.now() - datetime.timedelta(minutes=int(m))).strftime("%Y-%m-%d %H:%M")
        elif data.find("年") == -1 and data.find("月") > 0:
            date_time = time.strftime('%Y', time.localtime(time.time()))
            date_time = date_time + "-" + data.replace("月", "-").replace("日", "")
            date_time = date_time.replace(" ", "")
            string_lis = list(date_time)
            string_lis.insert(10, " ")
            date_time = ''.join(string_lis)
        elif data.find("年") > 0:
            date_time = data.replace("年", "-").replace("月", "-").replace("日", "")
            string_lis = list(date_time)
            string_lis.insert(10, " ")
            date_time = ''.join(string_lis)
        else:
            date_time = time.strftime('%Y-%m-%d %H:%M', time.localtime(time.time()))
        return date_time[0:17]

    # 写入到 excel
    def write_excel(self, data_list):
        print(data_list)
        excelpath = self.keywords + ".xls"  # 新建excel文件
        workbook = xlwt.Workbook(encoding='utf-8')  # 写入excel文件
        sheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)  # 新增一个sheet工作表
        headlist = [u'博主ID', u'微博正文', u'评论数', u'转发数', u'点赞数', u'发布时间']  # 写入数据头
        row = 0
        col = 0
        for head in headlist:
            sheet.write(row, col, head)
            col += 1
        for i in range(0, len(data_list)):
            sheet.write(i + 1, 0, data_list[i]["bz_id"])
            sheet.write(i + 1, 1, data_list[i]["content_text"])
            sheet.write(i + 1, 2, data_list[i]["conment_num"])
            sheet.write(i + 1, 3, data_list[i]["forward_num"])
            sheet.write(i + 1, 4, data_list[i]["give_num"])
            sheet.write(i + 1, 5, data_list[i]["date_time"])
        workbook.save(excelpath)


if __name__ == '__main__':
    keywords = input("请输入搜索关键字:")
    # 清除空左右两边空格
    if len(keywords.strip()) == 0:
        print("请输入搜索关键字")
    # 实例化类
    WeiBoText(keywords)
    exit()
