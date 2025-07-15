import helium as hl
import codecs
from lxml import etree
import time

driver=hl.start_chrome('https://www.dongchedi.com/community/4865/wenda-release')

def to_file(list,path):
    with codecs.open(path+'懂车帝问答'+'.txt','a+','utf-8') as t2:
        for i in list:
            t2.write('\t'.join(i).replace('\r','').replace('\n','')+'\r\n')

def extract():
    url=driver.current_url
    e1=driver.page_source
    e2=etree.HTML(e1)
    e3=e2.xpath('//div[@class="jsx-3811145822 data-wrapper tw-pt-12"]/section')
    z=[]
    for i in e3:
        z0=''.join(i.xpath('.//div[@class="tw-flex tw-items-center"]//div[@class="tw-overflow-hidden tw-flex tw-items-center tw-h-20"]//text()'))#评论人
        z1=''.join(i.xpath('.//div[@class="tw-flex tw-items-center"]//@href'))#评论人网址
        z2=''.join(i.xpath('.//div[@class="tw-flex tw-items-center"]//div[@class="tw-h-16 tw-mt-4 tw-flex tw-items-center"]//text()'))#评论人简介
        z3=''.join(i.xpath('.//div[contains(@class,"jsx-81802501")]//text()'))#内容
        z4=';'.join(i.xpath('.//div[contains(@class,"tw-grid tw-grid-cols-5 2xl:tw-grid-cols-6 tw-gap-8")]//@src'))#图片
        z5=''.join(i.xpath('.//div[@class="jsx-1875074220 right tw-flex tw-items-center tw-flex-none tw-text-12 md:tw-text-14 xl:tw-text-14"]/button[1]//text()'))#回答数
        z6=''.join(i.xpath('.//div[@class="jsx-1875074220 right tw-flex tw-items-center tw-flex-none tw-text-12 md:tw-text-14 xl:tw-text-14"]/button[2]//text()'))#收藏数
        z7=''.join(i.xpath('.//div[@class="jsx-1875074220 tw-flex tw-items-center tw-flex-1 tw-overflow-hidden tw-mr-24 tw-text-12 md:tw-text-14 xl:tw-text-14"]//text()'))#时间

        z.append([url,z0,z1,z2,z3,z4,z5,z6,z7])
    to_file(z,'D:/scrapy/网页/')

def run(url):
    hl.go_to(url)
    hl.scroll_down(4000)
    while 1:
        try:
            extract()
            time.sleep(1)
            hl.click(hl.S('//ul[@class="jsx-1325911405 tw-flex"]/li[last()]/a//i[@class="jsx-1325911405 DCD_Icon icon_into_12 tw-text-14"]'))
            time.sleep(2)
            hl.scroll_down(4000)
            time.sleep(1)
        except Exception as e:
            print(e)
            break

urls=['https://www.dongchedi.com/community/4865/wenda-release']
for url in urls:
    run(url)
