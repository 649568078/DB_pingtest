import requests
from selenium import webdriver
import time
import xlrd
import xlwt
import os


def get_screen():
    drivers = webdriver.Chrome()
    drivers.set_page_load_timeout(15)

    for i in row2:
        try:
            drivers.maximize_window()
            drivers.get(i)
            try:
                time.sleep(3)
                # 获取响应码
                r = requests.get(i, timeout=15)
                status_code = r.status_code
                if int(status_code) == 200:
                    print(i, "网页正常")
                else:
                    print(i, status_code)
                re_i = i.replace('/','').replace(':','').replace('\\','')
                print(re_i)
                picture_url = drivers.get_screenshot_as_file(os.getcwd() + "\\" + dir_name + "\\" + re_i + ".png")
                print(str(picture_url) + "截图成功")
            except BaseException as msg:
                print(msg)
        except:
            print("截图失败，可能被墙了")
    drivers.close()


def make_dir():
    # 获取当前系统时间
    global dir_name
    dir_name = time.strftime('%Y_%m_%d %H_%M_%S')
    # 执行判断
    isExists = os.path.exists(dir_name)
    # 如果不存在则创建一个以当前系统时间为名字的目录
    if not isExists:
        os.makedirs(dir_name)
        print(dir_name + ' 创建成功')
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(dir_name + ' 目录已存在')


def make_sample_excle():
    row1 = ['医科院图书馆数据库', '图书馆主页', '目录查询', '协和搜索', 'Sinomed', 'American Medical Association Journals', 'UPTODATE',
            'ClinicalKey', 'EMBASE', 'JoVE', 'Karger', 'Nature', 'Lippincott Williams & Wilkins Journals',
            'Oxford Journals',
            'Proquest', 'ScienceDirect', 'Scopus', 'Springer', 'Wiley', 'SCIE', '人卫临床知识库', 'Annual Reviews']

    global row2
    row2 = ['数据库链接地址', 'http://www.imicams.ac.cn/', 'http://opac.imicams.ac.cn:8090/opac/search.php',
            'http://discovery.imicams.ac.cn/', 'http://www.sinomed.ac.cn/', 'http://jamanetwork.com/',
            'http://www.uptodate.com/', 'http://www.clinicalkey.com/', 'http://www.embase.com/',
            'https://www.jove.com/',
            'http://www.karger.com/', 'http://www.nature.com/',
            'http://ovidsp.ovid.com/ovidweb.cgi?T=JS&NEWS=n&CSC=Y&PAGE=main&D=yrovft', 'http://www.oxfordjournals.org/',
            'http://search.proquest.com/', 'http://www.sciencedirect.com/', 'http://www.scopus.com/',
            'http://link.springer.com/', 'http://onlinelibrary.wiley.com/',
            'http://webofscience.com/?DestApp=WOS&editions=SCI', 'http://medbooks.ipmph.com/',
            'http://arjournals.annualreviews.org/action/showJournals']

    row3 = ['测试内容', '能否正常显示（未开通Shibboleth）', '能否检索（未开通Shibboleth）', '能否检索（未开通Shibboleth）', '能否检索（公共帐号：FnCoV 密码：666666）',
            '是否可下载全文（全库订购）（未开通Shibboleth）', '是否可检索和浏览结果（未开通Shibboleth）', '是否可下载全文（全库订购）', '是否可检索', '是否可浏览医学专辑视频',
            '是否可下载全文（全库订购）', '是否可下载医学期刊全文', '是否可下载全文（全库订购）', '是否可下载医学期刊全文', '是否检索和查看摘要', '是否可下载医学期刊全文', '是否可检索',
            '是否可下载医学期刊全文', '是否可下载医学期刊全文', '是否可检索', '是否可浏览图书章节（未开通Shibboleth）', '是否可下载医学期刊全文']

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('sheet1')
    # 写入excel
     # 参数对应 行, 列, 值
    for i in range(len(row1)):
        worksheet.write(i, 0, label=row1[i])
    for i in range(len(row2)):
        worksheet.write(i, 1, label=row2[i])
    for i in range(len(row3)):
        worksheet.write(i, 2, label=row3[i])
    # 保存
    workbook.save(os.getcwd() + "\\" + dir_name + "\\" + 'Excel_test.xls')


def main():
    # 创建文件夹
    make_dir()
    make_sample_excle()
    get_screen()

if __name__ == '__main__':
    main()


