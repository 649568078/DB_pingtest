import xlwt

# 创建一个workbook 设置编码

workbook = xlwt.Workbook(encoding='utf-8')
# 创建一个worksheet
worksheet = workbook.add_sheet('sheet1')
# 写入excel

row1 = ['医科院图书馆数据库', '图书馆主页', '目录查询', '协和搜索', 'Sinomed', 'American Medical Association Journals', 'UPTODATE',
        'ClinicalKey', 'EMBASE', 'JoVE', 'Karger', 'Nature', 'Lippincott Williams & Wilkins Journals',
        'Oxford Journals',
        'Proquest', 'ScienceDirect', 'Scopus', 'Springer', 'Wiley', 'SCIE', '人卫临床知识库', 'Annual Reviews']
row2 = ['数据库链接地址', 'http://www.imicams.ac.cn/', 'http://opac.imicams.ac.cn:8090/opac/search.php',
        'http://discovery.imicams.ac.cn/', 'http://www.sinomed.ac.cn/', 'http://jamanetwork.com/',
        'http://www.uptodate.com/', 'http://www.clinicalkey.com/', 'http://www.embase.com/', 'https://www.jove.com/',
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
# 参数对应 行, 列, 值

for i in range(len(row1)):
    worksheet.write(i, 0, label=row1[i])
for i in range(len(row2)):
    worksheet.write(i, 1, label=row2[i])
for i in range(len(row3)):
    worksheet.write(i, 2, label=row3[i])

# 保存
workbook.save('Excel_test.xls')
