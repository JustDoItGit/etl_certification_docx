# coding:utf-8
# 导入读取docx的包
import zipfile
import xml.etree.ElementTree
import time

# 用于读取docx的相关配置
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'


# 将传入的docx文档解析成tree结构
def make_docx_tree(file_path):
    try:
        with zipfile.ZipFile(file_path) as docx:
            tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))
        return tree
    except:
        return {'error': '无法解析'}


# 将内容放到到三维列表里面
def make_list_3dim(tree):
    list_3dim = []
    t = 0
    for table in tree.iter(TABLE):
        list_3dim.append([])
        tr = 0
        for row in table.iter(ROW):
            list_3dim[t].append([])
            td = 0
            for cell in row.iter(CELL):
                text = (''.join(node.text for node in cell.iter(TEXT)))
                list_3dim[t][tr].append(text.strip())
                # print('t=%d,tr=%d,td=%d' % (t, tr, td))
                # print(text)
                td += 1
            tr += 1
        t += 1
    return list_3dim


# 找到提取文档的目标内容起点所在的表格位置
def target_positon(list_3, target_str):
    result_list = []
    for t in range(len(list_3)):
        for tr in range(len(list_3[t])):
            for td in range(len(list_3[t][tr])):
                if target_str in list_3[t][tr][td].upper():
                    result_list.append({'t': t, 'tr': tr, 'td': td})
    if len(result_list) == 1:
        # 返回内容所在的下标
        # 例如 {'t':10, 'td': 1, 'tr': 1}
        return result_list[0]
    else:
        # 文件内缺少关键参数
        return 63004


# 提取issued_by（发布人）
def get_issued_by(list_3):
    issued_position = target_positon(list_3, 'ISSUED BY')
    print(issued_position)
    issued_by = list_3[issued_position['td']][issued_position['tr']][issued_position['t']]
    issued_people_sp = issued_by.split(':')
    if len(issued_people_sp) == 2:
        issued_by_people = issued_people_sp[1].strip()
        if issued_by_people != '':
            return issued_by_people
        else:
            return 63004, 'issued by error'
    else:
        return 63004, 'issued by error'


# 提取verified_by（审核人）
def get_verified_by(list_3):
    verified_position = target_positon(list_3, 'VERIFIED BY')
    print(verified_position)
    verified_by = list_3[verified_position['td']][verified_position['tr']][verified_position['t']]
    verified_people_sp = verified_by.split(':')
    if len(verified_people_sp) == 2:
        verified_by_people = verified_people_sp[1].strip()
        if verified_by_people != '':
            return verified_by_people
        else:
            return 63004, 'verified by error'
    else:
        return 63004, 'verified by error'


# 提取description


# main etl主函数
def etl_docx(file_path):
    start_time = time.time()
    tree = make_docx_tree(file_path)
    list_3 = []
    if tree != {'error': '无法解析'}:
        list_3 = make_list_3dim(tree)
    else:
        print('error')
    issued_by = get_issued_by(list_3)
    print(issued_by)

    verified_by = get_verified_by(list_3)
    print(verified_by)
    end_time = time.time()
    print('cost time: %f s' % (end_time - start_time))


etl_docx('../file_position/VCS-02590-A010-006_Rev3 签字终版.docx')
etl_docx('../file_position/VCS-DE-DNVGL-SE-0074-00802-015_rev.2.docx')
etl_docx('../file_position/VCS-03037-A010-006_Rev_00-Baseframe-reply20171012.docx')
# print(etl_docx(''))
