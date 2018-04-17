# coding:utf-8
# 导入读取docx的包
import zipfile
import xml.etree.ElementTree

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
                list_3dim[t][tr].append(text)
                print('t=%d,tr=%d,td=%d' % (t, tr, td))
                print(text)
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


def etl_docx(file_path):
    tree = make_docx_tree(file_path)
    if tree != {'error': '无法解析'}:
        return make_list_3dim(tree)
    else:
        print('error')


print(etl_docx('../file_position/VCS-02590-A010-006_Rev3 签字终版.docx'))
print(etl_docx(''))
