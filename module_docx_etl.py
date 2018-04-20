# coding:utf-8
# 导入读取docx的包
import zipfile
import xml.etree.ElementTree

import re
import datetime
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
    if file_path[-4:] != 'docx':
        # 无法解析，文件后缀不为docx
        return 66666
    try:
        with zipfile.ZipFile(file_path) as docx:
            tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))
        return tree
    except Exception, e:
        # 无法解析
        return 66667


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
        # 文件内缺少关键参数，关键参数数量不对
        return 63004, '文件内缺少关键参数：%s' % target_str


# 提取issued_by（发布人）
def get_issued_by(list_3):
    issued_position = target_positon(list_3, 'ISSUED BY')
    # print(issued_position)
    issued_by = list_3[issued_position['t']][issued_position['tr']][issued_position['td']]
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
    # print(verified_position)
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


# 上面两个函数有时间可以合并成一个

# 提取description
def get_description(list_3):
    description_position = target_positon(list_3, 'VERIFICATION COMMENTS')
    if description_position == 63004:
        return 63004
    description_list = []
    for row in list_3[description_position['t']]:
        # print(len(row))
        if len(row) >= 4 and row[1][:3] == 'DNV':
            # print(row[1])
            # print(row[1])
            description_list.append(row[1])
    return description_list


# 提取日期，并格式化为2017-12-12的格式
def get_time(comment_str):
    date_p_list = ['\d{4}-\d{2}-\d{2}']
    date_str_list = ['%Y-%m-%d']
    for i, p in enumerate(date_p_list):
        date_p = re.compile(p)
        date = date_p.findall(comment_str)
        if date:
            t_str = date[0]
            try:
                d = datetime.datetime.strptime(t_str, date_str_list[i]).strftime('%Y-%m-%d')
                return d
            except:
                return ''
    return ''


# 计算两个日期时间差
def compute_days(start_time, end_time):
    days_datetime = datetime.datetime.strptime(end_time, '%Y-%m-%d') - \
                    datetime.datetime.strptime(start_time, '%Y-%m-%d')
    days = days_datetime.days
    if days >= 0:
        return days
    else:
        return 0


def compute_loop_count(comment):
    '''
    计算loop_count（评审意见总轮次数）
    start_time（评审意见起始时间）
    end_time（评审意见结束时间）
    duration（评审意见持续天数）
    '''
    p = re.compile('DNVGL|DNV GL', re.I)
    dnv_list = p.split(comment)[1:]
    loop_count = len(dnv_list) - 1
    start_time = get_time(dnv_list[0])
    end_time = get_time(dnv_list[-1])
    if start_time != '' and end_time != '':
        duration = compute_days(start_time, end_time)
    else:
        duration = 0
    return {
        'start_time': start_time,
        'end_time': end_time,
        'duration': duration,
        'loop_count': loop_count,
    }


# main etl主函数
def etl_docx(file_path):
    start_time = time.time()
    tree = make_docx_tree(file_path)
    list_3 = []
    if tree == 66666 or tree == 66667:
        print(tree)
        return tree
    else:
        list_3 = make_list_3dim(tree)
    result_dic = {}
    issued_by = get_issued_by(list_3)
    result_dic['issued_by'] = issued_by

    verified_by = get_verified_by(list_3)
    result_dic['verified_by'] = verified_by
    # print(verified_by)

    description = get_description(list_3)
    # print(description[1])
    comment_all = []
    for i, comment in enumerate(description):
        comment_ = compute_loop_count(comment)
        # comment_['comment_id'] = i
        # print(comment_)
        comment_all.append(comment_)

    result_dic['comment_all'] = comment_all
    print(result_dic)
    end_time = time.time()
    print('cost time: %f s' % (end_time - start_time))


etl_docx('../file_position/VCS-02590-A010-006_Rev3 签字终版.docx')


# etl_docx('../file_position/VCS-02590-A010-006_Rev3 签字终版.doc')


# etl_docx('../file_position/VCS-DE-DNVGL-SE-0074-00802-015_rev.2.docx')
# etl_docx('../file_position/VCS-03037-A010-006_Rev_00-Baseframe-reply20171012.docx')


# print(etl_docx(''))
def go():
    return 1, '完成'


print(go())
