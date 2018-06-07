# coding:utf-8
'''
version:v1.0
auth:cainiao
question:时间匹配函数待优化
'''
# 导入读取docx的包
import zipfile
import xml.etree.ElementTree
import re
import datetime

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
        # 无法解析的docx文档
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
                if target_str == list_3[t][tr][td].upper()[:len(target_str)]:
                    result_list.append({'t': t, 'tr': tr, 'td': td})
    if len(result_list) == 1:
        # 返回内容所在的下标
        # 例如 {'t':10, 'td': 1, 'tr': 1}
        return result_list[0]
    else:
        # 文件内缺少关键参数，关键参数数量不对
        return 66668


# 提取issued_by（发布人）
def get_issued_by(list_3):
    issued_position = target_positon(list_3, 'ISSUED BY')
    if issued_position == 66668:
        return 66668
    issued_by = list_3[issued_position['t']][issued_position['tr']][issued_position['td']]
    issued_people_sp = issued_by.split(':')
    if len(issued_people_sp) == 2:
        issued_by_people = issued_people_sp[1].strip()
        if issued_by_people != '':
            return issued_by_people
        else:
            return ''
    else:
        return ''


# 提取verified_by（审核人）
def get_verified_by(list_3):
    verified_position = target_positon(list_3, 'VERIFIED BY')
    if verified_position == 66668:
        return 66668
    verified_by = list_3[verified_position['t']][verified_position['tr']][verified_position['td']]
    verified_people_sp = verified_by.split(':')
    if len(verified_people_sp) == 2:
        verified_by_people = verified_people_sp[1].strip()
        if verified_by_people != '':
            return verified_by_people
        else:
            return ''
    else:
        return ''


# 提取description
def get_description(list_3):
    description_position = target_positon(list_3, 'VERIFICATION COMMENTS')
    # 判断是否找到关键词
    if description_position == 66668:
        return 66668
    description_list = []
    for row in list_3[description_position['t']]:
        if len(row) >= 4 and row[1][:3] == 'DNV':
            description_list.append(row[1])
        elif len(row) >= 4 and row[1][:17] == 'Reviewed document':
            description_list.append(row[1])
    return description_list


# 提取日期，并格式化为2017-12-12的格式
def get_time(comment_str):
    # 正则匹配时间的列表
    date_p_list = [ur'19\d{2}-[01]{0,1}\d-[0-3]{0,1}\d|20\d{2}-[01]{0,1}\d-[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d-[01]{0,1}\d-19\d{2}|[0-3]{0,1}\d-[01]{0,1}\d-20\d{2}',
                   ur'19\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d|20\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d\.[01]{0,1}\d\.19\d{2}|[0-3]{0,1}\d\.[01]{0,1}\d\.20\d{2}',
                   ur'19\d{2}/[01]{0,1}\d/[0-3]{0,1}\d|20\d{2}/[01]{0,1}\d/[0-3]{0,1}\d',
                   ur'[01]{0,1}\d/[0-3]{0,1}\d/19\d{2}|[01]{0,1}\d/[0-3]{0,1}\d/20\d{2}',
                   ]
    # 正则规则对应的时间格式化列表 DNV GL 05-08-2016
    date_str_list = ['%Y-%m-%d', '%d-%m-%Y', '%Y.%m.%d', '%d.%m.%Y', '%Y/%m/%d', '%m/%d/%Y', ]
    for i, p in enumerate(date_p_list):
        date_p = re.compile(p)
        date = date_p.findall(comment_str)
        if date:
            t_str = date[0]
            try:
                d = datetime.datetime.strptime(t_str, date_str_list[i])
                return d
            except:
                continue
    return ''


# 计算两个日期时间差
def compute_days(start_time, end_time):
    days_datetime = end_time - start_time
    days = days_datetime.days
    if days >= 0:
        return days
    else:
        return 0


def compute_loop_count(comment_content):
    '''
    计算loop_count（评审意见总轮次数）
    start_time（评审意见起始时间）
    end_time（评审意见结束时间）
    duration（评审意见持续天数）
    comment_content: 评审内容
    '''
    loop_count = 0
    start_time = ''
    end_time = ''
    duration = 0
    p = re.compile(ur'DNVGL|DNV GL|DNV', re.I)
    p_find = p.findall(comment_content)
    if p_find:
        dnv_list = p.split(comment_content)[1:]
        dnv_list_new = [i for i in dnv_list if whether_contain_date(i[:30])]
        if len(dnv_list_new) == 0:
            start_time = ''
            end_time = ''
        else:
            loop_count = len(dnv_list_new) - 1
            start_time = get_time(dnv_list_new[0])
            end_time = get_time(dnv_list_new[-1])
        if start_time != '' and end_time != '':
            duration = compute_days(start_time, end_time)
    if start_time == '':
        start_time = None
    if end_time == '':
        end_time = None
    return {
        # 'comment_content': line_feed_comment_content(comment_content),  # 换行
        'start_time': start_time,
        'end_time': end_time,
        'duration': duration,
        'loop_count': loop_count,
    }


# 对comment_content评审意见内容换行
def line_feed_comment_content(comment_content):
    p = re.compile(ur'DNVGL|Goldwind|DNV GL|DNV|gold wind', re.I)
    group_list = p.findall(comment_content)
    line_feed = comment_content
    if group_list:
        say_list = p.split(comment_content)
        say_list.remove(say_list[0])
        if len(group_list) == len(say_list):
            # 规整对话内容
            say_clean = ''
            for i, val in enumerate(group_list):
                # \r\n\r\n用于段落间换行
                say_clean = say_clean + val + ' ' + say_list[i] + '\r\n\r\n'
            line_feed = say_clean
    return line_feed


# 用于判断每个DNV后面是否跟随时间
def whether_contain_date(comment_str):
    # 正则匹配时间的列表
    date_p_list = [ur'19\d{2}-[01]{0,1}\d-[0-3]{0,1}\d|20\d{2}-[01]{0,1}\d-[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d-[01]{0,1}\d-19\d{2}|[0-3]{0,1}\d-[01]{0,1}\d-20\d{2}',
                   ur'19\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d|20\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d\.[01]{0,1}\d\.19\d{2}|[0-3]{0,1}\d\.[01]{0,1}\d\.20\d{2}',
                   ur'19\d{2}/[01]{0,1}\d/[0-3]{0,1}\d|20\d{2}/[01]{0,1}\d/[0-3]{0,1}\d',
                   ur'[01]{0,1}\d/[0-3]{0,1}\d/19\d{2}|[01]{0,1}\d/[0-3]{0,1}\d/20\d{2}',
                   ]
    for i, p in enumerate(date_p_list):
        date_p = re.compile(p)
        date = date_p.findall(comment_str)
        if date:
            return True
    return False


# main etl主函数
def etl_docx(file_path):
    # 解析docx
    tree = make_docx_tree(file_path)
    list_3 = []
    # 判断解析成功结构化放入3层列表
    if tree == 66666 or tree == 66667:
        return tree, 'Fail'
    else:
        list_3 = make_list_3dim(tree)
    issued_by = get_issued_by(list_3)
    verified_by = get_verified_by(list_3)
    description = get_description(list_3)
    if issued_by == 66668 or verified_by == 66668 or description == 66668:
        return 66668, 'Fail'
    comment_all = []
    for i, comment_content in enumerate(description):
        comment_ = compute_loop_count(comment_content)
        comment_all.append(comment_)
    if comment_all == []:
        comment_all = [{
            'comment_content': '',  # 换行
            'start_time': None,
            'end_time': None,
            'duration': 0,
            'loop_count': 0,
        }]
    return 60000, {
        'issued_by': issued_by,
        'verified_by': verified_by,
        'comment_all': comment_all,
    }

# print(etl_docx('load validation_VCS-00985-A056-001_ rev 2.docx'))
