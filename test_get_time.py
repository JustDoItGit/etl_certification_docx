# coding:utf-8
import datetime
import re


def get_time(comment_str):
    # 正则匹配时间的列表
    date_p_list = [ur'19\d{2}-[01]{0,1}\d-[0-3]{0,1}\d|20\d{2}-[01]{0,1}\d-[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d-[01]{0,1}\d-19\d{2}|[0-3]{0,1}\d-[01]{0,1}\d-20\d{2}',
                   ur'19\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d|20\d{2}\.[01]{0,1}\d\.[0-3]{0,1}\d',
                   ur'[0-3]{0,1}\d\.[01]{0,1}\d\.19\d{2}|[0-3]{0,1}\d\.[01]{0,1}\d\.20\d{2}',
                   ur'19\d{2}/[01]{0,1}\d/[0-3]{0,1}\d|20\d{2}/[01]{0,1}\d/[0-3]{0,1}\d',
                   ur'[01]{0,1}\d/[0-3]{0,1}\d/19\d{2}|[01]{0,1}\d/[0-3]{0,1}\d/20\d{2}', ]
    # 正则规则对应的时间格式化列表
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


print get_time('sdf2016.03.04')
print get_time('sdf03.04.2016.')
print get_time('sdf2016-9-12')
print get_time('sdf9-12-2016')
print get_time('sdf9/12/2016')
print get_time('sdf/2016/09/12')
