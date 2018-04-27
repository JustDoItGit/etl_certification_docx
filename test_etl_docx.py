# coding:utf-8
from etl_certification_docx import module_docx_etl
import traverse_files_name_2_list

file_path = '../file_position/VCS-02590-A010-006_Rev3 签字终版.docx'
result = module_docx_etl.etl_docx(file_path)
print result


'''
file_name_l = traverse_files_name_2_list.file_name_list('../file_position/GW 121-2500 IEC IIIB Sinoma59.5 HH90/',
                                                        'docx')
print(file_name_l)
import time

start_time = time.time()
for i in file_name_l:
    path = '../file_position/GW 121-2500 IEC IIIB Sinoma59.5 HH90/' + i
    l = module_docx_etl.etl_docx(path)
    print(l)
    print i
end_time = time.time()
print('cost time: %f' % (end_time - start_time))
'''