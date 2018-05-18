# coding:utf-8
from etl_certification_docx import module_docx_etl
import traverse_files_name_2_list
import time

start_time = time.time()

file_path = '../file_position/Manuals PD-PP039140-17J3421-20.VCS.rev04.docx'
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

end_time = time.time()
print('cost time: %fs' %(end_time -start_time))