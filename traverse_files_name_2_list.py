# coding:utf-8
import os


# 遍历file_dir文件夹下的所有 file_tpye 文件 放到列表里
# 例如提取file_dir 下的所有 *.docx 文件
def file_name_list(file_dir_path, file_tpye):
    '''
    file_dir_path: 目标文件路径
    file_type: 文件类型 后缀
    '''
    for root, dirs, files in os.walk(file_dir_path):
        # print(root)  # 当前目录路径
        # print(dirs)  # 当前路径下所有子目录
        # print(files)  # 当前路径下所有非目录子文件
        '''
        file_name_l = [i for i in files if i.split('.')[-1] == file_tpye]
        # 处理文件名含有空格的文件
        for i, val in enumerate(file_name_l):
            print(i, val)
            if ' ' in val:
                os.system('mv ' + file_dir_path + '"' + val + '"' + ' ' + file_dir_path + val.replace(' ', '_'))
                file_name_l[i] = val.replace(' ', '_')
        '''
        return [i for i in files if i.split('.')[-1] == file_tpye]
