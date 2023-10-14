import argparse
import os
import time

file_path = r"C:\Users\lenovo\Desktop\MissionMerge\原神全 NPC 委托、轮次统计表.xlsx"

def read_args():
    parser = argparse.ArgumentParser(description='Get the path of a file')
    # 接受不定多个参数，存储在args.file中
    parser.add_argument('-o', nargs='+', help='文件路径')
    args = parser.parse_args()
    return args

def path_feature(file_path):
    os.stat(file_path).st_uid

if __name__ == '__main__':
    # args = read_args()
    # print(args.o)
    
    # for file_path in args.o:
    #     path_feature = path_feature(file_path)
    path_feature = path_feature(file_path)