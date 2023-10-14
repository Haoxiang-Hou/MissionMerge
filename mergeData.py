import pandas as pd
import numpy as np

old_file_path = './old.xlsx'
update_file_path = './update.xlsx'
output_data_path = './output_data.xlsx'

def read_excel(path, read_list=None):
    # 读取excel文件，返回一个字典，key为sheet_name，value为sheet的数据
    # 如果sheet_name为None，则读取所有sheet
    sheet_names = pd.ExcelFile(path).sheet_names
    if read_list is None:
        read_list = sheet_names
    else:
        read_list = [sheet_name for sheet_name in read_list if sheet_name in sheet_names]
        
    output_file = {}
    for sheet_name in read_list:
        output_file[sheet_name] = read_sheet(path, sheet_name)
    return output_file
    
    
def read_sheet(path, sheet_name):
    # 读取excel文件的A:J列，跳过前1行
    df = pd.read_excel(path, usecols='A:J', skiprows=1, sheet_name=sheet_name)
    # 将df中第0行的数据读取，如果不为NaN，则作为列名，否则保留原列名
    df.columns = df.iloc[0, :].where(df.iloc[0, :].notnull(), df.columns)
    # 删除第0行
    df.drop([0], inplace=True)
    # 重置索引
    df.reset_index(drop=True, inplace=True)
    return df

def merge_data(old_data, update_data):
    # 合并两个字典，如果有相同的key，则merge_sheet，如果没有重复，则两个字典的key都保留
    output_data = {}
    for key in old_data.keys():
        if key in update_data.keys():
            output_data[key] = merge_sheet(old_data[key], update_data[key])
        else:
            output_data[key] = old_data[key]
    for key in update_data.keys():
        if key not in old_data.keys():
            output_data[key] = update_data[key]
    return output_data

def merge_sheet(old_sheet, update_sheet):
    # 将两个sheet合并
    # 根据第一列的值Quest ID，将两个sheet合并
    # 如果ID不同，则将update_sheet的数据添加到old_sheet的最后一行
    # 如果ID相同，则将update_sheet的3,4,5列的数据对应覆盖到old_sheet的3,4,5列
    output_sheet = old_sheet.copy()
    for i in range(update_sheet.shape[0]):
        if update_sheet.iloc[i, 0] not in old_sheet['Quest ID'].values:
            output_sheet = output_sheet.append(update_sheet.iloc[i, :])
        else:
            # 找到update_sheet中第i行的Quest ID在old_sheet中的行号
            index = np.where(old_sheet['Quest ID'].values == update_sheet.iloc[i, 0])[0][0]
            output_sheet.iloc[index, 2] = update_sheet.iloc[i, 2]
            output_sheet.iloc[index, 3] = update_sheet.iloc[i, 3]
            output_sheet.iloc[index, 4] = update_sheet.iloc[i, 4]
    # output_sheet.reset_index(drop=True, inplace=True)
    return output_sheet

def save_excel(output_data, output_file_path):
    # 将output_data中的数据保存到output_file_path中，其余位置保持为old_file_path中的数据，并且保持原有的sheet顺序，数据格式、颜色等不变
    # 如果output_file_path中没有某个sheet，则添加该sheet
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    for key in output_data.keys():
        output_data[key].to_excel(writer, sheet_name=key, index=False)
    writer.save()
    writer.close()
    


if __name__ == '__main__':
    old_data = read_excel(old_file_path)
    update_data = read_excel(update_file_path)
    output_data = merge_data(old_data, update_data)
    save_excel(output_data, output_data_path)
    
    