import xlwings as xw
import mergeData
import time
import argparse

old_file_path = './old.xlsx'
update_file_path = './update.xlsx'
output_file_path = './new.xlsx'

region_list = ['蒙德', '璃月', '稻妻', '须弥', '枫丹']

def read_args():
    # 接受参数
    # -d，开启debug模式，默认关闭，加入后设置debug_state为True
    # -o，指定old_file的路径，默认为./old.xlsx
    # -u，指定update_file的路径，默认为./update.xlsx
    # -n，指定output_file的路径，默认为./new.xlsx
    parser = argparse.ArgumentParser(description='更新数据')
    parser.add_argument('-d', '--debug', action='store_true', help='开启debug模式')
    parser.add_argument('-o', '--old', default='./old.xlsx', help='指定old_file的路径，默认为./old.xlsx')
    parser.add_argument('-u', '--update', default='./update.xlsx', help='指定update_file的路径，默认为./update.xlsx')
    parser.add_argument('-n', '--output', default='./new.xlsx', help='指定output_file的路径，默认为./new.xlsx')
    args = parser.parse_args()
    
    global debug_state
    global old_file_path
    global update_file_path
    global output_file_path
    debug_state = args.debug
    old_file_path = args.old
    update_file_path = args.update
    output_file_path = args.output
    
    return args

if __name__ == '__main__':
    args = read_args()
    start_time = time.time()
    last_time = start_time
    current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print(current_time, '开始更新数据...')
    
    with xw.App(visible=debug_state, add_book=False) as app:
        # 复制update_file到output_file，保持原有的sheet顺序，数据格式、颜色等不变
        output_file = app.books.open(update_file_path)
        # 用户数据存储在old_data中
        old_data = mergeData.read_excel(old_file_path, region_list)
        if debug_state:
            print("数据读取完成！用时：", round(time.time() - last_time, 2), '秒')
            last_time = time.time()
        # 修改output_file中的数据，按照old_data中的数据进行修改
        for sheet_name, old_sheet_data in old_data.items():
            if sheet_name not in region_list:
                continue
            sheet = output_file.sheets[sheet_name]
            sheet_current_end_row = sheet.range('A1').current_region.rows.count
            sheet_ID_list = sheet.range((1, 1), (sheet_current_end_row, 1)).value
            sheet_ID_set = set(sheet_ID_list)
            old_ID_list = old_sheet_data['Quest ID'].values
            old_ID_dict = dict(zip(old_ID_list, range(len(old_ID_list))))
            
            if debug_state:
                print('\t', sheet_name, '开始更新...已累计用时：', round(time.time() - last_time, 2), '秒')
            
            # 将update_file中出现的数据从output_file中查找出来，覆盖到output_file中
            for i, Quest_ID in enumerate(sheet_ID_list):
                # 查找i行1列的Quest ID是否在old_data中
                if Quest_ID not in old_ID_dict or Quest_ID is None:
                    continue
                # 查找Quest ID在old_data中的行号
                index = old_ID_dict[Quest_ID]
                # 将old_data中(index, 5)到(index, 9)的数据覆盖到output_file中(i+1, 6)到(i+1, 10)的位置
                sheet.range((i+1, 6)).value = old_sheet_data.iloc[index, 5:10].values
                
            if debug_state:
                print('\t', sheet_name, '已更新', len(sheet_ID_list), '条数据，已累计用时：', round(time.time() - last_time, 2), '秒')
                    
            # 将update_file中没有出现，但在old_data中出现的数据添加到output_file末尾
            for i, Quest_ID in enumerate(old_ID_list):
                if Quest_ID not in sheet_ID_set or Quest_ID is None:
                    sheet.range((sheet_current_end_row + 1, 1)).value = old_sheet_data.iloc[i, :10].values
                    sheet_current_end_row += 1
            
            if debug_state:
                print('\t', sheet_name, '更新完成！用时：', round(time.time() - last_time, 2), '秒\n')
                last_time = time.time()
                    
        output_file.save(output_file_path)
        output_file.close()
        
        if debug_state:
            print('保存完成！用时：', round(time.time() - last_time, 2), '秒')
            last_time = time.time()
        
    # 显示更新用时，保留两位小数
    print('更新完成！累计用时：', round(time.time() - start_time, 2), '秒')
    
    
    
    
    