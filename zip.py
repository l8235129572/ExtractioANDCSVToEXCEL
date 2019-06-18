import os
import zipfile
import pandas as pd

# csv添加指定列,并将csv转为.xls
def pd_csv_to_xlsx(path, csv_nama):
    new_title = []
    f = pd.read_csv(path + csv_nama, header=0, low_memory=False, encoding='utf-8')
    titles = f.columns  # 获取表头
    for title in titles:
        new_title.append(title)
    # 在csv第二列添加  店铺
    new_title.insert(1, '店铺')
    # 指定店铺名
    f['店铺'] = 'olay官方旗舰店'
    f.to_csv(path + '\\new_' + csv_nama, columns=new_title, index=0, encoding="utf_8_sig")
    # 将csv转为excel
    csv = pd.read_csv(path + '\\new_' + csv_nama, encoding='utf-8')
    xls_name = csv_nama.replace('csv', 'xls')
    csv.to_excel(path + xls_name, sheet_name='data')

# 指定文件夹下的csv压缩文件,批量解压到指定文件
def jieya(path, path2):
    for name in os.listdir(path):  # 获取当前目录下的所有文件名
        file_path = path + name  # 拼接新路径,进行解压
        with zipfile.ZipFile(file_path, 'r') as f:
            for fn in f.namelist():
                filename = fn.encode('cp437').decode('gbk')
                f.extract(fn, path2)
                os.chdir(path2)  # 切换到目标目录
                os.rename(fn, filename)
        pd_csv_to_xlsx(path2, filename)



if __name__ == '__main__':
    jieya('C:\\Users\EDZ\\Desktop\\宝贝\\', 'C:\\Users\EDZ\\Desktop\\file\\')
