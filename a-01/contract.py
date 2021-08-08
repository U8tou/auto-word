# @File : contract.py
# @Description: 按模板生成文件
# @Author : mengbao
# @Time : 2021/8/8 17:51
from mailmerge import MailMerge
from datetime import datetime
import os

# 获取当前文件夹绝对路径
filePwd = os.getcwd()


# 生成合同
def GenerateCertify(tempPath, wordPath):
    # 打开模板
    document = MailMerge(tempPath)
    # 替换内容
    # document.merge(obj)
    document.merge(name='唐星', code='1010101010', year='2020', salary='13k', job='嵌入式软件开发工程师')
    # 保存文件
    document.write(wordPath)


if __name__ == "__main__":
    # 输入模板路径
    tempPath = filePwd + '\\mode\\薪资证明模板.docx'
    # 输入生成目录
    wordPath = filePwd + '\\static\\'
    # 输入文件名
    wordName = '薪资证明'
    # 输入生成份数
    count = 10
    # 模板填充对象
    obj = ''
    # 获得开始时间
    startTime = datetime.now()
    # 开始生成
    for i in range(count):
        fileName = f'{wordPath}{wordName}{i+1}.docx'
        GenerateCertify(tempPath, fileName)
    # 获取结束时间
    endTime = datetime.now()
    # 计算时间差
    allSeconds = (endTime - startTime).seconds
    print(f"生成{count}份合同一共用时: {str(allSeconds)} 秒")
    print("程序结束！")
