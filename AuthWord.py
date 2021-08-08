from docx import Document
import os
from win32com import client as wc
import win32com
from win32com.client import Dispatch, constants

# 测试文件夹
filePath = 'file/'
# 获取当前文件夹绝对路径
filePwd = os.getcwd()

# 创建一个文件
def GenerateNewWord(filename):
    document = Document()
    document.save(filePath + filename)


# 复制文档
def copyWord(filePath, copeFilePath):
    document = Document(filePath + filePath)
    document.save(filePath + copeFilePath)

# 转换格式
def TransDocToDocx(docName,docxName):
    # 获取当前目录完整路径
    currentPath = os.getcwd()
    # 获取 旧doc格式word文件绝对路径名
    docName = os.path.join(currentPath + '\\\\file\\', docName)
    # 设置新docx格式文档文件名
    docxName = os.path.join(currentPath + '\\\\file\\', docxName)

    # 打开word应用程序
    word = wc.Dispatch('Word.Application')
    # 打开 旧word 文件
    doc = word.Documents.Open(docName)
    # 保存为 新word 文件,其中参数 12 表示的是docx文件
    doc.SaveAs(docxName, 12)
    # 关闭word文档
    doc.Close()
    word.Quit()


# 创建新的word文档
def funOpenNewFile():
    word = Dispatch('Word.Application')
    # 或者使用下面的方法，使用启动独立的进程：
    # word = DispatchEx('Word.Application')

    # 如果不声明以下属性，运行的时候会显示的打开word
    word.Visible = 0  # 0:后台运行 1:前台运行(可见)
    word.DisplayAlerts = 0  # 不显示，不警告
    # 创建新的word文档
    doc = word.Documents.Add()
    # 在文档开头添加内容
    myRange1 = doc.Range(0, 0)
    myRange1.InsertBefore('在文档开头添加内容，\nhttp://t.cn/A6ILo1NC\n')
    # 在文档末尾添加内容
    myRange2 = doc.Range()
    myRange2.InsertAfter('在文档末尾添加内容\n')

    # 在文档i指定位置添加内容
    i = 3
    myRange3 = doc.Range(2, i)
    myRange3.InsertAfter("在文档i指定位置添加内容\n")

    # doc.Save()  # 保存
    doc.SaveAs(os.getcwd() + "\\file\\生成的新文件.docx")  # 另存为
    doc.Close()  # 关闭 word 文档
    word.Quit()  # 关闭 office


# 打开已存在的word文件
def funOpenExistFile():
    word = Dispatch('Word.Application')
    # 或者使用下面的方法，使用启动独立的进程：
    # word = DispatchEx('Word.Application')
    word.Visible = 0  # 0:后台运行 1:前台运行(可见)
    word.DisplayAlerts = 0  # 不显示，不警告
    # 打开一个已有的word文档
    doc = word.Documents.Open(os.getcwd() + "\\生成的新文件.docx")
    # 在文档开头添加内容
    myRange1 = doc.Range(0, 0)
    myRange1.InsertBefore('Hello word\n')
    # 在文档末尾添加内容
    myRange2 = doc.Range()
    myRange2.InsertAfter('Bye word\n，更多：http://t.cn/A6ILogFC')
    # 在文档i指定位置添加内容
    i = 0
    myRange3 = doc.Range(0, i)
    myRange3.InsertAfter("what's up, bro?\n")
    # doc.Save()  # 保存
    doc.SaveAs(os.getcwd() + "\\funOpenExistFile.docx")  # 另存为
    doc.Close()  # 关闭 word 文档
    word.Quit()  # 关闭 office


# 生成Pdf文件
def funGeneratePDF():
    word = Dispatch("Word.Application")
    word.Visible = 0  # 后台运行，不显示
    word.DisplayAlerts = 0  # 不警告
    # 打开一个已有的word文档
    doc = word.Documents.Open(os.getcwd() + "\\file\\格式待转换文件.doc")
    # txt=4, html=10, docx=16， pdf=17
    doc.SaveAs(os.getcwd() + "\\转换word为pdf.pdf", 17)
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    # 创建一个文件
    # GenerateNewWord('测试文档.docx')
    # 复制文档
    # copyWord('测试文档.docx', 'test.docx')
    # 文件格式转换
    # TransDocToDocx("格式待转换文件.doc", "格式转换完成文件.docx")
    # 创建具有内容的文档
    # funOpenNewFile()
    funGeneratePDF()