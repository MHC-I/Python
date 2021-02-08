import os 
from win32com import client as wc

def doc_to_docx(doc_path):
    word = wc.Dispatch("Word.Application") # 打开word应用程序
    word.Visible = 0   # 后台运行
    word.DisplayAlerts = 0    # 不警告
    doc = word.Documents.Open(doc_path)#打开word文件
    doc.SaveAs(doc_path+'x', 12)#另存为后缀为".docx"的文件，其中参数12指docx文件
    doc.Close() #关闭原来word文件
    word.Quit() #退出word
    os.remove(doc_path)#删除原文件
    print(doc_path+" 完成！")#显示结果

#调用walk方法遍历
path='E:\\确认表\\医检1班'
for root,dirs,files in os.walk(path):
    files = [f for f in files if not f[0] == '.']#忽略隐藏文件
    dirs[:] = [d for d in dirs if not d[0] == '.']
    for name in files:
        if name[-3:]=='doc':#判断文件类型是否为doc
            doc_path=os.path.join(root,name)
            doc_to_docx(doc_path)