import os
from docx import Document
'''该脚本将srt字幕中的句子转化为word文档'''

#输入filepath后修改srt字幕文件为txt文件
def newsrt(filepath):
    #修改文件名,使得文件可读
    os.rename(filepath,filepath+'.txt')
    filepath=filepath+'.txt'
    #在内存中创建word文档
    document = Document()
    #处理srt字幕内容，使其可读
    a = 1
    b = 2
    c = 3
    state = a
    text = ''
    with open(filepath, 'r', encoding='utf-8-sig') as f: #打开srt字幕文件，并去掉文件开头的\ufeff
        for line in f.readlines(): #遍历srt字幕文件
            if state == a: #跳过第一行
                state = b
            elif state == b: #跳过第二行
                state = c
            elif state == c: #读取第三行字幕文本
                if len(line.strip()) !=0:
                    text += ' ' + line.strip() #将同一时间段的字幕文本拼接
                    state = c
                elif len(line.strip()) ==0:
                    document.add_paragraph(text)
                    #将内容添加到word中
                    text = '\n'
                    state = a
    #保存word文档到原文件位置                
    len_filepath=len(filepath)
    docx_path=filepath[0:len_filepath-4]+'已修改.docx'
    document.save(docx_path)
    #删除原文件
    os.remove(filepath)
    #显示处理过程
    print(docx_path+'已完成')

#调用walk方法遍历
path='/Volumes/CoCo/1'
for root,dirs,files in os.walk(path):
    files = [f for f in files if not f[0] == '.']
    dirs[:] = [d for d in dirs if not d[0] == '.']
    for name in files:
        filepath=os.path.join(root,name)
        newsrt(filepath)