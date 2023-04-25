import numpy as np
from re import findall
from zipfile import ZipFile
import os
import shutil
from win32com import client as wc

map={'1':'引言适合读者，有吸引力和创新型',
              '2':'介绍主题和读者的关系',
              '3':'清晰提出作者的主要观点',
              '4':'对主题的关键词进行清晰的定义或解释',
              '5':'提出3个及以上的分论点来论证主旨',
              '6':'分论点能从不同角度论证主旨',
              '7':'论据与论点相关且充分，可在逻辑上充分支持论点',
              '8':'论据有不同类型，具有代表性或是解释意义',
              '9':'论据提供背景信息（来源，时间，出处及主要内容）',
              '10':'考虑到反例，反方和替代观点',
              '11':'指出反方观点中的逻辑错误',
              '12':'陈述对立/不同观点的本质',
              '13':'过渡语提示文章即将结束',
              '14':'概述问题的实质与作者独特的观点',
              '15':'能统一正反两方的观点得出最终结论',
              '16':'概念，关键词清楚，具体，一致',
              '17':'段落和篇章结构基本遵循“总-分-总”（即观点-论据-论证+结论）',
              '18':'句子与段落之间有逻辑连接的词或短语，过度合理流畅',
              '19':'论证结构为正-反-正',
              '20':'能明确指出分析性阅读中的错误假设',
              'A':'文本写作错误：流水句，在独立主句间使用逗号',
              'B':'文本写作错误：谓语动词错误',
              'C':'文本写作错误：指示代词指代混乱，指称不一致',
              'D':'文本写作错误：表意不清'}
N_subject=10
N_item=20
N_attribute=10
item_list=['1','2','3','4','5','6','7','8','9',
           '10','11','12','13','14','15','16',
           '17','18','19','20']
Q_matrix=[[0,0,0,0,0,0,0,1,0,1],
          [0,0,0,0,0,0,0,1,0,0],
          [0,0,0,1,1,0,0,0,0,0],
          [0,0,0,1,0,0,1,0,0,0],
          [0,0,0,1,0,0,0,0,0,0],
          [0,1,0,0,0,1,0,1,0,0],
          [0,1,0,0,0,0,1,0,0,0],
          [1,0,0,0,0,1,1,0,0,0],
          [0,0,0,0,1,1,0,1,0,0],
          [0,0,1,0,0,0,0,1,1,0],
          [1,1,1,0,0,0,0,0,1,0],
          [0,1,1,1,1,0,0,1,1,1],
          [0,0,0,1,0,0,0,1,0,0],
          [0,0,1,1,0,1,0,0,0,1],
          [0,0,0,0,0,0,0,0,1,1],
          [0,0,0,1,1,1,0,0,0,0],
          [1,1,0,0,0,1,1,0,0,1],
          [1,0,0,1,0,0,0,1,0,0],
          [1,1,1,0,0,0,0,0,0,0],
          [0,1,1,0,0,0,1,0,0,0]]
def doc2docx(path):
    # path=r'C:\Users\Vernon\Desktop\critical_writing\annotation_test\half-and-half(assessed)'
    word=wc.Dispatch('Word.Application')
    for file_name in os.listdir(path):
        file_path=os.path.join(path,file_name)
        doc=word.Documents.Open(file_path)
        doc.SaveAs('{}x'.format(file_path),12)
        doc.Close()
    word.Quit()
def deldoc(path):
    # path=r'C:\Users\Vernon\Desktop\critical_writing\annotation_test\half-and-half(assessed)'
    for file_name in os.listdir(path):
        if file_name.endswith('.doc'):
            file_path=os.path.join(path,file_name)
            os.remove(file_path)

def attribute_find(fn):
    comments=[]
    attributes=[]
    with ZipFile(fn) as fp:
        try:
            content = fp.read("word/comments.xml").decode("utf-8")
        except:
            content=''
        if not content:
            print(fn)
            print('该同学没有comments，略过')
            return -1
    for i in range(10):
        item1='<w:t>1</w:t></w:r><w:r><w:t>'+str(i)+'</w:t>'
        k='1'+str(i)
        item2='<w:t>'+k+'</w:t>'
        if content.find(item1)!=-1:
            comments.append(int(k))
            content=content.replace(item1,item2)
    item1='<w:t>2</w:t></w:r><w:r><w:t>0</w:t>'
    k='20'
    item2='<w:t>'+k+'</w:t>'
    if content.find(item1)!=-1:
        comments.append(int(k))
        content=content.replace(item1,item2)
    for i in range(1,10):
        item1='<w:t>'+str(i)+'</w:t>'
        if content.find(item1)!=-1:
            comments.append(int(i))

    comments=list(set(comments))
    for comment in comments:
        for i in range(N_attribute):
            if Q_matrix[comment-1][i]==1 & i not in attributes:
                attributes.append(i)
    return attributes

def main(path):
    # doc2docx(path)
    # deldoc(path)
    sub_attr_matrix=np.zeros((N_subject,N_attribute))
    index=0
    for file_name in os.listdir(path):
        file_path = os.path.join(path, file_name)
        attributes=attribute_find(file_path)
        if attributes==-1:
            continue
        if attributes is not list:
            sub_attr_matrix[index,attributes]=1
        else:
            for attribute in attributes:
                sub_attr_matrix[index,attribute]=1
        index+=1
    return sub_attr_matrix

matrix=main(r'C:\Users\Vernon\Desktop\critical_writing\half-and-half(assessed)')
print(matrix)