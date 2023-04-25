import os
from zipfile import ZipFile
import shutil
from win32com import client as wc

def doc2docx(path):
    # path=r'C:\Users\Vernon\Desktop\critical_writing\annotation_test\half-and-half(assessed)'
    word=wc.Dispatch('Word.Application')
    for file_name in os.listdir(path):
        if file_name.endswith('.doc'):
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
# <w:t>1</w:t></w:r><w:r><w:t>3</w:t>
def replace_comments(ori_comments):
    for i in range(10):
        item1='<w:t>1</w:t></w:r><w:r><w:t>'+str(i)+'</w:t>'
        k='1'+str(i)
        item2='<w:t>'+map[k]+'</w:t>'
        ori_comments=ori_comments.replace(item1,item2)
    item1='<w:t>2</w:t></w:r><w:r><w:t>0</w:t>'
    k='20'
    item2='<w:t>'+map[k]+'</w:t>'
    ori_comments = ori_comments.replace(item1, item2)
    for i in ['1','2','3','4','5','6','7','8',
              '9','A','B','C','D']:
        item1='<w:t>'+i+'</w:t>'
        item2='<w:t>'+map[i]+'</w:t>'
        ori_comments=ori_comments.replace(item1,item2)
    return ori_comments
def change_comments(fn):
    with ZipFile(fn) as fp:
        try:
            content = fp.read("word/comments.xml").decode("utf-8")
        except:
            content=''
        if not content:
            return 0
        rep_comments = replace_comments(content)

    delet_file = "word/comments.xml"
    zin = ZipFile(fn, 'r')  # 读取对象
    zout = ZipFile('1.docx', 'w')  # 被写入对象
    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if (item.filename != delet_file):  # 剔除要删除的文件
            zout.writestr(item, buffer)  # 把文件写入到新对象中
    zout.close()
    zin.close()
    shutil.move('1.docx', fn)

    with ZipFile(fn,'a') as fp:
        with fp.open("word/comments.xml",'w') as fp_com:
            fp_com.write(bytes(rep_comments,encoding='utf8'))
    return 1

def main(path):
    doc2docx(path)
    deldoc(path)
    for file_name in os.listdir(path):
        file_path = os.path.join(path, file_name)
        change_comments(file_path)

path=r'C:\Users\Vernon\Desktop\critical_writing\31班_essaywriting_1\WICKED BOYS'
main(path)