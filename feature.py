import xlrd
import jieba.analyse
import jieba

#读取数据
def read(file, sheet_index=0):
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(sheet_index)
    #cell = sheet.cell(1,6).value
    #print(cell)
    #print("---------")
    #return cell
    for i in range(1,1115):
        cell_value = sheet.cell(i,6).value
        out = seg_sentence(cell_value)
        print(keywords(out))

#去除停用词

def stopwordslist(filepath):
    stopwords = [line.strip() for line in open(filepath,'r').readlines()]
    return stopwords

#句子进行分词
def seg_sentence(sentence):
    sentence_seged = jieba.cut(sentence.strip())
    stopwords = stopwordslist('stopword.txt')  # 这里加载停用词的路径
    outstr = ''
    for word in sentence_seged:
        if word not in stopwords:
            if word != '\t':
                outstr += word
                outstr += " "
    return outstr

#提取关键词
def keywords(outstr):
    keys = jieba.analyse.extract_tags(outstr,topK=15,withWeight=False)
    return keys


if __name__ == '__main__':
    read('java.xlsx')
