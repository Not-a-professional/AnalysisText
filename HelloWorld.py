#!/usr/bin/python
# coding: UTF-8
import jieba.analyse
import jieba
import pytesseract
from PIL import Image
import fitz
from gensim.models import word2vec
import win32com
from win32com.client import Dispatch

jieba.suggest_freq(u'投资收益率', True)
jieba.suggest_freq(u'投资计划收益', True)
jieba.suggest_freq(u'固定利率', True)
jieba.suggest_freq(u'预期收益率', True)
jieba.suggest_freq(u'基准利率', True)
jieba.suggest_freq(u'调整', True)
jieba.suggest_freq(u'浮动利率', True)
jieba.suggest_freq(u'上限', True)
jieba.suggest_freq(u'下限', True)
jieba.suggest_freq(u'最高', True)
jieba.suggest_freq(u'最低', True)


def test_jieba(txt):
    # txt = u'欧阳建国是创新办主任也是欢聚时代公司云计算方面的专家'
    return u' '.join(jieba.cut(txt))


def test_pdf():
    read_pdf = 'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\4-昆明轨交-受托合同-浦发.pdf'
    doc = fitz.open(read_pdf)
    page_count = doc.pageCount
    print(page_count)

    target = ''

    for i in range(page_count):
        pytesseract.pytesseract.tesseract_cmd = 'D:\\Tesseract-OCR\\tesseract.exe'
        page = doc[i]
        zoom = int(300)
        rotate = int(0)
        trans = fitz.Matrix(zoom / 100.0, zoom / 100.0).preRotate(rotate)
        pm = page.getPixmap(matrix=trans, alpha=False)
        img = Image.frombuffer('RGB', [pm.width, pm.height], pm.samples,
                               "raw", 'RGB', 0, 1)
        text = pytesseract.image_to_string(img, lang='chi_sim')

        sentences = test_jieba(text.replace(' ', '').replace('\n', ''))

        target += sentences

    model = word2vec.Word2Vec(size=100, iter=1, workers=4)
    model.build_vocab()
    model.train(target)
    sim = model.wv.most_similar(u'收益', topn=100)
    for key in sim:
        print(key[0], key[1])

    return


def test_word():
    w = win32com.client.Dispatch('Word.Application')
    doc = w.Documents.Open(r'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\人保资本—南昌轨道交通债权投资计划受托合同-友邦江门.doc')
    for s in range(len(doc.Paragraphs)):
        sentences = doc.Paragraphs[s].Range.text
        if sentences.find(u'投资收益率') != -1 or sentences.find(u'预期收益率') != -1 or sentences.find(u'投资计划收益') != -1:
            print(sentences)
    return


def fast_test_pdf():
    f_read = open('C:\\Users\\user\\Desktop\\新建文本文档.txt', 'r')
    model = word2vec.Word2Vec(f_read, min_count=5, size=100, iter=5)
    # model.wv.vocab = {u'收益', 10}
    # model.wv.vectors = [[]]
    # model.train(f_read, total_examples=1, epochs=2, word_count=len(f_read))
    sim = model.wv.most_similar(u'收', topn=100)
    for key in sim:
        print(key[0], key[1])
    return


if __name__ == '__main__':
    fast_test_pdf()
