#!/usr/bin/python
# coding: UTF-8
import jieba.analyse
import jieba
import pytesseract
from PIL import Image
import fitz
from gensim.models import word2vec
import os

jieba.suggest_freq(u'收益率', True)
jieba.suggest_freq(u'投资收益', True)
jieba.suggest_freq(u'固定利率', True)
jieba.suggest_freq(u'预期收益率', True)
jieba.suggest_freq(u'基准利率', True)
jieba.suggest_freq(u'调整', True)


def test_jieba(txt):
    # txt = u'欧阳建国是创新办主任也是欢聚时代公司云计算方面的专家'
    return u' '.join(jieba.cut(txt))


def test_pdf():
    read_pdf = 'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\兖矿低热值燃料电厂项目受托合同_华泰招商环卫.pdf'
    doc = fitz.open(read_pdf)
    page_count = doc.pageCount
    print(page_count)

    target = 0
    upper = 0

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

        model = word2vec.Word2Vec(sentences, size=100, hs=1, min_count=1, window=3)

        current = model.most_similar([u'收益率'], topn=30)
        count = 0
        for key in current:
            count = count + 1
        if upper < count:
            upper = count
            target = i

    print(target)
    return


if __name__ == '__main__':
    test_pdf()
