#!/usr/bin/python
# -*- coding: utf-8 -*-
import jieba.analyse
import jieba
from PIL import Image
import fitz
from gensim.models import word2vec
import win32com
from win32com.client import Dispatch
import re
from aip import AipOcr
import os

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
    return ' '.join(jieba.cut(txt))


# def test_pdf():
#     read_pdf = 'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\哈尔滨不动产-工银安盛受托合同.pdf'
#     doc = fitz.open(read_pdf)
#     page_count = doc.pageCount
#     print(page_count)
#
#     target = ''
#
#     for i in range(page_count):
#         pytesseract.pytesseract.tesseract_cmd = 'D:\\Tesseract-OCR\\tesseract.exe'
#         page = doc[i]
#         zoom = int(300)
#         rotate = int(0)
#         trans = fitz.Matrix(zoom / 100.0, zoom / 100.0).preRotate(rotate)
#         pm = page.getPixmap(matrix=trans, alpha=False)
#         img = Image.frombuffer('RGB', [pm.width, pm.height], pm.samples,
#                                "raw", 'RGB', 0, 1)
#         text = pytesseract.image_to_string(img, lang='chi_sim')
#
#         target = target + text
#
#     target_list = re.split('[。]', test_jieba(target.replace(' ', '').replace('\n', '')))
#     sentences = []
#     for line in target_list:
#         sentences.append(line.split(' '))
#
#     model = word2vec.Word2Vec(sentences, min_count=1, window=100, size=100)
#
#     sim = model.wv.most_similar(u'投资收益率', topn=10)
#     for key in sim:
#         print(key[0], key[1])
#
#     return


def test_word():
    w = win32com.client.Dispatch('Word.Application')
    doc = w.Documents.Open(r'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\人保资本—南昌轨道交通债权投资计划受托合同-友邦江门.doc')
    for s in range(len(doc.Paragraphs)):
        sentences = doc.Paragraphs[s].Range.text
        if sentences.find(u'投资收益率') != -1 or sentences.find(u'预期收益率') != -1 or sentences.find(u'投资计划收益') != -1:
            print(sentences)
    return


def fast_test_pdf():
    with open('C:\\Users\\user\\Desktop\\新建文本文档.txt', 'r') as file_to_read:
        f_read = file_to_read.read()

    target_list = re.split(r'[。]', f_read)
    sentences = []
    for line in target_list:
        sentences.append(line.split(' '))

    model = word2vec.Word2Vec(sentences, min_count=5, size=100, iter=2, workers=4)
    sim = model.wv.most_similar(u'投资收益', topn=10)
    for key in sim:
        print(key[0], key[1])
    return


def baidu_orc():
    APP_ID = '11540842'
    API_KEY = 'aaDIvgPK6Ptb0itaS2mlUnOC'
    SECRET_KEY = 'iVoaZjKu69c11WIsQWKzj2bdiTy7i0Db'

    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    options = {}
    options["language_type"] = "CHN_ENG"
    options["detect_direction"] = "true"
    options["detect_language"] = "true"
    options["probability"] = 'true'

    read_pdf = 'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\2-受托合同（工银安盛）.pdf'
    doc = fitz.open(read_pdf)
    page_count = doc.pageCount

    i = 12
    page = doc[i]
    zoom = int(320)
    rotate = int(0)
    trans = fitz.Matrix(zoom / 100.0, zoom / 100.0).preRotate(rotate)
    pm = page.getPixmap(matrix=trans, alpha=False)
    img = Image.frombuffer('RGB', [pm.width, pm.height], pm.samples, "raw", 'RGB', 0, 1)

    img.save(str(i) + '.png')
    image = get_file_content(str(i) + '.png')
    res = client.basicGeneral(image, options)
    analysis_json(res)  #处理百度返回的json对象
    delete_file(str(i) + '.png')

    return


def get_file_content(file_path):
    with open(file_path, 'rb') as fp:
        return fp.read()


def delete_file(file_path):
    os.remove(file_path)
    return


def analysis_json(json_str):
    for i in range(json_str['words_result_num']):
        print(json_str['words_result'][i]['words'])
    return


if __name__ == '__main__':
    #j = '{"log_id": 2471272194,"words_result_num": 2,"words_result":[{"words": " TSINGTAO"},{"words": "青島睥酒"}]}'
    # analysis_json(json.loads(j))
    baidu_orc()
