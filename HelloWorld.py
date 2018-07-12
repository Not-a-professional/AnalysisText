#!/usr/bin/python
# encoding: UTF-8
import jieba.analyse
import jieba
import pytesseract
from PIL import Image
import fitz

jieba.suggest_freq('收益率', True)
jieba.suggest_freq('投资收益', True)
jieba.suggest_freq('固定利率', True)
jieba.suggest_freq('预期收益率', True)
jieba.suggest_freq('基准利率', True)
jieba.suggest_freq('调整', True)


def test_jieba(txt):
    # txt = u'欧阳建国是创新办主任也是欢聚时代公司云计算方面的专家'
    print(','.join(jieba.cut(txt)))


def test_pdf():
    read_pdf = 'C:\\Users\\user\\Documents\\WXWork\\1688851382425401\\Cache\\File\\2018-07\\test.PDF'
    doc = fitz.open(read_pdf)
    page_count = doc.pageCount
    print(page_count)

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
        print(text.replace(' ', ''))
    return


if __name__ == '__main__':
    test_jieba(u'欧阳建国是创新办主任也是欢聚时代公司云计算方面的专家')
