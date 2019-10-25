# Find the poster title in pdf
# Author: Wen Zhang (wzhang53@utk.edu)
# Version: 0.3.0
# Date: 10/25/2019
# License: MIT

import fitz
from operator import itemgetter


def getPosterTitle(pdfFilename, rect=[290, 0, 1728, 290]):
    doc = fitz.open(pdfFilename)
    words = doc[0].getTextWords()
    title = [w for w in words if fitz.Rect(w[:4]) in fitz.Rect(rect)]
    title.sort(key=itemgetter(3, 0))
    return ' '.join(w[4] for w in title)