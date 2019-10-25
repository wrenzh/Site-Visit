# Insert QR code to PDF
# Author: Wen Zhang (wzhang53@utk.edu)
# Version: 0.3.0
# Date: 10/25/2019
# License: MIT

import os
import sys
import io
import fitz
import qrcode


def insertQRCode(srcPDF, link, location):
    """Insert the QR code to link on pdfFile"""
    buf = io.BytesIO()
    qrcode.make(link).save(buf)

    doc = fitz.open(srcPDF)
    doc[0].insertImage(fitz.Rect(*location), stream=buf.getvalue())

    path = os.path.dirname(srcPDF)
    tmp = os.path.join(path, "tmp.pdf")
    doc.save(tmp)
    doc.close()
    os.replace(tmp, srcPDF)


def batchQRCode(srcDir, baseLink, location, verbose):
    """Insert QR code to every pdf file in srcDir"""
    allFiles = []
    for root, _, files in os.walk(srcDir):
        for f in files:
            allFiles.append(os.path.join(root, f))

    if baseLink[-1] != '/':
        baseLink += '/'

    for f in allFiles:
        if verbose:
            print(f"Inserting QR code to {f}")

        base = os.path.basename(f)
        link = baseLink + base
        insertQRCode(f, link, location)

        if verbose:
            print("Success, replacing original file")