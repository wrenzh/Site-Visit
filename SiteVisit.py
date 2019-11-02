# Process the submission files
# Author: Wen Zhang (wzhang53@utk.edu)
# Version: 0.3.0
# Date: 10/25/2019
# License: MIT

import os
import sys
import subprocess
import shutil
import re
import io
import datetime
import comtypes.client
import fitz
import qrcode
from operator import itemgetter


def decode(filename):
    """Decode the filename with the expected filename convention"""
    # Replace multiple underscores and space with a single one and split
    fields = re.sub(r"\s*\_+\s*", "_", filename).split('_')

    if len(fields) not in range(4, 7):
        raise ValueError(f"{filename} uses wrong convention")

    last, first, univ, prof = fields[0], fields[1], fields[2], fields[3]
    indx = "1" if len(fields) < 5 else fields[4].strip()
    revs = "R0" if len(fields) < 6 else fields[5].strip()

    if (not indx.isdigit()) or (not revs[1:].isdigit()):
        raise ValueError(f"{filename} uses wrong convention")

    return last, first, univ, prof, indx, revs


def categorize(filepath):
    """Find the poster/paper category and area"""
    c = re.compile("Core|Non-core|Associated", flags=re.IGNORECASE)
    a = re.compile(r"Hardware\sTestbed|Actuation\sand\sHVDC|"
                   r"Large\sScale\sTestbed|Other\sCategories|"
                   r"Power\sConverter\sDesign\sand\sControl|"
                   r"Power\sElectronics\sDevices\sand\sComponents|"
                   r"Power\sSystem\sControl|Power\sSystem\sEstimation|"
                   r"Power\sSystem\sModeling|Power\sSystem\sMonitoring",
                   flags=re.IGNORECASE)

    if c.search(filepath):
        category = c.search(filepath).group(0)
    else:
        category = None

    if a.search(filepath):
        area = a.search(filepath).group(0)
    else:
        raise ValueError(f"Filepath {filepath} does not contain research area")

    return category, area


def getAllFiles(dir):
    """Get a list of all files in dir including sub-dir"""
    allFiles = []
    for root, _, files in os.walk(dir):
        for f in files:
            allFiles.append(os.path.join(root, f))
    
    return allFiles

def batch2pdf(srcDir, verbose):
    """Autocatically replace ppt/doc with pdf in srcDir"""
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    word = comtypes.client.CreateObject("Word.Application")

    allFiles = getAllFiles(srcDir)

    for f in allFiles:
        path = os.path.dirname(f)
        name, ext = os.path.splitext(os.path.basename(f))

        if ext.lower() == ".pdf":
            pass
        elif ext.lower() == ".csv":
            # Likely the list of papers and posters so don't do anything
            pass
        elif ext.lower() in [".ppt", ".pptx", ".doc", ".docx"]:
            if verbose:
                print(f"Replacing {name} with pdf...", end=" ")

            try:
                if ext.lower() in [".ppt", ".pptx"]:
                    deck = powerpoint.Presentations.Open(f)
                    # Powerpoint formatType code 32 for pdf
                    deck.SaveAs(os.path.join(path, name + ".pdf"), 32)
                    deck.Close()
                else:
                    doc = word.Documents.open(os.path.join(srcDir, f))
                    # Word formatType code 17 for pdf
                    doc.SaveAs(os.path.join(path, name + ".pdf"), 17)
                    doc.Close()
            except comtypes.COMError:
                print(f"\nERROR: Converting {f} failed")
                print("Please check if the file has Visio drawings")
                print("Also check PDF export settings (PDF/A must be off)")
                print("Also maximize the Powerpoint/Word application")
                quit()

            if verbose:
                print(f"Success")

            os.remove(f)
        else:
            raise ValueError(f"Unexpected filetype .{ext} in {path}")


def scan(srcDir):
    """Process papers and posters from root folder"""
    allFiles = getAllFiles(srcDir)

    # Determine if a file is paper or poster and record in dictionary
    posters = []
    papers = []
    for f in allFiles:
        base = os.path.basename(f)
        path = os.path.dirname(f)
        name, ext = os.path.splitext(base)

        isPoster = "Poster" in path or ext.lower() in [".ppt", ".pptx"]
        isPaper = "Paper" in path or ext.lower() in [".doc", ".docx"]

        if (not (isPoster or isPaper)) and ext.lower() == ".pdf":
            # Determine by other files in the same folder
            for a in os.listdir(path):
                isPaper |= ".doc" in a.lower()
                isPoster |= ".ppt" in a.lower()

        if (isPoster and isPaper) and (not (isPoster or isPaper)):
            raise ValueError(f"Ambiguous type {f}: paper or poster?")

        category, area = categorize(path)
        last, first, univ, prof, indx, revs = decode(name)
        # Formatted name
        fname = '_'.join([last, first, univ, prof, indx, revs])
        # Base name without revision number
        bname = '_'.join([last, first, univ, prof, indx])
        revs = int(revs[1:])
        # Modified time
        mtime = os.path.getmtime(f)
        d = {"file": f, "fname": fname, "bname": bname, "revs": revs,
             "last": last, "first": first, "univ": univ, "prof": prof,
             "indx": indx, "ext": ext, "category": category, "area": area, 
             "mtime": mtime}

        if isPoster:
            posters.append(d)
        else:
            papers.append(d)

    return papers, posters


def removeOldRevisions(fileList):
    """Find and delete older revisions"""
    baseNames = [f["bname"] for f in fileList]
    revisions = [f["revs"] for f in fileList]

    # Sorted indices with baseNames to aggeregate same paper/poster
    sIndx = sorted(range(len(baseNames)), key=baseNames.__getitem__)
    sNames = [baseNames[i] for i in sIndx]
    sRevs = [revisions[i] for i in sIndx]

    prevName = ''
    lastRev = -1
    indxToDel = []
    for i, n in enumerate(sNames):
        if n != prevName:
            prevName = n
            lastRev = sRevs[i]
        elif sRevs[i] > lastRev:
            print(f"Remove ver. {lastRev} < latest {sRevs[i]} for {n}")

            indxToDel.append(sIndx[i-1])
            lastRev = sRevs[i]
        elif sRevs[i] < lastRev:
            print(f"Remove ver. {sRevs[i]} < latest {lastRev} for {n}")

            indxToDel.append(sIndx[i])

    return [fileList[i] for i in set(sIndx) - set(indxToDel)]


def checkDuplicate(fileList):
    """Find duplicate submissions in fileList"""
    duplicate = False

    uniqs = []
    for f in fileList:
        if f["fname"] not in uniqs:
            uniqs.append(f["fname"])
        else:
            duplicate = True
            print(f"Duplicate found: {f['file']}")

    if duplicate:
        print("Duplicate submission file found. Stopping...")
        quit()


def copyFormated(papers, posters, dstDir, verbose):
    """Copy files in papers and posters to dstDir as structured"""
    for p in papers:
        # No categories for papers
        dstSubdir = os.path.join(dstDir, "Papers", p["area"])
        os.makedirs(dstSubdir, exist_ok=True)
        dstFile = os.path.join(dstSubdir, p["bname"] + p["ext"])
        if os.path.isfile(dstFile):
            raise ValueError(f"File already exists: {dstFile}")

        shutil.copy2(p["file"], dstFile)

        if verbose:
            print(f"Paper {p['fname']} copied to {dstSubdir}")

    for p in posters:
        dstSubdir = os.path.join(dstDir, "Posters", p["category"], p["area"])
        os.makedirs(dstSubdir, exist_ok=True)
        dstFile = os.path.join(dstSubdir, p["bname"] + p["ext"])
        if os.path.isfile(dstFile):
            raise ValueError(f"File already exists: {dstFile}")

        shutil.copy2(p["file"], dstFile)

        if verbose:
            print(f"Poster {p['fname']} copied to {dstSubdir}")


def getPosterTitle(pdfFile, rect=[0, 0, 1728, 290]):
    """Extract poster title from pdf"""
    # [0, 0, 1728, 290] works for the poster template
    doc = fitz.open(pdfFile)
    words = doc[0].getTextWords()
    title = [w for w in words if fitz.Rect(w[:4]) in fitz.Rect(rect)]
    title.sort(key=itemgetter(3, 0))
    return ' '.join(w[4] for w in title).strip().encode("utf-8", 'ignore')


def saveList(papers, posters, dstDir):
    """Write list of papers and posters to csv files in dstDir"""
    paperFile = os.path.join(dstDir, "Papers.csv")
    with open(paperFile, "w", encoding="utf-8") as f:
        f.write("File name, Area, Last name, First name, University, "
                "Professor, Index, Revision, Last Modified, Title, Abstract\n")

        for p in papers:
            # Title and abstract needs mannual input
            title = " "
            abstract = " "
            f.write(f"{p['fname']}, "
                    f"{p['area']}, "
                    f"{p['last']}, "
                    f"{p['first']}, "
                    f"{p['univ']}, "
                    f"{p['prof']}, "
                    f"{p['indx']}, "
                    f"{p['revs']}, "
                    f"{p['mtime']}, "
                    f"{title}, "
                    f"{abstract}\n")

    posterFile = os.path.join(dstDir, "Posters.csv")
    with open(posterFile, "w", encoding="utf-8") as f:
        f.write("File name, Category, Area, Last name, First name, "
                "University, Professor, Index, Revision, "
                "Last Modified, Title\n")

        for p in posters:
            title = ""
            try:
                pdfFile = os.path.abspath(os.path.join(dstDir, 
                                                       "Posters", 
                                                       p["category"], 
                                                       p["area"], 
                                                       p["bname"] + ".pdf"))
                title = getPosterTitle(pdfFile)
            except:
                title = " "

            f.write(f"{p['fname']}, "
                    f"{p['category']}, "
                    f"{p['area']}, "
                    f"{p['last']}, "
                    f"{p['first']}, "
                    f"{p['univ']}, "
                    f"{p['prof']}, "
                    f"{p['indx']}, "
                    f"{p['revs']}, "
                    f"{p['mtime']}, "
                    f"{title}\n")


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
    allFiles = getAllFiles(srcDir)

    if baseLink[-1] != '/':
        baseLink += '/'

    for f in allFiles:
        if verbose:
            print(f"Inserting QR code to {f}...", end=" ")

        base = os.path.basename(f)
        link = baseLink + base
        insertQRCode(f, link, location)

        if verbose:
            print("Success")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        root = os.path.abspath(os.path.join(os.path.expanduser('~'), 
                                            "Downloads",  "Site Visit"))
    else:
        root = sys.argv[1]

    timenow = datetime.datetime.now().strftime("%b %d %y %H%M%S")
    newDir = os.path.join(root, os.pardir, "NO UPLOAD Generated " + timenow)
    newDir = os.path.abspath(newDir)

    verbose = True

    if verbose:
        print(f"Scanning files in {root}...")

    papers, posters = scan(root)

    if verbose:
        print("Checking duplicates...")

    checkDuplicate(papers)
    checkDuplicate(posters)

    if verbose:
        print("No duplicate found, removing old revisions...")

    papers = removeOldRevisions(papers)
    posters = removeOldRevisions(posters)

    if verbose:
        print(f"Creating folder structure in {newDir}")

    os.makedirs(os.path.join(newDir, "Papers"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Core"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Non-core"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Associated"), exist_ok=True)

    if verbose:
        print("Copying files to formatted forlder structure...")

    copyFormated(papers, posters, newDir, verbose)

    if verbose:
        print("Converting files to pdf...")

    batch2pdf(newDir, verbose)

    if verbose:
        print("Inserting QR codes to posters...")

    year = datetime.datetime.now().strftime("%Y")
    # Base link where file should point to
    baseLink = "https://curent.utk.edu/" + year + "SiteVisit"
    # Location for placing the QR code on the poster
    QRLocation = [1400, 2415, 1540, 2555]

    batchQRCode(os.path.join(newDir, "Posters"), baseLink, QRLocation, verbose)

    if verbose:
        print(f"Creating submission files list in {newDir}...")

    saveList(papers, posters, newDir)

    if verbose:
        print("Done")