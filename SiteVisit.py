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
import datetime
import comtypes.client
from pdfQRcode import batchQRCode
from pdfTitle import getPosterTitle


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
        print(indx)
        print(revs)
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


def batch2pdf(srcDir, verbose):
    """Autocatically replace ppt/doc with pdf in srcDir"""
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    word = comtypes.client.CreateObject("Word.Application")

    allFiles = []
    for root, dirs, files in os.walk(srcDir):
        for f in files:
            allFiles.append(os.path.join(root, f))

    for f in allFiles:
        path = os.path.dirname(f)
        name, ext = os.path.splitext(os.path.basename(f))

        if ext.lower() == ".pdf":
            pass
        elif ext.lower() == ".csv":
            # Likely the list of papers and posters so don't do anything
            pass
        elif ext.lower() == ".ppt" or ext.lower() == ".pptx" or \
             ext.lower() == ".doc" or ext.lower() == ".docx":
            if verbose:
                print(f"Converting {f} to pdf...")

            try:
                if ext.lower() == ".ppt" or ext.lower() == ".pptx":
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
                print(f"ERROR: Converting {f} failed")
                print("Please check if the file has Visio drawings")
                print("Also check PDF export settings (PDF/A must be off)")
                quit()

            if verbose:
                print(f"Conversion success, removing {f}")

            os.remove(f)
        else:
            raise ValueError(f"Unexpected filetype .{ext} in {path}")


def generateList(srcDir, dstDir, verbose):
    """Process papers and posters from root folder"""
    allFiles = []
    for root, _, files in os.walk(srcDir):
        for f in files:
            allFiles.append(os.path.join(root, f))

    # Determine if a file is paper or poster and record in dictionary
    posters = []
    papers = []
    for f in allFiles:
        base = os.path.basename(f)
        path = os.path.dirname(f)
        name, ext = os.path.splitext(base)

        isPoster = "Poster" in path or \
                   ext.lower() == ".ppt" or \
                   ext.lower() == ".pptx"
        isPaper = "Paper" in path or \
                  ext.lower() == ".doc" or \
                  ext.lower() == ".docx"

        if (not (isPoster or isPaper)) and ext.lower() == ".pdf":
            # Determine by other files in the same folder
            for a in os.listdir(path):
                isPaper |= ".doc" in a.lower()
                isPoster |= ".ppt" in a.lower()

        if (isPoster and isPaper) and (not (isPoster or isPaper)):
            raise ValueError(f"Cannot determine {f}: paper or poster?")

        category, area = categorize(path)
        last, first, univ, prof, indx, revs = decode(name)
        # Formatted name
        fname = '_'.join([last, first, univ, prof, indx, revs])
        # Base name
        bname = '_'.join([last, first, univ, prof, indx])
        revs = int(revs[1:])
        mtime = os.path.getmtime(f)
        d = {"file": f, "name": fname, "base": bname, "revs": revs, 
             "ext": ext, "category": category, "area": area, "mtime": mtime}

        if isPoster:
            posters.append(d)
        else:
            papers.append(d)

    return papers, posters


def removeOldRevisions(fileList, verbose):
    """Find and delete older revisions"""
    baseNames = [f["base"] for f in fileList]
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
            if verbose:
                print(f"Remove ver. {lastRev} < latest {sRevs[i]} for {n}")

            indxToDel.append(sIndx[i-1])
            lastRev = sRevs[i]
        elif sRevs[i] < lastRev:
            if verbose:
                print(f"Remove ver. {sRevs[i]} < latest {lastRev} for {n}")

            indxToDel.append(sIndx[i])

    return [fileList[i] for i in set(sIndx) - set(indxToDel)]


def findDuplicate(fileList):
    """Find duplicate submissions in fileList"""
    duplicate = False

    uniqs = []
    for f in fileList:
        if f["name"] not in uniqs:
            uniqs.append(f["name"])
        else:
            duplicate = True
            print(f"Duplicate found: {f['file']}")

    if duplicate:
        raise ValueError("Duplicate submission file found")


def copyFormated(papers, posters, dstDir, verbose=True):
    """Copy files in papers and posters to dstDir as structured"""
    for p in papers:
        # No need to differentiate categories for papers
        dstSubdir = os.path.join(dstDir, "Papers", p["area"])
        os.makedirs(dstSubdir, exist_ok=True)
        # Revision numbers are removed
        dstFile = os.path.join(dstSubdir, p["base"] + p["ext"])
        if os.path.isfile(dstFile):
            # File already exists, likely duplicates or older submission
            raise ValueError(f"File already exists: {dstFile}")

        shutil.copy2(p["file"], dstFile)

        if verbose:
            print(f"Paper {p['file']} copied to {dstFile}")

    for p in posters:
        # Categories for posters
        dstSubdir = os.path.join(dstDir, "Posters", p["category"], p["area"])
        os.makedirs(dstSubdir, exist_ok=True)
        # Revision numbers are removed
        dstFile = os.path.join(dstSubdir, p["base"] + p["ext"])
        if os.path.isfile(dstFile):
            # File already exists, likely duplicates or older submission
            raise ValueError(f"File already exists: {dstFile}")

        shutil.copy2(p["file"], dstFile)

        if verbose:
            print(f"Poster {p['file']} copied to {dstFile}")


def writeList2CSV(papers, posters, dstDir):
    """Write list of papers and posters to csv files in dstDir"""
    try:
        paperFile = open(os.path.join(dstDir, "Papers.csv"), "w")
        posterFile = open(os.path.join(dstDir, "Posters.csv"), "w")
    except OSError:
        print("Cannot write .csv files. Are they open in Excel?")

    # Print table headers
    print("File name, Area, Last name, First name, University, "
          "Professor, Index, Revision, Last Modified, Title", file=paperFile)
    print("File name, Category, Area, Last name, First name, University, "
          "Professor, Index, Revision, Last Modified, Title", file=posterFile)

    for p in papers:
        last, first, univ, prof, indx, _ = decode(p["name"])
        print(f"{p['name']}, {p['area']}, {last}, {first}, {univ}, "
              f"{prof}, {indx}, {p['revs']}, {p['mtime']}, ",
              file=paperFile)

    for p in posters:
        last, first, univ, prof, indx, _ = decode(p["name"])
        try:
            pdfFile = os.path.abspath(os.path.join(dstDir, "Posters", 
                                                p["category"], 
                                                p["area"], 
                                                p["base"] + ".pdf"))
            title = getPosterTitle(pdfFile, [0, 0, 1728, 290])
            print(f"{p['name']}, {p['category']}, {p['area']}, {last}, {first}, "
                f"{univ}, {prof}, {indx}, {p['revs']}, {p['mtime']}, {title}", 
                file=posterFile)
        except:
            print(f"{p['name']}, {p['category']}, {p['area']}, {last}, {first}, "
                f"{univ}, {prof}, {indx}, {p['revs']}, {p['mtime']}, ", 
                file=posterFile)

    paperFile.close()
    posterFile.close()


if __name__ == "__main__":
    year = datetime.datetime.now().strftime("%Y")
    if len(sys.argv) < 2:
        root = os.path.join(os.path.expanduser('~'), "Downloads", 
                            year + " Annual Site Visit")
    else:
        root = sys.argv[1]

    timenow = datetime.datetime.now().strftime("%b %d %y %H%M%S")
    newDir = os.path.join(root, os.pardir, "NO UPLOAD Generated " + timenow)
    newDir = os.path.abspath(newDir)

    verbose = True

    if verbose:
        print("Preprocessing original files...")

    papers, posters = generateList(root, newDir, verbose)

    if verbose:
        print("Looking for duplicate submissions...")

    findDuplicate(papers)
    findDuplicate(posters)

    if verbose:
        print("No duplicate found, looking for old revisions...")

    papers = removeOldRevisions(papers, verbose)
    posters = removeOldRevisions(posters, verbose)

    if verbose:
        print(f"Creating folder structure in {newDir}")

    os.makedirs(os.path.join(newDir, "Papers"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Core"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Non-core"), exist_ok=True)
    os.makedirs(os.path.join(newDir, "Posters", "Associated"), exist_ok=True)

    if verbose:
        print("Copying to formatted forlder structure...")

    copyFormated(papers, posters, newDir, verbose)

    if verbose:
        print("Converting files to pdf...")

    batch2pdf(newDir, verbose)

    if verbose:
        print("Inserting QR codes for posters...")

    # Base link where file should point to
    baseLink = "https://curent.utk.edu/" + year + "SiteVisit"
    # Location for placing the QR code on the poster
    QRLocation = [1400, 2415, 1540, 2555]

    batchQRCode(os.path.join(newDir, "Posters"), baseLink, QRLocation, verbose)

    if verbose:
        print(f"Creating submission files list in {newDir}...")

    writeList2CSV(papers, posters, newDir)

    if verbose:
        print("Done")