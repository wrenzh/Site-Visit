# Manage paper and poster submission for CURENT ERC Annual Site Visit
# Author: Wen Zhang (wzhang53@utk.edu)
# Version: 0.1.0
# Date: 10/11/2019
# License: MIT
#
# Functionalities:
# - Generate list of papers and posters
# - Convert doc/docx and ppt/pptx to pdf
# - Generate QR code for each poster
# - Insert QR code onto each poster
#
# Usage:
#   python3 SiteVisit.py submission-directory
# Argument:
# - submission-directory is the full path to the site visit folder
#
# Limitations:
# Only works on Windows platform and Office must be installed as PDF creation
# relies on Office Powerpoint and Word
# Note the PDF/A compliance option must be disabled to avoid Office crashing
# It is known some Visio drawings will prehibit PDF creation by Office. 
# Mannual editting them to figures is required.
#
# Detailed functionalities:
# The generated pdf files for each paper and poster will be placed under the
# following folder structure:
#
# Generate at **time**
# -> Papers
# ---> Hardware Testbed
# -----> Ma_Yiwei_UTK_Wang_1.pdf
# -----> ... ...
# ---> HVDC and FACTs
# ---> ... ...
# -> Posters
# ---> Core
# -------> Hardware Testbed
# -------> HVDC and FACTs
# -------> ... ...
# ---> Non-core
# -------> ... ...
# ---> Associated
# -------> ... ...
#
# Simply put, under Posters, there are 3 categories (Core, Non-core,
# Associated) and under each category, there are numerous detailed research
# areas (Hardware Testbed, HVDC and FACTs, etc.). Under Papers, there is no
# category differentiation and research areas are immediately under Papers.
#
# Expected file convention:
# The same filename convention for both papers and posters
# Lastname_Firstname_SchoolAbbreviation_ProfessorLastname_Index_RevisionNo
# e.g. Zhang_Wen_UTK_Wang_1_R0.pptx
# The index is used to differentiate several submissions from a same author
# as well as to correlate the poster with paper. For example,
# Zhang_Wen_UTK_Wang_1_R0.docx
# is the corresponding paper for the above poster.
# The revision number is to help keep track of the latest submission without
# risking deleting previous version. Using a revision number in submission
# to Google Drive is advised. For Confluence submission, it is not necessary
# because Confluence has built-in version control, which is not always true
# for Google Drive.
# Note the paper index and revision number are optional. If not given, the
# default numbers are assumed, shown in the above example.
# In reality, because how human work, the filenames are going to be messey.
# For now, multiple underscores and any space before or after underscore are
# allowed. Still, do expect to clean up manually.
#
# Duplication checking:
# It is expected that there will be similar submission filenames for the same
# paper/poster. For example, one may submit both the .doc and .pdf. In this
# case, the program will stop and ask for mannual intervention. Checking file
# creation/modification time is not trustworthy as it may be easily changed.
#
# Removing older revisions:
# Only the most recent revision is kept while older revisions are removed from
# the generated folder. The generated files also only contain everything but 
# the revision number.
#
# Barcode generation:
# Barcode will be inserted at [1400, 2415, 1540, 2555] with 120x120 pixels.
# This seems consistent with the given powerpoint template page size.
# The base url link, barcode location can be changed at the end of the file.
# For now, the base url is: https://curent.utk.edu/2019SiteVisit

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


def decode(filename):
    """Decode the filename with the expected filename convention"""
    # Replace multiple underscores and space with a single one and split
    fields = re.sub(r"\s*\_+\s*", "_", filename).split('_')

    if len(fields) not in range(4, 7):
        raise ValueError(f"{filename} uses wrong convention")

    last, first, univ, prof = fields[0], fields[1], fields[2], fields[3]
    indx = "1" if len(fields) < 5 else fields[4]
    revs = "R0" if len(fields) < 6 else fields[5].upper()

    return (last, first, univ, prof, indx, revs)


def categorize(filepath):
    """Find the poster/paper category and area"""
    # Category (only for posters)
    c = re.compile("Core|Non-core|Associated")
    # Research area (for both posters and papers)
    a = re.compile(r"Hardware\sTestbed|HVDC\sand\sFACTS|Large\sScale\sTestbed|"
                   r"Other\sCategories|Power\sConverter\sDesign\sand\sControl|"
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

    # Find all files
    allFiles = []
    for root, dirs, files in os.walk(srcDir):
        for f in files:
            allFiles.append(os.path.join(root, f))

    for f in allFiles:
        base = os.path.basename(f)
        path = os.path.dirname(f)
        name, ext = os.path.splitext(base)

        if ext.lower() == ".pdf":
            pass
        elif ext.lower() == ".ppt" or ext.lower() == ".pptx":
            if verbose:
                print(f"Converting {f} to pdf...")

            try:
                deck = powerpoint.Presentations.Open(f)
                # Powerpoint formatType code 32 for pdf
                deck.SaveAs(os.path.join(path, name + ".pdf"), 32)
                deck.Close()
            except comtypes.COMError:
                print(f"ERROR: Converting {f} failed")
                print("Please check if the file has Visio drawings")
                print("Also check PDF export settings (PDF/A must be off)")
                quit()

            if verbose:
                print(f"Conversion success, removing {f}")

            os.remove(f)
        elif ext.lower() == ".doc" or ext.lower() == ".docx":
            if verbose:
                print(f"Converting {f} to pdf...")

            try:
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
        elif ext.lower() == ".csv":
            # Likely the list of papers and posters
            pass
        else:
            raise ValueError(f"Unexpected filetype .{ext} in {path}")

    powerpoint.Quit()
    word.Quit()


def insertQRCode(srcPDF, dstPDF, link, location):
    """Insert the QR code to link on pdfFile"""
    img = qrcode.make(link)
    buf = io.BytesIO()
    img.save(buf)
    img_stream = buf.getvalue()

    pdfFile = fitz.open(srcPDF)
    pdfFile[0].insertImage(fitz.Rect(*location), stream=img_stream)
    pdfFile.save(dstPDF)
    pdfFile.close()


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
        path = os.path.dirname(f)
        link = baseLink + base
        tmp = os.path.join(path, "tmp.pdf")
        insertQRCode(f, tmp, link, location)

        if verbose:
            print(f"Success, replacing original file")

        os.replace(tmp, f)


def generateList(srcDir, dstDir, verbose):
    """Process papers and posters from root folder"""
    allFiles = []
    for root, _, files in os.walk(srcDir):
        for f in files:
            allFiles.append(os.path.join(root, f))

    # Determine if a file belong to paper or poster and record in dictionary
    posters = []
    papers = []
    # Used to separate revision number
    pattern = re.compile(r"(\w+)_R(\d+)")
    for f in allFiles:
        base = os.path.basename(f)
        path = os.path.dirname(f)
        name, ext = os.path.splitext(base)

        isPoster = "Poster" in path
        isPaper = "Paper" in path

        if ext.lower() == ".doc" or ext.lower() == ".docx":
            # .doc/.docx is definitely a paper
            isPaper = True
        elif ext.lower() == ".ppt" or ext.lower() == ".pptx":
            # .ppt/.pptx is definitely a poster
            isPoster = True
        elif ext.lower() == ".pdf":
            # Determine by other files in the same folder
            # If there is any other file with .doc/.docx, the file is paper
            # Otherwise, the file is poster
            if not (isPoster or isPaper):
                files = os.listdir(path)
                paperHere = False
                for f in files:
                    if ".doc" in f.lower() or ".docx" in f.lower():
                        paperHere = True
                        break
                isPaper = paperHere
                isPoster = not paperHere
        else:
            raise ValueError(f"Unsupported file format {ext} in {base}")

        if isPoster and isPaper:
            raise ValueError(f"Confused {f} for paper or poster?")

        category, area = categorize(path)
        properName = '_'.join(decode(name))
        results = pattern.search(properName)
        if isPoster:
            posters.append({"Path": f,
                            "Name": properName,
                            "BaseName": results.group(1),
                            "Revision": int(results.group(2)),
                            "Extension": ext,
                            "Category": category,
                            "Area": area})
        else:
            papers.append({"Path": f,
                           "Name": properName,
                           "BaseName": results.group(1),
                           "Revision": int(results.group(2)),
                           "Extension": ext,
                           "Area": area})

    return papers, posters


def removeOldRevisions(fileList, verbose):
    """Find and delete older revisions"""
    baseNames = [p["BaseName"] for p in fileList]
    revisions = [p["Revision"] for p in fileList]

    # Sorted indices with baseNames to aggeregate same paper/poster
    sortedIndx = sorted(range(len(baseNames)), key=baseNames.__getitem__)
    sortedNames = [baseNames[i] for i in sortedIndx]
    sortedRevs = [revisions[i] for i in sortedIndx]

    prevName = ''
    prevRev = -1
    indicesToDelete = []
    for i, n in enumerate(sortedNames):
        if n != prevName:
            prevName = n
            prevRev = sortedRevs[i]
        elif sortedRevs[i] > prevRev:
            if verbose:
                print(f"Removing version {prevRev} < latest "
                        f"{sortedRevs[i]} for {n}")

            indicesToDelete.append(sortedIndx[i-1])
            prevRev = sortedRevs[i]
        elif sortedRevs[i] < prevRev:
            if verbose:
                print(f"Removing version {sortedRevs[i]} < latest "
                      f"{prevRev} for {n}")

            indicesToDelete.append(sortedIndx[i])

    return [fileList[i] for i in set(sortedIndx) - set(indicesToDelete)]


def findDuplicate(fileList):
    """Find duplicate submissions in fileList"""
    duplicate = False

    uniqNames = []
    for f in fileList:
        if f["Name"] not in uniqNames:
            uniqNames.append(f["Name"])
        else:
            duplicate = True
            print(f"Duplicate found: {f['Path']}")

    if duplicate:
        raise ValueError("Duplicate submission file found.")


def copyFormated(papers, posters, dstDir, verbose=True):
    """Copy files in papers and posters to dstDir as structured"""
    # Copy files to dstDir
    for p in papers:
        dstSubdir = os.path.join(dstDir, "Papers", p["Area"])
        dstFile = os.path.join(dstSubdir, p["BaseName"] + p["Extension"])
        os.makedirs(dstSubdir, exist_ok=True)
        shutil.copy2(p["Path"], dstFile)

        if verbose:
            path = p["Path"]
            print(f"Paper {path} copied to {dstFile}")

    for p in posters:
        dstSubdir = os.path.join(dstDir, "Posters", p["Category"], p["Area"])
        dstFile = os.path.join(dstSubdir, p["BaseName"] + p["Extension"])
        os.makedirs(dstSubdir, exist_ok=True)
        if os.path.isfile(dstFile):
            # File already exists, likely duplicates
            raise ValueError(f"File already exists: {dstFile}")

        shutil.copy2(p["Path"], dstFile)

        if verbose:
            path = p["Path"]
            print(f"Poster {path} copied to {dstFile}")



def writeList2CSV(papers, posters, dstDir):
    """Write list of papers and posters to csv files in dstDir"""
    try:
        paperFile = open(os.path.join(dstDir, "Papers.csv"), "w")
        posterFile = open(os.path.join(dstDir, "Posters.csv"), "w")
    except OSError:
        print("Cannot write .csv files. Are they open in Excel?")

    # Print table headers
    print("File name, Area, Last name, First name, "
          "University, Professor, Index, Revision", file=paperFile)
    print("File name, Category, Area, Last name, First name, "
          "University, Professor, Index, Revision", file=posterFile)

    for p in papers:
        name = p["Name"]
        area = p["Area"]
        last, first, univ, prof, indx, revs = decode(name)
        print(f"{name}, {area}, {last}, {first}, "
              f"{univ}, {prof}, {indx}, {revs}", file=paperFile)

    for p in posters:
        name = p["Name"]
        category = p["Category"]
        area = p["Area"]
        last, first, univ, prof, indx, revs = decode(name)
        print(f"{name}, {category}, {area}, {last}, {first}, "
              f"{univ}, {prof}, {indx}, {revs}", file=posterFile)

    paperFile.close()
    posterFile.close()


def main(rootDir, newDir, baseLink, QRLocation, verbose):
    """Organize and export files for Site Visit"""
    if verbose:
        print("Preprocessing original files...")

    papers, posters = generateList(rootDir, newDir, verbose)

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
        print(f"Creating submission files list in {newDir}...")

    writeList2CSV(papers, posters, newDir)

    if verbose:
        print("Copying to formatted forlder structure")

    copyFormated(papers, posters, newDir, verbose)

    if verbose:
        print("Converting files to pdf...")

    batch2pdf(newDir, verbose)

    if verbose:
        print("Inserting QR codes for posters...")

    batchQRCode(os.path.join(newDir, "Posters"), baseLink, QRLocation, verbose)

    if verbose:
        print("Done")


if __name__ == "__main__":
    year = datetime.datetime.now().strftime("%Y")
    if len(sys.argv) < 2:
        root = os.path.join(os.path.expanduser('~'), 
                            "Downloads", year + " Annual Site Visit")
    else:
        root = sys.argv[1]

    timenow = datetime.datetime.now().strftime("%b %d %y %H%M%S")
    newDir = os.path.abspath(os.path.join(root, os.pardir, 
                                          "Generated " + timenow))

    # Base link where file should point to
    baseLink = "https://curent.utk.edu/" + year + "SiteVisit"
    # Location for placing the QR code on the poster
    QRLocation = [1400, 2415, 1540, 2555]

    main(root, newDir, baseLink, QRLocation, True)