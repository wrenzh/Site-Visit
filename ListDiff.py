# Compare two list of posters or papers 
# Author: Wen Zhang (wzhang53@utk.edu)
# Version: 0.2.0
# Date: 10/17/2019
# License: MIT

import os
import sys
import csv
import re

def listDiff(A, B):
    """Compare two list of files and find changes"""
    # A is the reference, and B is supposedly the newer version
    a = list(csv.DictReader(open(A, 'r'), skipinitialspace=True))
    b = list(csv.DictReader(open(B, 'r'), skipinitialspace=True))

    # Get name keys
    getName = re.compile(r"(.*)_R\d+")
    namesA = [getName.match(f["File name"]).group(1) for f in a]
    namesB = [getName.match(f["File name"]).group(1) for f in b]

    # Get category and area
    catA = [f["Category"] for f in a]
    catB = [f["Category"] for f in b]
    areaA = [f["Area"] for f in a]
    areaB = [f["Area"] for f in b]

    # Get last modified time
    mtimeA = [float(f["Last Modified"]) for f in a]
    mtimeB = [float(f["Last Modified"]) for f in b]

    # Find new and modified submissions
    newSubs = []
    contentModifies = []
    categoryChanges = []
    for i, n in enumerate(namesB):
        if n not in namesA:
            newSubs.append(n)
        else:
            indxA = namesA.index(n)
            if mtimeA[indxA] != mtimeB[i]:
                contentModifies.append(n)
            if catA[indxA] != catB[i] or areaA[indxA] != areaB[i]:
                categoryChanges.append(n)

    # Find deleted submissions
    delSubs = []
    for n in namesA:
        if n not in namesB:
            delSubs.append(n)

    # Report changes
    print(f"Total new submissions in {B}: {len(newSubs)}")
    for n in newSubs:
        print(f"+ {n}")

    print(f"\nDeleted submissions in {A}: {len(delSubs)}")
    for n in delSubs:
        print(f"- {n}")

    print(f"\nContent modified in {B}: {len(contentModifies)}")
    for n in contentModifies:
        print(f"@ {n}")

    print(f"\nCategory change in {B}: {len(categoryChanges)}")
    for n in categoryChanges:
        print(f"# {n}")



if __name__ == "__main__":
    if len(sys.argv) > 3:
        raise ValueError("Too many input arguments")
    elif len(sys.argv) == 3:
        listDiff(sys.argv[1], sys.argv[2])
    elif len(sys.argv) <= 2:
        folder = os.path.join(os.path.expanduser('~'), 'Downloads') \
                 if (len(sys.argv) == 1) else sys.argv[1]

        csvFiles = []
        for f in os.listdir(folder):
            _, ext = os.path.splitext(f)
            if ext.lower() == ".csv":
                csvFiles.append(os.path.join(folder, f))

        if len(csvFiles) != 2:
            raise ValueError("Exactly 2 csv files required")

        listDiff(csvFiles[0], csvFiles[1])