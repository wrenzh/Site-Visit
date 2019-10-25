# Site-Visit

UTK CURENT ERC Annual Site Visit Paper and Poster Submission Processing

## Functionalities

- Generate list of papers and posters
- Convert doc/docx and ppt/pptx to pdf
- Generate QR code for each poster
- Insert QR code onto each poster

## Usage

`python3 SiteVisit.py submission-directory`

- submission-directory is the full path to the site visit folder

## Limitations

Only works on Windows platform and Office must be installed as PDF creation
relies on Office Powerpoint and Word "Save as" functionality.
Note the PDF/A compliance option must be disabled to avoid Office crashing.
It is known some Visio drawings will prehibit PDF creation.
The Powerpoint and Word applications must be maximized for some reason for
some conversion.

## Detailed description

### Output folder structure

The generated pdf files for each paper and poster will be placed under the
following folder structure:

Generate at **time**\
--> Papers\
----> Hardware Testbed\
------> Ma_Yiwei_UTK_Wang_1.pdf\
----> HVDC and FACTs\
----> ... ...\
--> Posters\
----> Core\
--------> Hardware Testbed\
--------> HVDC and FACTs\
--------> ... ...\
----> Non-core\
----> Associated\

Simply put, under Posters, there are 3 categories (Core, Non-core,
Associated) and under each category, there are numerous detailed research
areas (Hardware Testbed, HVDC and FACTs, etc.). Under Papers, there is no
category differentiation and research areas are immediately under Papers.

### Expected filename convention

The same filename convention for both papers and posters
Lastname_Firstname_SchoolAbbreviation_ProfessorLastname_Index_RevisionNo\
e.g. Zhang_Wen_UTK_Wang_1_R0.pptx\
The index is used to differentiate several submissions from a same author
as well as to correlate the poster with paper. For example,\
Zhang_Wen_UTK_Wang_1_R0.docx\
is the corresponding paper for the above poster.
The revision number is to help keep track of the latest submission without
risking deleting previous version. Using a revision number in submission
to Google Drive is advised. For Confluence submission, it is not necessary
because Confluence has built-in version control, which is not always true
for Google Drive.
Note the paper index and revision number are optional. If not given, the
default numbers are assumed, shown in the above example.
In reality, because how human work, the filenames are going to be messy.
Do expect to clean up manually.

### Duplication

It is expected that there will be similar submission filenames for the same
paper/poster. For example, one may submit both the .doc and .pdf. In this
case, the program will stop and ask for mannual intervention. Checking file
creation/modification time is not trustworthy as it can be easily changed.

### Removing older revisions

Only the most recent revision is kept while older revisions are removed from
the generated folder. The generated filename also does not include the
revision number.

### Barcode generation

Barcode will be inserted at [1400, 2415, 1540, 2555] with 120x120 pixels.
This seems consistent with the given powerpoint template page size.
The base url link, barcode location can be changed at the end of the file.
For the year of 2019, the base url is: <https://curent.utk.edu/2019SiteVisit>
