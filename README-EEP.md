
Quick Start
======================
* `make combine`
* Check and fix errors in the combined Excel file.
* Open up combine sorted file.  Fix names and save file.
* `make eeplist` or `make eeplist-t`.  `make eeplist-t`  generates Taiwan only schools and combines checking and letter submission lists.
* Delete all files on local machine /Users/caber/Documents/eep/tmp/doc
* Delete all files on local machine /Users/caber/Documents/eep/tmp/pdf
* Open virtualbox and run windows.
* Open student-labels-with-macro-for-windows-only.xls
* Execute macro "GenerateLabelsForSchools".
* Look at files under /Users/caber/Documents/eep/tmp/doc,pdf



Detailed Steps
======================
* pipenv shell

* cd src

* (Now optional) Init a directory for season.
	python eep_shared.py
	This creates a document_inspection folder.

* Make sure ROWS_USED_BY_HEADING in "eep-merge-sheets-from-raw-excel.py" is correct.  Default is 3.

* Take Wen's raw xlsm file, clean it up (change to readable font, 宋体) and make a copy.
	Save as {year}{seasonLetter}_eep.xls

* Open up {year}{seasonLetter}_eep.xls and clean it up too if needed.
	Clear ending rows for example.  Make sure China tab is first and Taiwan is 2nd

* Create year{seasonLetter}_eep_combined.xls.
	Use `make combine` if possible.  Otherwise, use following commands.

	./eep-merge-sheets-from-raw-excel.py --sheetnums 13 14
	python eep-merge-sheets-from-raw-excel.py  ~/Documents/eep/{year}{season}/{year}{season}_eep.xls --sheetnums {sheet_1_index} ... #0 based.
	./eep-merge-sheets-from-raw-excel.py /eep/2019s/2019s_eep.xls --sheetnums 13 14

	cd ~/Documents/eep/2012f
	/Users/cc/projects/eep-automation-suite/src/eep-merge-sheets-from-raw-excel.py --sheetnums 1 2
	/Users/cc/projects/eep-automation-suite/src/eep-merge-sheets-from-raw-excel.py ~/Documents/eep/2017s/2017s_eep.xls --sheetnums 31 32

	/Users/cc/Documents/eep/scripts/python_scripts/eep/eep-merge-sheets-from-raw-excel.py  ~/Documents/eep/2012f/2012f_eep.xls --sheetnums 29 30

* Open up x_eep_combined.xls, clean up(font problem if exists) and save using Excel program.
	This forces calculation of formulas that can be used in the next step.

* (This is now created automatically) Sort x_eep_combined.xls by:
	sch-na-len,
	school-na,
	donor-id
	Save as x_eep_combined_sorted.xls

	Clean up student-label-name column as needed.

* Open up x_eep_combined_sorted.xls in Excel, clean up for final output and resave the file.
	The file created automatically can't be used with mail merge for unknown reasons.
  NOTE: Check student_name_extra column, put the value back into name column if it was
	to distinguish a student from another.

* Create all the lists using:
  make eeplist
	-- If no Excel file is specified, application will attempt to look for current season's file.
	python ./eep-generate-lists.py  ~/Documents/eep/2013s/2013s_eep_combined_sorted.xls
	./eep-generate-lists.py /eep/2019s/2919s_eep_combined_sorted.xls



Open up word-tools/student-labels.doc to generate label files.
NOTE: Use BiauKaiTee if possible.
5. Use WORD older version. Use student-name-labels.docx to create student labels.  Avery US Letter 5164/8164 label template
	 - Filter by schl_na_len
	 - Merge to new document
	 - Inspect new document
	 - Save as PDF to _toPrint/student-labels as student-labels-0Xschar.pdf
	 - Notes:
	 - If update label is grayed out, make sure mail merge type is "label".  This can be done by
	   selecting "start mail merge" and then "label".

	regular name length, use 48pt
	if long names, use 26pt
	- Letters / Pt Size (新細明體):
	- 4 / 60pt
	- 5 / 49pt
	- 6 / 41
	- 7 / 35
	- 8 / 30
	- 9 / 27
	- 10+ / 24
	- 12 / 20
	- 13 / 18 (none after 2019)

May have to print to PDF for all files so format remains the same when printed at EEP office.

6. Print volunteer list.  Use eep/templates/volunteer-name-tags.docx.  Remember to point to the right donor file in the corresponding year.


Volunteer Name Tags
===============
If necessary, use the following two Excel mail merge files to generate
volunteer name tags.  Both Excel utilizes the same volunteer-names.xls loated
in et/cdata.

- tools/volunteer-labels.docx
- tools/volunteer-sticky-labels.docx

EEP Donor Report Generator
==========================
* Organiza student photos into the season directory
* Brighten photos if necessary
* Rename all extensions to lower case
  rename 's/\.JPG$/.jpg/' *.JPG
* Manually rename files with Chinese names
* Use eep-photo-cropper.py to rename photos
* Do a second pass with eep-photo-cropper.py to crop photos
* In Windows VM, use eep-photo-resizer.py to resize photos from host directory to VM directory
* In Windows VM, use eep-donor-report-generator.py to generate Word files
* Move the Word files out to host machine
* Check Word files for errors
