Steps
======================
* (Now optional) Init a directory for season.
	python eep_shared.py

* Make sure ROWS_USED_BY_HEADING in "eep-merge-sheets-from-raw-excel.py" is correct.  Default is 3.

* Take Wen's raw xlsm file, clean it up (change to readable font, 宋体) and make a copy.
	Save as {year}{seasonLetter}_eep.xls

* Open up {year}{seasonLetter}_eep.xls and clean it up too if needed.
	Clear ending rows for example.  Make sure China tab is first and Taiwan is 2nd

* Create year{seasonLetter}_eep_combined.xls.
	python eep-merge-sheets-from-raw-excel.py  ~/Documents/eep/{year}{season}/{year}{season}_eep.xls --sheetnums {sheet_1_index} ... #0 based.

	cd ~/Documents/eep/2012f
  /Users/cc/Documents/eep/scripts/python_scripts/eep/eep-merge-sheets-from-raw-excel.py  ~/Documents/eep/2012f/2012f_eep.xls 16,17

* Open up x_eep_combined.xls, clean up(font problem if exists) and save using Excel program.
	This forces calculation of formulas that can be used in the next step.

* (This is now created automatically) Sort x_eep_combined.xls by:
	sch-na-len,
	school-na,
	donor-id
	Save as x_eep_combined_sorted.xls

* Open up x_eep_combined.xls in Excel and resave the file.
	The file created automatically can't be used with mail merge for unknown reasons.

* Create all the lists using:
	python ./eep-generate-lists.py  ~/Documents/eep/2013s/2013s_eep_combined_sorted.xls
	This creates a document_inspection folder.



Open up tools/student-labels.doc to generate label files.
NOTE: Use BiauKaiTee if possible.
5. Use WORD older version. Use student-name-labels.docx to create student labels.  Avery 5164/8164 label template
	 - Filter by schl_na_len
	 - Merge to new document
	 - Inspect new document
	 - Save as PDF to _toPrint/student-labels as student-labels-0Xschar.pdf

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
		- 13 / 18

May have to print to PDF for all files so format remains the same when printed at EEP office.

6. Print volunteer list.  Use eep/templates/volunteer-name-tags.docx.  Remember to point to the right donor file in the corresponding year.


Volunteer Name Tags
===============
If necessary, use the following two Excel mail merge files to generate
volunteer name tags.  Both Excel utilizes the same volunteer-names.xls loated
in et/cdata.

- tools/volunteer-labels.docx
- tools/volunteer-sticky-labels.docx