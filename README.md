Enlightenment Foundation tools
===============
This project provides a few tools that I use to semi-automate some tedious and
repetitive tasks at the organization that I volunteer at.  The script mainly
deal with reading and creating Excel files that are needed every six months.

While these tools work, they do require a LOT of cleanup and refactoring.


Requirements
===============
* Python 2.7.
* xlrd, xlwt and xlutils library.
  Source: https://pypi.python.org/pypi/xlrd
          https://pypi.python.org/pypi/xlwt
          https://pypi.python.org/pypi/xlutils

Tool Specific Requirements
===============
* eep-donor-report-generator.py
  - Requires MS Office installed.
  - Windows (tested with Windows 7) as it uses Windows COM APIs to automate Word
    file creation.

Available Scripts
===============
* src/
    * eep-donor-report-generator.py (Windows only)
    * eep-generate-lists.py
    * eep-merge-sheets-from-raw-excel.py
    * eep-photo-cropper.py
    * eepphotoresizer.py

Either do pipenv shell first and invoke the individual scripts.
Or use `pipenv run`. e.g. pipenv run python src/eep-photo-cropper.py

Usage
===============
I'm not going to provide usage here (at least not now) as these tools are made
specifically for my organization and I doubt other people can use it without
heavy modifications.  But if you are curious as to how I use these tools, refer
to the readme.txt.  That's my personal cheat sheet.
