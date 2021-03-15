# Readme.md
Original version by Tushar Suradkar
https://community.sw.siemens.com/s/question/0D54O00006BtAnKSAV/code-rename-sheets-to-the-referenced-model-name

This program renames the sheets in a draft to the model name of the first drawing view on the sheet.

If multiple sheets have the same first model, it appends "Copy(X)" to the name.  Where X is a sequential integer value starting at 1.

Sheets with no model drawing views are not renamed.

It is possible that the program has already been run on the file.  In that case, the above scheme does not work as expected.  To avoid this problem, the sheets with model views are first renamed to random integer values, then processed as above.