# Rename Sheets
Original version by Tushar Suradkar

https://community.sw.siemens.com/s/question/0D54O00006BtAnKSAV/code-rename-sheets-to-the-referenced-model-name

This program renames the sheets in a draft to the model name 
of the first drawing view on the sheet.  If the drawing view 
is set to match a configuration, the configuration name is 
appended to the model name.

If multiple sheets would otherwise have the same name, the 
program appends a suffix `-Copy(X)` to the name.  Where `X` 
is a sequential integer value starting at `1`.  The suffix 
text `Copy` is a string read from the file `suffix.txt`. 
It can be changed to whatever is appropriate for your 
circumstances.

Sheets with no model drawing views are not renamed.

If the program has processed the file previously, the above 
scheme does not work as expected.  To avoid this issue, the 
sheets with model views are first renamed to random integer 
values, then processed as above.