# CompareFlatAndModelVolumes
**Algorithm contributed by @o_o ....码**  Thank you!

Computes the difference in volume of a bent sheetmetal part and its flat pattern.

The program addresses an issue where a flat pattern is created in the Synchronous environment.  If Ordered features are then added, they are not carried over to the flat pattern. Compounding the problem, even though it is out-of-date, the flat pattern is not flagged as such.

The program performs the following calculation:

`NVD = (|MV - FPV|) / MV`, where
- NVD is the normalized volume difference
- MV is the model volume
- FPV is the flat pattern volume
- || indicates absolute value

The normalized volume difference is compared to a threshold specified by the user.  The threshold is set in `program_settings.txt`.  If `NVD > Threshold`, an error is reported.

To avoid some complications, the file must contain exactly one model body and one flat pattern.

Setting the threshold requires some experimentation. That is because `NVD` depends on the neutral factor. It also depends on the number of bends relative to the overall size of the part.  In a part with no bends, `NVD` should be very close to 0.  In a small part with many bends, `NVD` will be orders of magnitude higher. Finally, it depends on whether you add or remove material from the flat pattern itself.

One way to do the experiment is to set the threshold to a tiny value, maybe something like `1e-15`. Then run the program and inspect the NVD values reported in the log file.  You can then set the threshold to a practical value, maybe close to the high end of the numbers you saw.

As you can see, this method is not foolproof. You can get false positives and/or false negatives. However, if you are running into this issue with flat patterns, it can be a useful check on your files.


