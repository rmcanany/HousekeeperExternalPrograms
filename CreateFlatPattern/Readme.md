# Create Flat Pattern

Example program for the Solid Edge Housekeeper `Run External Program` command.  

Creates a flat pattern of a formed sheet metal part.  Processes `Ordered` and `Sync` parts.  Processes `*.psm` and `*.par` parts as long as the `SheetMetal` environment is active.

Selects the largest planar face and the longest linear edge of that face for creation and placement.

The user is alerted through Housekeeper's reporting mechanism if:
- The file's active environment is not `SheetMetal`
- A part with no geometry is encountered
- A flat pattern already exists in the file
- Problems processing the geometry are encountered
- No planar faces and/or linear edges are found

