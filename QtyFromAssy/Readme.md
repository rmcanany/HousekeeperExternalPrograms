# QtyFromAssy

Gets the total quantity of each model in a specified assembly and
subassemblies.  

The program checks the `IncludeInBOM` flag and does not process files 
where it is set to `NO`.  It also checks for quantity overrides and 
proceeds accordingly.

The quantity and assembly file name are added as custom properties to 
each model, making them available, for example, in a Draft file.  The 
property names are `QtyFromAssy_Qty` and `QtyFromAssy_Assy`, 
holding the quantity and assembly file name, respectively.

The program can be run from Solid Edge Housekeeper or stand-alone.  In 
stand-alone mode, the assembly file must be open in Solid Edge before 
starting the command.