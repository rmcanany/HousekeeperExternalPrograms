# Replace Blocks

Example program for the Solid Edge Housekeeper `Run External Program` task.  The program has three operations it can perform: `Find/Replace`, `Add`, and `Delete`.  The operations are performed in the order given.

The program reads `program_settings.txt` for the processing parameters.  Edit that file according to your needs.  Here is an example.

```
TemplateName = C:\data\CAD\BlockLibrary.dft

ReplaceBlock = Block1, Block1
ReplaceBlock = TitleBlock, NewTitleBlock

AddBlock = square
AddBlock = circle

DeleteBlock = slot
DeleteBlock = triangle
```

- The `TemplateName` variable tells the program what file contains the replacement blocks, or blocks to be added.
- The `ReplaceBlock` variable is a comma-delimited tuple.  The first entry in the tuple is the name of the block to be replaced.  The second is the name of its replacement.  There can be any number `ReplaceBlock` entries in the file.
- The `AddBlock` and `DeleteBlock` variables define what blocks are to be processed for their respective action.  There can be any number of these, as well.
 
There are a couple of limitations in the current implementation.  

First, it cannot replace a differently-named block if a block with the new name already exists in the file.  That situation is reported in the log file.

Second, the `Find/Replace` option can't process block names that contain a comma.  That is also reported in the log file.