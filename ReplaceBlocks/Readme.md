# Replace Blocks

Example program for the Solid Edge Housekeeper `Run External Program` task.  The program scans the file for the presence of named blocks and replaces them if found.  

The program reads `program_settings.txt` for the processing parameters.  Edit that file according to your needs.  Here is an example.

```
TemplateName = C:\data\CAD\BlockLibrary.dft

ReplaceBlock = Block1, Block1
ReplaceBlock = TitleBlock, NewTitleBlock
```

- The `TemplateName` variable tells the program what file contains the replacement blocks.
- The `ReplaceBlock` variable is a comma-delimited tuple.  The first entry is the name of the block to be replaced.  The second is the name of its replacement.  There can be any number `ReplaceBlock` entries in the file.
 
There are a couple of limitations in the current implementation.  

First, it cannot replace a differently-named block if a block with the new name already exists in the file.  That situation is reported in the log file.

Second, it will have trouble with any block name that contains a comma.