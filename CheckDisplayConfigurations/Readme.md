# CheckDisplayConfigurations

Reports assembly display configurations that are out of date.

When a part is deleted from an assembly, previously saved display configurations may still refer to it. Solid Edge doesn't flag it at the time. You only find out when you activate one and get the warning *One or more parts have been deleted from the assembly since the configuration was saved.* 

This program finds them without having to activate each configuration by hand.  It checks every level of the assembly because a part deleted from a subassembly also makes the higher-level configurations out of date.

For each out-of-date configuration, a message like this is reported:

`Display configuration 'ELEVATOR INSTALLATION' out of date`

To fix it, open the assembly in Solid Edge, activate the reported configuration, check that it looks right, and save the updated result.

## Notes

- The configurations are read from the `.cfg` file that Solid Edge creates for every assembly, with the same name. If that file is missing, an error is reported. Configuration files with a different name are not supported.
- The program doesn't change the assembly or the `.cfg` file. Solid Edge locks the `.cfg` file while the assembly is open, so the program reads a temporary copy.
- Family of Assemblies files are not currently supported.
- Configurations saved in very old versions of Solid Edge use a format the program can't read, and are reported as not checked. To update one, in SE activate the display configuration and save the assembly.
- If a subassembly's file can't be found, it is reported in the log file.  Configurations aren't checked below that point.

## Troubleshooting

Details are written to `configuration_diagnostic.txt` in the program's directory. For each configuration, it lists every saved occurrence ID, what it matches in the current assembly, and which ones are missing.

The `.cfg` file format was worked out by experiment.  The logic is described in `ConfigurationFile.vb`. If a file doesn't match a known format, the program reports an error, for example:

`Could not read 'MyAssembly.cfg': Configuration 'VIEW 1': Unrecognized node at byte offset 12708: flags 0x30000004, second value 0x00000000, ID 0x00000232`

If you get one of those, please report it.  We will need the offending `.cfg` file to investigate.
