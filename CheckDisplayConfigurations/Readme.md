# CheckDisplayConfigurations

Reports assembly display configurations that are out of date.

When a part is deleted from an assembly, display configurations saved before the deletion still refer to it. Solid Edge doesn't flag those configurations. You only find out when you activate one and get the warning *One or more parts have been deleted from the assembly since the configuration was saved.* The program finds them without having to activate each configuration by hand.

It checks every level of the assembly. A part deleted from a subassembly makes the top-level assembly's configurations out of date too, and is reported the same way.

For each out-of-date configuration, a message like this is reported:

`Display configuration 'ELEVATOR INSTALLATION': One or more parts have been deleted from the assembly since the configuration was saved`

To fix it, activate the configuration, check that it looks right, and save it again.

## Notes

- The `default,Solid Edge` configuration is updated every time the assembly is saved. It is only reported if the assembly has changed since it was last saved.
- The configurations are read from the `.cfg` file Solid Edge keeps next to the assembly, with the same name. If that file is missing, an error is reported. Configuration files with a different name are not supported.
- The program doesn't change the assembly or the `.cfg` file. Solid Edge locks the `.cfg` file while the assembly is open, so the program reads a temporary copy.
- Family of Assemblies files are not currently supported.

## Troubleshooting

Details are written to `configuration_diagnostic.txt` in the program's directory. For each configuration, it lists every saved occurrence ID, what it matches in the current assembly, and which ones are missing.

The `.cfg` file format is not documented by Siemens. It was worked out by experiment, and the details are in the comments at the top of `ConfigurationFile.vb`. If a file doesn't match the known format, the program reports an error instead of guessing, for example:

`Could not read 'MyAssembly.cfg': Configuration 'VIEW 1': Unrecognized node at byte offset 12708: flags 0x30000004, second value 0x00000000, ID 0x00000232`

If you get one of those, please report it, along with the version of Solid Edge used.
