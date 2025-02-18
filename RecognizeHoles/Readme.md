# Recognize Holes

Finds cylindrical cutouts in an imported model and converts them into hole features. For the command to work correctly, the model must be in a freshly-imported state, with no subsequent modifications performed in Solid Edge. 

As the first step of the conversion process, the Optimize command is run on the imported geometry. While not strictly necessary, it is considered good practice for any imported file. 

The conversion is only possible in Synchronous mode. Ordered files are switched to Sync before the conversion, then switched back. Note, the imported body and the new hole features remain in Sync after the transition.
