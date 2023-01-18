# Costing Data Project
A software made to automate data entry into the costing_data.xml file for Chemstations.

# Running costing_data_script.py - Arguments
The program has a few arguments/parameters.

`--filename`, `-fn`
* Name of the temporary update file to be saved. Default can be changed in file as well (i.e. `costing_data_updated.xml`)

# Running browser automation
The file, `costing_data_script.py` automates it so that the browser will automatically open to download the newest CEPCI data file. Chrome is required in order to run this feature. There are certain wait times, but the browser should not stall at one web page without any action for more than ten seconds.
If the user does not want to run the browser automation, they should first download the data file they would like to use and place it into the `data/` folder in the main repository. They should then run `costing_data_script_no_browser.py`, which will pull the most recent data file in the `data/` repository if another file is not specified (through the `--datafile` or `-df` parameter, which asks for the path of the raw data file to be used).

# Saved files
The updated data file can be saved directly into `costing_data.xml` or a file of a chosen name, specified with the `--filename` or `-fn` arguments. After the preview file is shown to the user, the user is given the option to either save the file into `costing_data.xml` or into the specified filename. 
