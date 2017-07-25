# Spyglass
Nationstates Sheet-generator for easily finding approximate update times!

## Running Spyglass

### Windows

Open up the Spyglass folder in your extracted .zip and run the executable there. It'll prompt you for a nation name, 
so input that, and wait. Within a couple minutes, a Spyglass spreadsheet will appear in that folder. And with that, 
you're done!

### Mac/Linux

Here, things are a bit more involved. First off, you'll need python 2.7, so if you don't already have it, go download 
it. Next off, you'll need one of the required dependencies. To do that, head into your terminal/command line and enter 
the following:

```
sudo easy_install pip
sudo pip install openpyxl
```

A couple warnings might crop up, but those shouldn't cause any actual problems as far as running the program goes
Now, navigate to the Source subfolder of Spyglass, and enter the following:

```
python2 spyglass.py
```

As above, you'll be prompted for your nation name, and then things will chug along for a couple minutes, before creating 
the final sheet.

Caveat: Times given by Spyglass are only approximate, since update time varies from update to update due to nation 
creation/deletion and whatever else. Don't rely on given times in exclusivity; it won't work very well.

### Using the Spreadsheet

Regions are color-coded and marked in columns A and B based on their value as targets. These markers can be used with
Excel's built in filter function to quickly identify targets.

- GREEN -- Region is founderless and passwordless
- YELLOW -- Region has an executive WA delegate and is passwordless
- RED -- Region has non-executive WA delegate and is passwordless

- `~` -- Region can be tagged
- `*` -- Region has non-executive WA delegate

Additionally, regions are color-coded in column H based on their endorsement count.

- RED -- Region has zero WAD endorsements

### Calibration

Spyglass now automatically calibrates itself based on the most recent update.

### Supported Terminal Arguments

If you need Spyglass to run unattended (e.g. through a scheduled task or Cron job), you can use the following arguments
to control its behavior.

- `-n NATION` Specify Nation to identify user by. In order to comply with NationStates API rules, this must be the user's nation. Use underscores instead of spaces.
- `-o OUTFILE` Specify path to write the generated timesheet in XLSX format to.
- `-s` Suppress creating a debug log file. Log files are written to the current working directory.
- `-l PATH` Write debug log to specified path

If run without arguments, Spyglass runs in interactive mode and outputs to the working directory.
