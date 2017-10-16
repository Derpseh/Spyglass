Nationstates Sheet-generator for easily finding approximate update times

[Latest release](https://github.com/khronion/Spyglass/releases)

# Running Spyglass

You will need a working Python 2.x installation with [OpenPyXL](https://openpyxl.readthedocs.io/en/default/) to run Spyglass. Simply run the following to use Spyglass:

```
python2 spyglass.py
```

You will be prompted for your nation name, and then Spyglass will take several minutes to generate a new sheet.

**Caveat:** Times given by Spyglass are only approximate, since update time varies from update to update due to nation creation/deletion and server-side variance. These times should be used to place triggers which can be used in turn to determine a given region's true update time.

### Using the Spreadsheet

Regions are color-coded and marked in columns A and B based on their value as targets. These markers can be used with
Excel's built in filter function to quickly identify targets.

- GREEN — Region is founderless and passwordless
- YELLOW — Region has an executive WA delegate and is passwordless
- RED — Region has non-executive WA delegate and is passwordless

- `~` — Region can be tagged
- `*` — Region has non-executive WA delegate

Additionally, regions are color-coded in column H based on their endorsement count.

- RED — Region has zero WAD endorsements

Spyglass now also lists a region's embassies and the first 140 characters from their WFE in columns I and J. These columns can be filtered to quickly identify regions that are not yet or currently tagged.

### Supported Terminal Arguments

If you need Spyglass to run unattended (e.g. through a scheduled task or Cron job), you can use the following arguments
to control its behavior.

- `-n NATION` Specify Nation to identify user by. In order to comply with NationStates API rules, this must be the user's nation. Use underscores instead of spaces.
- `-o OUTFILE` Specify path to write the generated timesheet in XLSX format to.
- `-s` Suppress creating a debug log file. Log files are written to the current working directory.
- `-l PATH` Write debug log to specified path

If run without arguments, Spyglass runs in interactive mode and outputs to the working directory.