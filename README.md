# Spyglass

*Nationstates Sheet-generator for easily finding approximate update times*

[Latest release](https://github.com/Aptenodyte/Spyglass/releases)

## Running Spyglass

To run Spyglass from source, you will need a working Python 3.6+ installation with Openpyxl, Requests, BeautifulSoup4, and LXML. You can obtain these packages through pip:  
```
pip install openpyxl requests beautifulsoup4 lxml
```
Execute `Spyglass.py` with your Python 3.6+ interpreter.

```
python Spyglass.py
```

Follow the subsequent prompts to generate a sheet. You must provide a valid nation name in order to use Spyglass.

### Recommended execution method
To run Spyglass via EXE, download the binary from the releases page. You will not need to install any dependencies. It will launch in a GUI mode.

As a reminder, times given by Spyglass are only approximate, since update time varies from update to update due to nation creation/deletion and server-side variance. These times should be used to place triggers which can be used in turn to determine a given region's true update time.

## Using the Spreadsheet

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

## Supported Terminal Arguments

If you need Spyglass to run unattended (e.g. through a scheduled task or Cron job), you can use the following arguments
to control its behavior.

- `-n NATION` Specify Nation to identify user by. In order to comply with NationStates API rules, this must be the user's nation. Use underscores instead of spaces.
- `-o OUTFILE` Specify path to write the generated timesheet in XLSX format to.
- `-s` Suppress creating a debug log file. Log files are written to the current working directory.
- `-l PATH` Write debug log to specified path.
- `-m` Suppress the generation of embassies. This makes it easier to search a sheet for a specific nation by name.

If run without arguments, Spyglass runs in interactive mode and outputs to the working directory.
