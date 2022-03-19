# Spyglass
Spyglass generates Excel Spreadsheets that can be used to predict the update time of a NationStates region, usually to coordinate an invasion, liberation, or other military gameplay action.

* **Estimated update times**: View estimated major and minor update times for all regions.
* **R/D metagame prioritization**: Identify regions based the presence/absence of a founder, WA delegate, and password.
* **WFE and Embassy preview**: Preview world factbook entries and regional embassies to quickly determine if a region has been impacted by the R/D metagame.

## Usage
Windows and macOS users should use the pre-packaged release files for their operating system. These versions contain all the files needed to run Spyglass.

**Important note**: Microsoft Defender Smartscreen may block execution of programs it doesn't recognize, like Spyglass. You may bypass these warnings by clicking "Run Anyway" on the Smartscreen warning dialog.

After downloading the latest version of Spyglass, run the program and follow the on-screen prompts. Spyglass will then create an Excel spreadsheet in the same directory where it is located.

Spyglass currently runs in a command prompt window, but a graphical user interface is planned for the future.

*Disclaimer: Although Spyglass maintainers make every effort to comply with NS API rules, players are ultimately responsible for the behavior of any scripts, tools, or other programs that they use. Players are encouraged to report any issues or unexpected behavior in this forum thread, or by filing an issue on GitHub.*

### Runtime Flags
Advanced users can use the following arguments to control Spyglass.

- `-n NATION` Specify Nation to identify user by. In order to comply with NationStates API rules, this must be the user's nation. Use underscores instead of spaces.
- `-o OUTFILE` Specify path to write the generated timesheet in XLSX format to.
- `-s` Suppress creating a debug log file. Log files are written to the current working directory.
- `-l PATH` Write debug log to specified path.
- `-m` Suppress the generation of embassies. This makes it easier to search a sheet for a specific nation by name.

### Interpreting a Spyglass Sheet

To help users identify regions of interest in the R/D metagame, regions in Spyglass sheets are color-coded and marked in columns A and B based on if they have a founder, if they have a World Assembly delegate, and if the delegate has executive powers. These markers can be used with Excel's built-in filter function to quickly identify regions of interest.

-   GREEN — Region is founderless and passwordless
-   YELLOW — Region has an executive WA delegate and is passwordless
-   RED — Region has non-executive WA delegate and is passwordless
-   `~` — Region can be tagged
-   `*` — Region has non-executive WA delegate

Additionally, regions with zero WAD endorsements are color-coded red in column H, which shows the region's endorsement count.

Spyglass also lists a region's embassies and the first 140 characters from their WFE in columns I and J. These columns can be used to identify regions that have been tagged by military gameplayers.

## Licensing
Spyglass version 2.0 and onward are licensed under the terms of the [GNU General Public License version 3.0](https://choosealicense.com/licenses/gpl-3.0/).

## Credits

Spyglass was originally developed by Devi and is now maintained by [Devi](https://github.com/Derpseh/), [Khronion](https://github.com/khronion), [Zizou](https://github.com/Aptenodyte), and [Aav](https://github.com/AavHRF). The icon was contributed by [Vleerian](https://github.com/vleerian/).
