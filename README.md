# Spyglass
Spyglass generates Excel Spreadsheets that can be used to predict the update time of a NationStates region, usually to coordinate an invasion, liberation, or other military gameplay action.

* **Estimated update times**: View estimated major and minor update times for all regions.
* **R/D metagame prioritization**: Identify regions based the presence/absence of a founder, WA delegate, and password.
* **WFE and Embassy preview**: Preview world factbook entries and regional embassies to quickly determine if a region has been impacted by the R/D metagame.

To learn more about what information Spyglass provides, [view the documentation](docs/interpreting_sheets.md).

## Usage
Windows and macOS users should use the pre-packaged release files for their operating system. These versions contain all the files needed to run Spyglass. Linux users should see the [instructions for setting up a development environment](docs/dev_environment.md).

> **Important note**: Microsoft Defender Smartscreen may block execution of programs it doesn't recognize, like Spyglass. They may also block the first attempt to connect to the internet. Browsers, including Chrome, may also flag the download of unknown executables as risky. You may bypass these warnings by clicking "Run Anyway" on the Smartscreen warning dialog.

After downloading the latest version of Spyglass, run the program and follow the on-screen prompts. Spyglass will then create an Excel spreadsheet in the same directory where it is located.

Spyglass currently runs in a command prompt window, but a graphical user interface is planned for the future.

*Disclaimer: Although Spyglass maintainers make every effort to comply with NS API rules, players are ultimately responsible for the behavior of any scripts, tools, or other programs that they use. Players are encouraged to report any issues or unexpected behavior in this forum thread, or by filing an issue on GitHub.*

### Runtime Flags

Some runtime flags are available for advanced users.

```commandline
$ python spyglass.py -h

usage: Spyglass [-h] [-n NATION] [-o OUTFILE] [-m] [-s] [--minor MINOR] [--major MAJOR] [-d] [-p PATH]

Spyglass 3.0.0: Generate NationStates region update timesheets. Developed by Panzer Vier, with additions by Khronion, Zizou, and Aav

optional arguments:
  -h, --help            show this help message and exit
  -n NATION, --nation NATION
                        Nation to use for authentication. Use underscores instead of spaces.
  -o OUTFILE, --outfile OUTFILE
                        Name of the output file. Defaults to 'spyglass.xlsx'.
  -m, --minimize        Generate a minimized sheet without WFEs and embassies
  -s, --suppress        Suppresses the creation of a debug log file
  --minor MINOR         The length of minor update in seconds.
  --major MAJOR         The length of major update in seconds.
  -d, --dump            Do not download the latest data dump. Use the one in the current directory.
  -p PATH, --path PATH  Path to the data dump. Defaults to 'regions.xml.gz'.


If run without arguments, Spyglass runs in interactive mode and outputs to its
working directory.
```

## Licensing
Spyglass version 2.0 and onward are licensed under the terms of the [GNU General Public License version 3.0](https://choosealicense.com/licenses/gpl-3.0/).

## Credits

Spyglass was originally developed by Devi and is now maintained by [Devi](https://github.com/Derpseh/), [Khronion](https://github.com/khronion), [Zizou](https://github.com/Aptenodyte), and [Aav](https://github.com/AavHRF). The icon was contributed by [Vleerian](https://github.com/vleerian/).
