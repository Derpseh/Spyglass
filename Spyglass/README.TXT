Spyglass: Generate NationStates region update timesheets.
Developed by Panzer Vier, with additions by Khronion

USING SPYGLASS

To run Spyglass, double-click "Spyglass" or "Spyglass.exe" and input your nation name when prompted. Spyglass will not
run without a valid nation. Spyglass will then generate an Excel spreadsheet with a list of regions and their estimated
update times. Spyglass requires an internet connection to run, so it can download the latest daily dump and query the
API to gather all the data necessary for generating a spreadsheet of region update times.

READING A SPREADSHEET

Regions are color-coded and marked in columns A and B based on their value as targets. These markers can be used with
Excel's built in filter function to quickly identify targets.

GREEN -- Region is founderless and passwordless
YELLOW -- Region has an executive WA delegate and is passwordless
RED -- Region has non-executive WA delegate and is passwordless

~ -- Region can be tagged
* -- Region has non-executive WA delegate

Additionally, regions are color-coded in column H based on their endorsement count.

RED -- Region has zero WAD endorsements