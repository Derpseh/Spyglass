# Interpreting a Spyglass Sheet

| Column | Title              | Description                                                                                                     |
|--------|--------------------|-----------------------------------------------------------------------------------------------------------------|
| A      | Region Name        | Region name in NationStates                                                                                     |
| B      | Region Link        | Clickable hyperlink to region in NationStates                                                                   |
| C      | # of Nations       | Region population                                                                                               |
| D      | Tot Nations        | Total population for all regions up to and including this region<br/>_Used to calculate estimated update times_ |
| E      | Minor Upd. (est)   | Predicted time of region's minor update                                                                         |  
| F      | Major Upd. (true)  | Most recently observed region major update time<br/>_Only shown if update lengths are not manually specified_   |
| F      | Major Upd. (est)   | Predicted time of region's major update <br/>_Only shown if update lengths are manually specified_              | 
| G      | Del. Votes         | Voting power for the region's world assembly delegate (one less than the delegate's endorsements)               |
| H      | Del. Endos         | Number of endorsements on the region's world assembly delegate                                                  |
| I      | Embassies          | Comma-delimited list of regional embassies                                                                      |
| J      | WFE                | Preview of region world factbook entry                                                                          |


## Color-coding
To help users identify regions of interest in the R/D metagame, regions in Spyglass sheets are color-coded and marked in columns A and B based on if they have a founder, if they have a World Assembly delegate, and if the delegate has executive powers. These markers can be used with Excel's built-in filter function to quickly identify regions of interest.

Columns A and B are color-coded and suffixed based on a region's relevance to the military gameplay metagame, based on if they have a founder, a password, and an executive world assembly delegate.

| Indicator      | Description                                                                                                                                                |
|----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Green (color)  | Region is founderless and has no password.                                                                                                                 |
| Yellow (color) | Region has a founder and an executive world assembly delegate, but has no password                                                                         |  
| Red (color)    | Region has a non-executive world assembly delegate and has no password                                                                                     |
| `~` (suffix)   | Region is potentially subject to military gameplay actions.<br/>_Has an executive world assembly delegate and no password. May or may not have a founder._ |
| `*` (suffix)   | Region is not subject to military gameplay actions.<br/>_Has a non-executive world assembly delegate._                                                     |

Additionally, regions with zero world assembly endorsements are color-coded red in column H.