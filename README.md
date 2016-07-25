# Spyglass
Nationstates Sheet-generator for easily finding approximate update times!

Running:

Windows:
Open up the Spyglass folder in your extracted .zip and run the executable there. It'll prompt you for a nation name, so input that, and wait. Within a couple minutes, a Spyglass spreadsheet will appear in that folder. And with that, you're done!

Mac/Linux:
Here, things are a bit more involved. First off, you'll need python 2.7, so if you don't already have it, go download it.
Next off, you'll need one of the required dependencies. To do that, head into your terminal/command line and enter the following:
sudo easy_install pip
sudo pip install openpyxl
A couple warnings might crop up, but those shouldn't cause any actual problems as far as running the program goes
Now, navigate to the Source subfolder of Spyglass, and enter the following:
python Sheet.py
As above, you'll be prompted for your nation name, and then things will chug along for a couple minutes, before creating the final sheet.

Caveat: Times given by Spyglass are only approximate, since update time varies from update to update due to nation creation/deletion and whatever else. Don't rely on given times in exclusivity; it won't work very well.

Calibration: Spyglass, at its present state, relies on manual input to determine the approximate lengths of major and minor update. Since this number is unlikely to stay the same, update times may decrease in accuracy over time. It is, however, possible to recalibrate by following these steps:
1: Find the length of major and minor updates. This can be accomplished by going to the activity page of a late-updating region (I like to use Spear Danes for this)
Look for messages like the following: [nation]'s influence has increased from X to Y. Pick up the timestamp there, for the last two updates. Major happens at around 12 AM, EST, and minor at around 12 PM.
2: Convert the time into seconds. This is fairly easy- 3600*hours + 60*minutes + seconds
3: In your folder, you'll find a text file labelled UpdTime. Open that up, and plug in your numbers; minor and major, respectively. Save and close, and it'll use those numbers the next time you run Spyglass.
