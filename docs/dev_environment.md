# Setting up Python for Spyglass

It is recommended that users set up a virtual environment using [venv](https://virtualenv.pypa.io/en/latest/), [pip-env](https://pipenv.pypa.io/en/latest/), [conda](https://docs.conda.io/en/latest/), or a similar tool when using Spyglass.

Spyglass requires Python 3. Users may automatically install dependencies by running the following command:

```commandline
$ pip install -r requirements.txt
```

If `pip` fails, make sure that you are using the command associated with Python 3.x. If `pip` is not installed on your system, see the[ documentation for installing pip](https://pip.pypa.io/en/stable/installation/).

Once packages are installed, Spyglass can be executed from the terminal directly:

```commandline
$ python Spyglass-cli.py
```

## Runtime Flags

Some runtime flags are available for advanced users.

```commandline
$ python Spyglass-cli.py -h

Spyglass 2.0: Generate NationStates region update timesheets.

Developed by Panzer Vier, with additions by Khronion, Zizou, and Aav

usage: .\Spyglass-cli.py [-h] [-n NATION] [-o OUTFILE] [-s | -l PATH]

Optional arguments:
-h           Show this help message and exit.
-n NATION    Specify Nation to identify user by. In order to comply with
NationStates API rules, this must be the user's nation. Use
underscores instead of spaces.
-o OUTFILE   File to output the generated timesheet in XLSX format to.
-s           Suppress creating a debug log file. Log files are written to
the current working directory.
-l PATH      Write debug log to specified path.
-m           Generate a minimized sheet without WFEs and embassies

If run without arguments, Spyglass runs in interactive mode and outputs to its
working directory.
```

As of 2.0, a user must still manually confirm if they wish to re-use an existing `regions.xml.gz` file with the interactive prompt even if they are using runtime flags.
