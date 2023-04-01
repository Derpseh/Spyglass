# Setting up Python for Spyglass

It is recommended that users set up a virtual environment using [venv](https://virtualenv.pypa.io/en/latest/), [pip-env](https://pipenv.pypa.io/en/latest/), [conda](https://docs.conda.io/en/latest/), or a similar tool when using Spyglass.

Spyglass requires Python 3.9 Users may automatically install dependencies by running the following command:

```commandline
$ pip install -r requirements.txt
```

If `pip` fails, make sure that you are using the command associated with Python 3.x. If `pip` is not installed on your system, see the[ documentation for installing pip](https://pip.pypa.io/en/stable/installation/).

Once packages are installed, Spyglass can be executed from the terminal directly:

```commandline
$ python spyglass.py
```

## Building Spyglass

OS-specific builds of Spyglass are generated using [pyInstaller](https://pyinstaller.readthedocs.io/en/stable/).

### macOS (Intel)

Install pyInstaller and UPX, then run:

```commandline
$ pyinstaller --clean spyglass.py -F -n Spyglass -c -i assets/Spyglass.icns
```

### Windows
```commandline
$ pyinstaller --clean spyglass.py -F -n Spyglass.exe -c -i Spyglass.ico
```