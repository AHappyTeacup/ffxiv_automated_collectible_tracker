# Final Fantasy XIV Automated Collectible Tracker
Inspired by my FC's use of Google Sheets to track mount farming progress amongst the members, 
I asked myself "But what if this updated itself?".

This is the beginning of an answer to that question.

FFXIV ACT uses Lodestone, and Google Sheets to automatically track and compare Collectibles progress between members of an FC, a raid static, or just a group of friends. \
You just need a list of character names, and your API authentication key for Google Sheets.

## Usage
Eventually, this will have a GitHub action to deliver an executable.\
For now you will need to be comfortable with using python.

The entry point for the python code is `act.py`.\
Running `act.py --help` should give an overview of the available commands and expected parameters.

**get-fc-members** \
The subcommand `get-fc-members` will populate a yaml file with a list of a specified Free Comapny's members. \
The command will require the World name, the Free Company name (in quotes if it contains spaces), and optionally the name of the file to write to (characters.yaml by default). \
For example, filling the file characters.yaml for my FC, The BLUs Brothers would require the command:
```bash
act.py get-fc-members Phantom "The BLUs Brothers"
```

**fill-sheet-data**\
The subcommand `fill-sheet-data` does the main work of this project. \
For each Google Sheet in the config.yaml file, the sheet be stripped down to "Sheet1", then the Sheets will be built with the headings specified, then the data for each character in the characters file will be filled in.

This takes three arguments.
 - --credentials-file\
 This is the filepath to the Google Sheets API JSON key. If left blank, it will look for `.credentials.json` in the current directory.
 - --sheet-config-file
 This is the filepath to the config.yaml file that details the sheet config. If left blank, it will look for `config.yaml` in the current directory.
 - --characters-file
 This is the filepath to the characters.yaml file that lists the FFXIV characters to be catalogued in the sheet. If left blank, it will look for `characters.yaml` in the current directory.

```bash
act.py fill-sheet-data --credentials-file .credentials.json --sheet-config-file config.yaml --characters-file characters.yaml
```

## SETUP

### Python
Install python on your machine. \
This project was written in python 3.7, and the constantly evolving nature of async in python may mean you need to stick close to that.\
Installing python on your machine should automatically install pip.\
Using pip, or otherwise, install [poetry](https://python-poetry.org/) (`pip install poetry`).\
In the directory where you've downloaded this repo, run `poetry install` in your terminal to install the requirements listed in poetry.lock.

### Google Sheets API Key
This assumes you have a google account since you're using Google Spreadsheets. \
Just follow [this video](https://www.youtube.com/watch?v=ddf5Z0aQPzY&t=63s) from 1:03 to 3:40.

### config.yaml
There is an example configuration file in the repo. \
It has the basic configuration for a Mounts spreadsheet (as of patch 6.3). \
The most important change to make to this is to add the ID of a spreadsheet in the example location prompted.

Edit the colours as you see fit, and edit any other sheet details as desired. \
You can add more Spreadsheets, Sheets, or columns as you desire.

### characters.yaml
There is an example configuration file in the repo. \
You can extend this yourself for a raid group, or create a file from your FC using this tool.


# Known Bugs
- Characters who have unlocked no mounts will produce empty boxes for mounts instead of 'N'.