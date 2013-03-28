GetOutlook
==========

GetOutlook is a python script which downloads all mail from an Outlook
account.

It is initially based on the work of GetLive
(http://sourceforge.net/projects/getlive/). The interface almost backwards
compatible to support drop-in replacement.

The mail will be downloaded in mbox format, where each foldername is represented as one file.
Use these files to load in your favorite email program

## Installation

Use a Python 2.7 installation and install 2 extra packages (mechanize and configobj) with

    pip install mechanize
    pip install configobj
    
Download GetOutlook.py with the sample config.
Modify the config file with your desired options (see below) and start fetching.

## Usage

Usage: GetOutlook.py [options]

    Options:
        -h, --help             show this help message and exit
        --config-file=CONFIGFILE
                               Configuration file (mandatory)
        --verbosity=VERBOSITY
                               Verbosity of messages (1/2/10/100)

## Config file

Config file:

    Username = user.name                (your username)
    Domain = outlook.com                (the domain of the email address)
    Password = your_secret_password     (secret)
    Downloaded = download.txt           (for old Getlive accounts, use this once)
    DestinationDir = /path/to/mail      (directory to store email in mbox format, folder name as filename)
    StatusFile = status.txt             (status of of downloaded message ids and a bit more)
    BreakOnAlreadyDownloaded = 40       (break pagescan when seen number of existing messages, 0 = always scan)


Every option in mandatory, except Downloaded and BreakOnAlreadyDownloaded

## Feedback

Please leave feedback at github page(https://github.com/ParseGuy/GetOutlook/)

Parse Guy
