GetOutlook
==========

GetOutlook is a python script which downloads all mail from an Outlook
account.

It is initially based on the work of GetLive
(http://sourceforge.net/projects/getlive/). The interface almost backwards
compatible to support drop-in replacement.

The mail will be downloaded in mbox format, which is compatible with many
mail programs.

## Installation

Get the GetOutlook.py file and install the required python modules
(mechanize etc) with your favorite python package installer.

Modify the config file with your desired options (see below) and start
fetching.

## Usage

Usage: GetOutlook.py [options]

Options:
  -h, --help            show this help message and exit
  --config-file=CONFIGFILE
                        Configuration file (mandatory)
  --verbosity=VERBOSITY
                        Verbosity of messages (1/2/10/100)

## Config file

Config file:

Username = user.name                
Domain = outlook.com
Password = your_secret_password
Downloaded = download.txt           (for old Getlive accounts)
DestinationDir = /path/to/mail      (directory to store email)
StatusFile = status.txt             (new format downloaded messages)
BreakOnAlreadyDownloaded = 40       (break pagescan when seen number of
                                     existing messages, 0 = always scan)
                  
Every option in mandatory, except Downloaded and BreakOnAlreadyDownloaded

## Feedback

Please leave feedback at github page(https://github.com/ParseGuy/GetOutlook/)

Parse Guy
