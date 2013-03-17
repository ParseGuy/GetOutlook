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
        -h, --help             show this help message and exit
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

## Donations

GetOutlook was written in sole effort to get you the mail backup you need.
If you have the means, please support me with a donation. More donations mean
more time to support you, add features and feel happy about yourself!

### Donate via PayPal

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_SM.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=8U34AKDF35D8E)

## Feedback

Please leave feedback at github page(https://github.com/ParseGuy/GetOutlook/)

Parse Guy
