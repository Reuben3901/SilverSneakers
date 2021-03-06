# SilverSneakers
### Version 1.1.2

This program is a tool to automatically process SilverSneaker member visit reports generated by both the gym software MindBody and the front door key fob entry system KEYS.

The automation of this task saved 20+ hours of work each month for employees and the generated statistics revealed a missed opportunity of more than $1200 in monthly revenue.

### Sample reports and member data included to try the program for yourself:
* FrontDoorKeysReport.txt
* MindBodyReport.xlsx
* CompleteMemberList.xlsx

## Installation

### Requirements:
* Python 3
* Virtualenv (optional)
* OpenPyXL

1. Open a CMD Command Promt
Press `⊞Win`+`s` in windows and search for 'cmd'
2. If you don't have venv, you will need to install that first <br />
`pip install virtualenv`
3. Navigate to a folder you would like to download this project into <br />
`cd <folder location>` <br />
  example: `C:\Windows\System32>cd C:Users\Reuben\Desktop\SilverSneakers`
4. Download this project into the folder <br />
`git clone https://github.com/Reuben3901/SilverSneakers.git`
5. Create the virtual environment, named virtenv <br />
`python -m venv virtenv`
6. Activate the virtual environment <br />
`virtenv\Scripts\activate.bat` <br />
  You will now see: `(virtenv)` in front: `(virtenv) C:Users\Reuben\Desktop\SilverSneakers>`
7. Always good to upgrade Python's package handler pip <br />
`python -m pip install --upgrade pip`
8. Install the required modules to run this program <br />
`python -m pip install -r requirements.txt`
9. Run the program <br />
`python SilverSneakers.py`
10. To deactivate the virtual environment <br />
`deactivate`

### - alternatively - 
Simply add the required modules to your global Python installation <br />
`pip install openpyxl` <br />
Then follow steps #3, 4, 9

## ChangeLog

### Version 1.1.2
* Updated instructions: Removed wrongful mention of Pygame

### Version 1.1.1
* Removed usage of Send2Trash module
* No longer creates files: SilverSneakersReportsCombined.xlsx, SilverSneakersReportClean.xlsx

### Version 1.1.0
* Added functionality to support Silver&Fit and OptumFitness

## License & Copyright
© Reuben W. Young