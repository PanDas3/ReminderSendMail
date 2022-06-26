# Project Name
* SendMail - version 1.1.37

Project Date: March 2021


## Additional infromation

- This project reminds me of when I should send in my request renewal permission to services.
- I wrote this program after my job, because it is an element of my self-development.
- If you find a bug you can write to me :)


## Table of Contents
* [General Info](#general-information)
* [Additional infromation](#additional-infromation)
* [General Information](#general-information)
* [Technologies Used](#technologies-used)
* [Usage](#usage)
* [Project Status](#project-status)
* [Contact](#contact)
<!-- * [License](#license) -->


## General Information
The program has:
- Getting params from configuration file
- Reading date from excel
- Match current date to date from excel
- Building frame in HTML (to content mail)
- Sending mail with table from excel
- Logged status


## Technologies Used
- Python - version 3.8
- Pandas - version 1.2.2
- ConfigParser - version 5.0.2
- Pretty_Html_table - version 0.9.dev0
- Logger - version 2.3.0
- PyInstaller - version 4.2

<!--
## Features
None


## Screenshots
![Example screenshot](./img/screenshot.png)


## Setup
What are the project requirements/dependencies? Where are they listed? A requirements.txt or a Pipfile.lock file perhaps? Where is it located?

Proceed to describe how to install / setup one's local environment / get started with the project.
-->

## Usage
If you would like to run from script:
1. Run main.py
2. Complete config.ini
3. Run main.py
4. Look at the log file

If u would like to run EXE (and then example put it into Task Scheduler, like me):

Generate EXE file - CMD ->
```batch
pip install pyinstaller --proxy http://user:pass@proxy.pl:3128
cd Building
pyinstaller ReminderSendMail.spec
```

1. Run SendMail.exe
2. Complete config.ini
3. Run again SendMail.exe
4. Look at the log file


## Project Status
Project is: _completed / discontinue_

<!-- _complete_ / _no longer being worked on_ (and why) -->

<!-- 
## Room for Improvement
No plans
-->

## Contact
Created by [@Majster](mailto:rachuna.mikolaj@gmail.com)