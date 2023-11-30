# bot_excel
a excel automation in python using openpyxl and pyautogui

## how it works?
**program.py** - a simple GUI to represent some program that you have in your company or something like that<br/>
**app.py** - reads the file tables.xlsx and insert each line into the program.py<br/>
**mouse.py** - this is a file that i used to get the coordinates for the pyautogui, if the program dont work really well in your machine use this file to get the correct coordinates and replace in the app.py

## how to use?
* git clone https://github.com/RaelzeraXD/bot_excel
* cd bot_excel
* pip3 install -r requirements.txt
* ./program.py
* ./app.py
after execute the app.py click on the window program.py 

## problems
if you are running in linux execute this command to avoid x11 problems
* xhost +local:$USER
