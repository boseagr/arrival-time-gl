# arrival-time-gl

using python 3

external lib used : 
- eel
- winevt

create executable using pyinstaller

please make sure to create 'web' folder inside the script then empty html arrv12k.hmtl inside 'web' folder

python -m eel main.py web --exclude tkinter --exclude win32com --exclude PyQt5 --exclude numpy --exclude cryptography --exclude PIL --onefile --noconsole
