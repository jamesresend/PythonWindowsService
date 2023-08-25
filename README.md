# PythonWindowsService

I've implemented this code using Pycharm with Python 3.11. 

This service will write to a log file every 5 seconds, located [ActualServicePath]\Logs\PythonWindowsService.log

Install the following dependencies 
>pip install pywin32
>
>pip install pyinstaller

Compile your code into a executable (this will create a \dist folder in you project folder with a PythonWindowsService.exe)
>pyinstaller --onefile --hidden-import=win32timezone --clean PythonWindowsService.py

In a new command line (note to change you binPath to your actual path)
>sc create "Python Windows Service" binPath="Z:\PythonWindowsService\dist\PythonWindowsService.exe"

To change the description of you service
>sc description "Python Windows Service" "Example Python Service"

To start the Windows Service
>sc start "Python Windows Service"

To stop the Windows Service
>sc stop "Python Windows Service"

To delete the Windows Service
>sc delete "Python Windows Service"
