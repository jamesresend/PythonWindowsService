# pip install pywin32
''' If PyCharm highlights win32serviceutil, win32service or win32event with error after pywin32 install, just restart PyCharm '''

# pip install pyinstaller

'''Compile your code into a executable (this will create a \dist folder in you project folder with a PythonWindowsService.exe)'''
# pyinstaller --onefile --hidden-import=win32timezone --clean PythonWindowsService.py

'''Open a new command line for the code bellow'''
# sc create "Python Windows Service" binPath="Z:\PythonWindowsService\dist\PythonWindowsService.exe"

'''To change the description of you service'''
# sc description "Python Windows Service" "Example Python Service"

'''To start the Windows Service'''
# sc start "Python Windows Service"

'''To stop the Windows Service'''
# sc stop "Python Windows Service"

'''To delete the Windows Service'''
# sc delete "Python Windows Service"

import os
import sys
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import logging


class PythonWindowsService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PythonWindowsService"
    _svc_display_name_ = "Python Windows Service"

    def __init__(self, args):
        super().__init__(args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.logger = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        self.logger = self.setup_logger()
        self.logger.info("Service is starting...")

        while self.is_alive:
            self.logger.info("Service is running...")
            win32event.WaitForSingleObject(self.hWaitStop, 5000)  # Wait for 5 seconds or until stop signal

        self.logger.info("Service is stopping...")

    def setup_logger(self):
        log_dir = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'Logs')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f'PythonWindowsService.log')

        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] - %(message)s',
        )

        return logging.getLogger()


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PythonWindowsService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(PythonWindowsService)
