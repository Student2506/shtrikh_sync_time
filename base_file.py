import sys
import time
import logging
from datetime import datetime as dt

import pythoncom
import servicemanager
import win32service
import win32serviceutil
import win32com.client

FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logger = logging.getLogger(__name__)


class MyService:
    def stop(self):
        self.running = False

    def run(self):
        self.running = True
        while self.running:
            logging.basicConfig(level=logging.DEBUG, format=FORMAT)
            drvfr = win32com.client.Dispatch(
                'AddIn.Drvfr', pythoncom.CoInitialize()
            )
            drvfr.GetCountLD()
            ld_numbers = []
            for i in range(drvfr.LDCount):
                drvfr.LDIndex = i
                drvfr.EnumLD()
                ld_numbers.append(drvfr.LDNumber)

            servicemanager.LogInfoMsg(
                f'Текущее число принтеров: {drvfr.LDCount}'
            )
            for i in ld_numbers:
                drvfr.LDNumber = i
                drvfr.SetActiveLD()
                drvfr.Connect2()
                servicemanager.LogInfoMsg(
                    f'Код соединения: {drvfr.ResultCode}'
                )
                if drvfr.ResultCode:
                    continue
                drvfr.Password = 30
                drvfr.GetECRStatus()
                if drvfr.ECRMode in (4, 7, 9):
                    drvfr.TimeStr = dt.now().strftime('%H:%M:%S')
                    drvfr.SetTime()
                    servicemanager.LogInfoMsg(
                        f'Код установки времени: {drvfr.ResultCode}'
                    )
                drvfr.Disconnect()

            time.sleep(86400)
            servicemanager.LogInfoMsg('Service Running...')


class MyServiceFramework(win32serviceutil.ServiceFramework):

    _svc_name_ = 'shtrikh_time_sync'
    _svc_display_name_ = 'Time Sync Shtrikh'

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.service_impl.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        self.service_impl = MyService()
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.service_impl.run()


def init():
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MyServiceFramework)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(MyServiceFramework)


if __name__ == '__main__':
    init()
