
from PyQt5 import uic, QtWidgets
from PyQt5.QtCore import QThread
from EngHours import EngHrs
from OvtHours import OvtHrs
from ProjDbd import PrjDbd
from Misc import Misc
import sys
import ctypes

qtCreatorFile = "dashboard.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class Dashboard(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)  # Ui set-up
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        # Initialize objects for thread (for progress bar)
        self.thread = []
        self.worker = []
        def is_admin():     # Run script as admin
            try:
                return ctypes.windll.shell32.IsUserAnAdmin()
            except:
                return False

        if not is_admin():
            # Re-run the program with admin rights, 0 hide, 1 show (need 1 for .exe)
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit(0)

    def CovWTSSetup(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.WTSSetupCov)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def DavWTSSetup(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.WTSSetupDav)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def CovOvtSetup(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.SetupCov)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def DavOvtSetup(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.SetupDav)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def CovOvtHrs(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.OvtHrsCov)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def DavOvtHrs(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.OvtHrsDav)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def CovPopSap(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.PopSapCov)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def DavPopSap(self):
        # Thread objects
        self.thread = QThread()
        self.worker = OvtHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.PopSapDav)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxTS)
        self.worker.progress.connect(self.reportProgressTS)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def UpdEngHrs(self):
        # Thread objects
        self.thread = QThread()
        self.worker = EngHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.EngHrs)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxEH)
        self.worker.progress.connect(self.reportProgressEH)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def UpdEngShp(self):
        # Thread objects
        self.thread = QThread()
        self.worker = EngHrs()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.EngShp)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxEH)
        self.worker.progress.connect(self.reportProgressEH)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def EngH2ProjD(self):
        # Thread objects
        self.thread = QThread()
        self.worker = PrjDbd()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.EngH2ProjD)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMax)
        self.worker.progress.connect(self.reportProgress)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def UpdEco(self):
        # Thread objects
        self.thread = QThread()
        self.worker = PrjDbd()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.UpdateEco)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMax)
        self.worker.progress.connect(self.reportProgress)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )
        self.btnEng2Prj.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnEng2Prj.setEnabled(True)
        )
        self.btnEco.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnEco.setEnabled(True)
        )
        self.btnRelData.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnRelData.setEnabled(True)
        )

    def RelData2Wks(self):
        # Thread objects
        self.thread = QThread()
        self.worker = PrjDbd()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.ReleaseData)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMax)
        self.worker.progress.connect(self.reportProgress)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )
        self.btnEng2Prj.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnEng2Prj.setEnabled(True)
        )
        self.btnEco.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnEco.setEnabled(True)
        )
        self.btnRelData.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.btnRelData.setEnabled(True)
        )

    def CovRemSpc(self):
        # Thread objects
        self.thread = QThread()
        self.worker = Misc()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.RemSpcCov)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxM)
        self.worker.progress.connect(self.reportProgressM)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def DavRemSpc(self):
        # Thread objects
        self.thread = QThread()
        self.worker = Misc()
        # Move subthread to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.RemSpcDav)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.maximum.connect(self.getMaxM)
        self.worker.progress.connect(self.reportProgressM)
        # Start the thread
        self.thread.start()
        # Disable all buttons while subthread is running
        self.tabTimeSht.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabTimeSht.setEnabled(True)
        )
        self.tabEngHrs.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabEngHrs.setEnabled(True)
        )
        self.tabProjDash.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabProjDash.setEnabled(True)
        )
        self.tabMisc.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabMisc.setEnabled(True)
        )

    def WinReg(self):
        self.tabTimeSht.setEnabled(False)
        self.tabEngHrs.setEnabled(False)
        self.tabProjDash.setEnabled(False)
        self.tabMisc.setEnabled(False)
        Misc().WinRegIssue()
        self.tabTimeSht.setEnabled(True)
        self.tabEngHrs.setEnabled(True)
        self.tabProjDash.setEnabled(True)
        self.tabMisc.setEnabled(True)

    def ClsXl(self):
        self.tabTimeSht.setEnabled(False)
        self.tabEngHrs.setEnabled(False)
        self.tabProjDash.setEnabled(False)
        self.tabMisc.setEnabled(False)
        Misc().CloseAllXl()
        self.tabTimeSht.setEnabled(True)
        self.tabEngHrs.setEnabled(True)
        self.tabProjDash.setEnabled(True)
        self.tabMisc.setEnabled(True)

    def getMaxTS(self, value):          # Used to update maximum for loading bar
        self.progressBar_ts.setMaximum(value)

    def reportProgressTS(self, value):  # Used to update loading bar
        self.progressBar_ts.setValue(value)

    def getMaxEH(self, value):          # Used to update maximum for loading bar
        self.progressBar_eh.setMaximum(value)

    def reportProgressEH(self, value):  # Used to update loading bar
        self.progressBar_eh.setValue(value)

    def getMax(self, value):          # Used to update maximum for loading bar
        self.progressBar.setMaximum(value)

    def reportProgress(self, value):  # Used to update loading bar
        self.progressBar.setValue(value)

    def getMaxM(self, value):          # Used to update maximum for loading bar
        self.progressBar_m.setMaximum(value)

    def reportProgressM(self, value):  # Used to update loading bar
        self.progressBar_m.setValue(value)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Dashboard()
    window.show()
    sys.exit(app.exec_())
