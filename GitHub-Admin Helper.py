import sys
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QSplashScreen, QApplication


app = QApplication(sys.argv)
splash = QSplashScreen(QPixmap("O:pathway to png for splashloading screen"))
splash.show()

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import Admin_Helper_Single_AK_Checks
import Admin_Helper_Iterate_Admin_AKsv2
import Admin_Helper_PKsv2
import Admin_Helper_Iterate_Admin_FAs
import Admin_Helper_Iterate_Admin_POs
import autoconformityV2
import autopoa
import autona
import tolerance_evaluationv22
import cropnresize_insert_to_excel
import Admin_Helper_fix_supplier
import Admin_Helper_conformitydoubles
import Admin_helper_inspections
import Admin_Helper_Inspection_Performancev2
import Admin_Interactive_Parts_Diagram
import Admin_Helper_certtodo
import Admin_Helper_QP_Check
import easygui
import subprocess

#a lot of these programs use image grabs to find where to click on GUI's, i need
#to set imagegrab to False or True for whichever program is activates as some support
#image grabs across all screen and some only setup to support the main  image grab
from PIL import ImageGrab
from functools import partial

# The following folders/files necessary to run programs that are in my Quality folder,
# =============================================================================================
# "Auto AK Image",
# "Inspections Raw Data" #file in "Inspections" folder,
# numerous files within "Interactive_Parts_diagram" folder,
# "Tesseract OCR" folder,
# and the "poppler" folder within the "Insert_images_excel" folder.
# ===============================================================================================
# ALl these folders within my folder in quality drive needed #to run programs for this suite of programs

#main GUI
class Actions(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.setGeometry(400, 200, 350, 75)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setWindowTitle("Admin Helper Suite")

        self.menubar = QMenuBar()
        self.setMenuBar(self.menubar)

        actionFile = self.menubar.addMenu("File")

        check_issue_menu = actionFile.addMenu("Admin Incoming")
        check_issue_menu.addAction("Iterate AK Admin list", self.admin_ak_checks)
        check_issue_menu.addAction("iterate PK Admin list", self.admin_pk_checks)
        check_issue_menu.addAction("iterate FAs Admin list", self.admin_fa_checks)
        check_issue_menu.addAction("iterate POs Admin list", self.admin_po_checks)
        actionFile.addAction("AK Check Issued", self.single_checks)
        actionFile.addAction("Brian's FAIs to Review Finder", self.fai_review)
        check_conformity_menu = actionFile.addMenu("Conformity")
        check_conformity_menu.addAction("Complete Simple Items", self.conformity_simple)
        check_conformity_menu.addAction("Complete Double PK Lines", self.conformity_double_lines)
        check_conformity_menu.addAction("Remove NA Items", self.conformity_NA)
        check_conformity_menu.addAction("Remove POA Items", self.conformity_POA)
        actionFile.addAction("Copy Inspects to Folder", self.inspects)
        actionFile.addAction("Fix PO/Supplier Admin", self.po_supplier)
        actionFile.addAction("Insert Drawing to FAI", self.drawing)
        actionFile.addAction("Search Cert-to-do", self.cert)
        actionFile.addAction("Tolerance FAI Checker", self.tolerance)
        actionFile.addAction("QP Waiting Check", self.qp_waiting)
        miscellaneous = actionFile.addMenu("Misc")
        miscellaneous.addAction("Inspection Performance", self.performance)
        miscellaneous.addAction("Interactive Parts Diagram", self.interactive)
        actionFile.addSeparator()
        actionFile.addAction("Quit", self.close_window)

        myFont = QFont()
        myFont.setBold(True)
        myFont.setPointSize(10)

        self.about = QLabel("Pick an activity in file drop down", self)
        self.about.setFont(QFont('Times'))
        self.about.setFont(myFont)
        self.about.adjustSize()
        self.about.move(70, 35)
        self.show()


    def single_checks(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.w = Admin_Helper_Single_AK_Checks.Actions()
            self.w.show()
        except:
            easygui.msgbox(msg="Error 1337; try again and check instructions!", title="ERROR")

    def admin_ak_checks(self):
        try:
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
            self.admin_ak = Admin_Helper_Iterate_Admin_AKsv2.Actions()
            self.admin_ak.show()
        except:
            easygui.msgbox(msg="Error 807; try again and check instructions!", title="ERROR")

    def admin_fa_checks(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.admin_fa = Admin_Helper_Iterate_Admin_FAs.Actions()
            self.admin_fa.show()
        except:
            easygui.msgbox(msg="Error 8337; try again and check instructions!", title="ERROR")

    def admin_po_checks(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.admin_po = Admin_Helper_Iterate_Admin_POs.Actions()
            self.admin_po.show()
        except:
            easygui.msgbox(msg="Error 31337; try again and check instructions!", title="ERROR")

    def admin_pk_checks(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
            self.admin_pk = Admin_Helper_PKsv2.Actions()
            self.admin_pk.show()
        except:
            easygui.msgbox(msg="Error 8008; try again and check instructions!", title="ERROR")

    def conformity_simple(self):
        try:
            #see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.conformity_single = autoconformityV2.Actions()
            self.conformity_single.show()
        except:
            easygui.msgbox(msg="Error 101; try again and check instructions!", title="ERROR")

    def conformity_double_lines(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.conformity_double = Admin_Helper_conformitydoubles.Actions()
            self.conformity_double.show()
        except:
            easygui.msgbox(msg="Error 10101; try again and check instructions!", title="ERROR")

    def conformity_NA(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.conformity_nas = autona.Actions()
            self.conformity_nas.show()
        except:
            easygui.msgbox(msg="Error 707; try again and check instructions!", title="ERROR")

    def conformity_POA(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.conformity_poas = autopoa.Actions()
            self.conformity_poas.show()
        except:
            easygui.msgbox(msg="Error 07734; try again and check instructions!", title="ERROR")

    def drawing(self):
        try:
            self.insertdrawing = cropnresize_insert_to_excel.Actions()
            self.insertdrawing.show()
        except:
            easygui.msgbox(msg="Error 200; try again and check instructions!", title="ERROR")

    def tolerance(self):
        try:
            self.check_tolerance = tolerance_evaluationv22.Actions()
            self.check_tolerance.show()
        except:
            easygui.msgbox(msg="Error 8008135; try again and check instructions!", title="ERROR")

    def po_supplier(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.fai_suppliers = Admin_Helper_fix_supplier.Actions()
            self.fai_suppliers.show()
        except:
            easygui.msgbox(msg="Error 8007; try again and check instructions!", title="ERROR")

    def inspects(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.copy_inspects = Admin_helper_inspections.Actions()
            self.copy_inspects.show()
        except:
            easygui.msgbox(msg="Error 37331; try again and check instructions!", title="ERROR")

    def fai_review(self):
        try:
            file_location = "O:\\Quality\\Brian J\\Automation\\Released\\Process MDL list.exe"
            subprocess.call(file_location)
        except:
            easygui.msgbox(msg="Error 80; Couldn't find Program!", title="ERROR")

    def performance(self):
        try:
            self.inspect_performance = Admin_Helper_Inspection_Performancev2.Actions()
            self.inspect_performance.show()
        except:
            easygui.msgbox(msg="Error 37331; try again and check instructions!", title="ERROR")

    def interactive(self):
        try:
            self.interactive_diagram = Admin_Interactive_Parts_Diagram.Actions()
            self.interactive_diagram.show()
        except:
            easygui.msgbox(msg="Error DESTRUCTO; try again and check instructions!", title="ERROR")

    def cert(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.cert_to_do = Admin_Helper_certtodo.Actions()
            self.cert_to_do.show()
        except:
            easygui.msgbox(msg="Error DRAGONBALL; Try again and check instructions!", title="ERROR")

    def qp_waiting(self):
        try:
            # see note at top where image grab module is called for why this line is in here
            ImageGrab.grab = partial(ImageGrab.grab, all_screens=False)
            self.qp_check = Admin_Helper_QP_Check.Actions()
            self.qp_check.show()
        except:
            easygui.msgbox(msg="Error 666; WHAT ERROR?!", title="ERROR")

    def close_window(self):
        self.close()

if __name__ == "__main__":
    #app = QApplication(sys.argv)
    window = Actions()
    splash.finish(window)
    sys.exit(app.exec_())
