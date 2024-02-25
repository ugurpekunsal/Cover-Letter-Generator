import win32gui
import win32.lib.win32con as win32con

from pyforms.utils.settings_manager import conf;
import settings
conf+=settings

import pyforms
from   pyforms          import BaseWidget
from   pyforms.controls import *

from fpdf import FPDF

the_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(the_program_to_hide , win32con.SW_HIDE)

class CoverLetterGenerator(BaseWidget):

    def __init__(self):
        super(CoverLetterGenerator,self).__init__('Cover Letter Generator')

        self._flags = {'company':False, 'position':False, 'portal':False}

        #Define progress bar
        self._progress          = ControlProgress()

        #Definition of the forms fields
        self._companyName       = ControlText('Company name:')
        self._positionName      = ControlText('Position name:  ')
        self._canadaJobBank     = ControlCheckBox('Canada job bank', 'True')
        self._indeed            = ControlCheckBox('Indeed', 'False')
        self._linkedin          = ControlCheckBox('Linkedin', 'False')
        self._ziprecruiter      = ControlCheckBox('Ziprecruiter', 'False')
        self._button            = ControlButton('Generate')
        
        #Define actions
        self._companyName.changed_event = self.__companyAction
        self._positionName.changed_event = self.__positionAction
        self._canadaJobBank.changed_event = self.__canadaJobBankAction
        self._indeed.changed_event = self.__indeedAction
        self._linkedin.changed_event = self.__linkedinAction
        self._ziprecruiter.changed_event = self.__ziprecruiterAction

        #Define the button action
        self._button.value = self.__buttonAction

        #Define label
        self._portalQuestion    = ControlLabel("Please select the job portal:")
        self._statusMessage    = ControlLabel()

        self.formset = [
                        ('_progress'),
                        ('_companyName'),
                        ('_positionName'),
                        ('_portalQuestion'),
                        ('_canadaJobBank'),
                        ('_indeed'),
                        ('_linkedin'),
                        ('_ziprecruiter'),
                        ' ',
                        (' ', '_button', ' '),
                        (' ', '_statusMessage', ' ')
                        ]
    
    def __companyAction(self):
        textActionHandler(self, 'company', self._companyName.value)

    def __positionAction(self):
        textActionHandler(self, 'position', self._positionName.value)
        
    def __canadaJobBankAction(self):
        portalActionHandler(self, 'canadaJobBank')

    def __indeedAction(self):
        portalActionHandler(self, 'indeed')

    def __linkedinAction(self):
        portalActionHandler(self, 'linkedin')

    def __ziprecruiterAction(self):
        portalActionHandler(self, 'ziprecruiter')

    def __buttonAction(self):
        companyName = self._companyName.value
        positionName = self._positionName.value
        if (self._canadaJobBank.value == True):
            jobPortalName = "the Canada job bank"
        elif (self._indeed.value == True):
            jobPortalName = "Indeed"
        elif (self._linkedin.value == True):
            jobPortalName = "Linkedin"
        elif (self._ziprecruiter.value == True):
            jobPortalName = "Ziprecruiter"
        else:
            self._statusMessage.value = "No portal selected!"
            return
        generateCoverLetter(companyName, positionName, jobPortalName)        
        self._statusMessage.value = "Generated succesfully!"

def textActionHandler(self, source, value):
    self._statusMessage.value = ""
    if (len(value) == 0):
        self._progress.__sub__(33)
        self._flags[source] = False
        return
    if (self._flags[source] == True):
        return
    self._progress.__add__(33)
    self._flags[source] = True

def portalActionHandler(self, source):
    self._statusMessage.value = ""
    if (self._flags['portal'] == True):
        return
    checkBoxes = {'canadaJobBank': self._canadaJobBank, 'indeed': self._indeed, 'linkedin': self._linkedin, 'ziprecruiter': self._ziprecruiter}
    self._flags['portal'] = True
    if (checkBoxes[source].value == False):
        self._progress.__sub__(34)
        self._flags['portal'] = False
        return
    del checkBoxes[source]
    if (all(value.value == False for value in checkBoxes.values())):
        print('first time checking')
        self._progress.__add__(34)
        self._flags['portal'] = False
        return
    for key, value in checkBoxes.items():
        value.value = False
    self._flags['portal'] = False

def generateCoverLetter(companyName, positionName, jobPortalName):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size = 12)

    text = open("coverletter.txt", "r")

    for x in text:
        x = x.replace("companyName", companyName)
        x = x.replace("positionName", positionName)
        x = x.replace("jobPortalName", jobPortalName)
        pdf.multi_cell(0, 8, txt = x)
        pdf.cell(0, 4, ln = 1)
    pdf.output("C:/Users/ugurp/Downloads/Ugur (Ian) Pekunsal - Cover Letter.pdf")

if __name__ == "__main__":   pyforms.start_app( CoverLetterGenerator )
    
