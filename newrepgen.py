import sys
from PyQt4 import QtGui,QtCore
import pypyodbc
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch, landscape,cm
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.rl_config import defaultPageSize
import time
import locale
import csv
import pandas as pd




class Example(QtGui.QMainWindow):

    def errorMsg(self):
        reply = QtGui.QMessageBox.information(self, 'Error',
                "Unable to generate report : \n Please close existing report or Account number does not exist. ")
        

    def __init__(self):
        super(Example, self).__init__()

        self.initUI()


    def close_application(self): 
        sys.exit()


    def AboutMeRenan(self):
        
        reply = QtGui.QMessageBox.information(self, 'About New Report Gen',
                "New Report Gen Version 1.5 \n Copyright c2016. All rights reserved.\n This program is licensed from Mr. Renan Rivera (email:renzo031109@gmail.com) \n to : Avon Angeles Branch")


    def centerOnScreen (self):
        '''centerOnScreen()
         Centers the window on the screen.'''
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
    
    def initUI(self):

        #MENU
        extractAction = QtGui.QAction("&Quit", self)
        extractAction.setShortcut("Ctrl+Q")
        extractHelp = QtGui.QAction("&About", self)      

        extractAction.triggered.connect(self.close_application)
        extractHelp.triggered.connect(self.AboutMeRenan)

        self.statusBar()

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(extractAction)
        fileMenu = mainMenu.addMenu('&Help')
        fileMenu.addAction(extractHelp)
        

        pic = QtGui.QLabel(self)
        pic.setGeometry(180,0,500,380)
        pixmap = QtGui.QPixmap('REPGEN2.PNG')
        pixmap = pixmap.scaledToHeight(400)
        pic.setPixmap(pixmap)

        #Widgets Main
        LabelTitleReport = QtGui.QLabel("REPORT : ",self)
        LabelTitleReport.move(10,40)
        LabelTitleReport.setStyleSheet('font-size: 12pt')
       
        BtnMasterList = QtGui.QPushButton('Master List Report',self)
        BtnMasterList.move(10,70)
        BtnMasterList.setStyleSheet('font-size: 10pt')
        BtnMasterList.resize(160,25)       
        
        BtnNewFD = QtGui.QPushButton('New FD Tracking Report',self)        
        BtnNewFD.move(10,100)
        BtnNewFD.setStyleSheet('font-size: 10pt')
        BtnNewFD.resize(160,25)

        BtnDueDate = QtGui.QPushButton('Current Due Report',self)        
        BtnDueDate.move(10,130)
        BtnDueDate.setStyleSheet('font-size: 10pt')
        BtnDueDate.resize(160,25)

        BtnKPIUpdate = QtGui.QPushButton('SL KPI Report',self)        
        BtnKPIUpdate.move(10,160)
        BtnKPIUpdate.setStyleSheet('font-size: 10pt')
        BtnKPIUpdate.resize(160,25)

        BtnKPICritical = QtGui.QPushButton('SL w/ Less Personal',self)        
        BtnKPICritical.move(10,190)
        BtnKPICritical.setStyleSheet('font-size: 10pt')
        BtnKPICritical.resize(160,25)

        BtnNewSLtracking = QtGui.QPushButton('New SLC Report w/ info',self)        
        BtnNewSLtracking.move(10,220)
        BtnNewSLtracking.setStyleSheet('font-size: 10pt')
        BtnNewSLtracking.resize(160,25)

        BtnNewSLtrackingKPI = QtGui.QPushButton('New SLC Report KPI',self)        
        BtnNewSLtrackingKPI.move(10,250)
        BtnNewSLtrackingKPI.setStyleSheet('font-size: 10pt')
        BtnNewSLtrackingKPI.resize(160,25)

        #Button Events
        BtnMasterList.clicked.connect(self.MasterListClick)
        BtnNewFD.clicked.connect(self.NewFDTrackingMenu)
        BtnDueDate.clicked.connect(self.DueDateUpdate)
        BtnKPIUpdate.clicked.connect(self.KPIUpdate)
        BtnKPICritical.clicked.connect(self.KPIUpdateCritical)
        BtnNewSLtracking.clicked.connect(self.NewSLCTracking)
        BtnNewSLtrackingKPI.clicked.connect(self.NewSLTrackingKPI)
                       
        #self.move(500,300)
        self.centerOnScreen()
        self.setWindowTitle('New Report Generator version 1.5')
        self.setWindowIcon(QtGui.QIcon('replogo.png'))
        self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)
        self.setFixedSize(470,300)
        self.show()

#--------------------------------------------------------------------------------------

    def MasterListClick(self):
        
        self.app2 = QtGui.QMainWindow(self)
        self.app2.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app2.setWindowTitle(self.tr('Master List Report'))

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app2.move((resolution.width() / 2) - (self.frameSize().width() / 1.70),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
        

        self.LabelFDAcct = QtGui.QLabel("SL Account : ")
        self.EntryFDAcct = QtGui.QLineEdit(self)
        self.LabelStatus = QtGui.QLabel("Status : ")
        self.ComboStatus = QtGui.QComboBox(self)
        self.ComboStatus.addItem(" ")
        self.ComboStatus.addItem('A')
        self.ComboStatus.addItem('I')
        self.ComboStatus.addItem('R')
        self.ComboStatus.addItem('D')
        
        self.BtnGenerateMstr = QtGui.QPushButton('Generate')
        self.BtnGenerateMstrCSV = QtGui.QPushButton('Export')

        #Command that will display the widget's
        self.CentralWidget = QtGui.QWidget()
        self.CentralWidgetLayout = QtGui.QHBoxLayout()
        
        self.CentralWidgetLayout.addWidget(self.LabelFDAcct)
        self.CentralWidgetLayout.addWidget(self.EntryFDAcct)
        self.CentralWidgetLayout.addWidget(self.LabelStatus)
        self.CentralWidgetLayout.addWidget(self.ComboStatus)
        self.CentralWidgetLayout.addWidget(self.BtnGenerateMstr)
        self.CentralWidgetLayout.addWidget(self.BtnGenerateMstrCSV)
        
        self.CentralWidget.setLayout(self.CentralWidgetLayout)        
        self.app2.setCentralWidget(self.CentralWidget)

        #Button Event
        self.BtnGenerateMstr.clicked.connect(self.MasterListCall)
        self.BtnGenerateMstrCSV.clicked.connect(self.ExportCSVMaster) 

        
        self.app2.setFixedSize(430,100)        
        self.app2.show()

#-----------------------------------------------------------------------------------------------
    def ConnectToSQL(self):       

        #Connect to SQL
        self.connection = pypyodbc.connect('Driver={SQL Server};'
                                      'Server=MNLW3DRM-ANG;'
                                      'Database=DRMPOS;'
                                      'uid=dsdvp;pwd=password')
        self.cursor=self.connection.cursor()

        
#-----------------------------------------------------------------------------------------------
    def MasterListCall(self):
                
        self.ConnectToSQL()
       
        if self.ComboStatus.currentText()==' ':
            self.statusMaster=("A','I")
        else:
            self.statusMaster=self.ComboStatus.currentText()
       
        
        #SQL Scripts
        self.SQLCommand=(
            """select
            A.DistId AccountNo, 
            A.Name,
            A.Status,
            CONVERT(VARCHAR(10),A.AppointDate,10) AppointDate,
            (A.DealerCreditAmt-A.ARBalanceAmt) AvailableCL,
            A.Street Address,
            A.Zip,
            A.MobilePhoneNo,
            A.PastdueAmt,
            A.LOA,
            (Select Name From Dealer where DistID = A.ManagerID) SL_Name
            from Dealer A          
            where ManagerID='"""+self.EntryFDAcct.text()+"""'
            and status in ('"""+self.statusMaster+"""')
            Order by NAME"""            
                )

        self.cursor.execute(self.SQLCommand)
        self.results=self.cursor.fetchone()

        #PDF format

        self.styleP = getSampleStyleSheet()
        self.styleBH = self.styleP["Normal"]
        self.styleBH.alignment = TA_CENTER
        self.styleBH.fontSize=7


        self.doc = SimpleDocTemplate("sql.pdf", pagesize = letter, rightMargin=5,lefMargin=10, topMargin=45,bottomMargin=12)
        self.doc.pagesize=landscape(letter)
        self.elements=[]
        self.data=[]

        self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                        ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                        ('BOX',(0,0),(-1,-1),0.50,colors.red),
                        ('FONTSIZE',(0,0),(-1,-1),7),
                        ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                        ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                        ])

        self.data.append(['ACCOUNT NO.',
                     'NAME',
                     'STATUS',
                     'REC.DATE',
                     'AVAILABLE CL',
                     'ADDRESS',         
                     'CONTACT',
                     'PASTDUE',
                     'LOA'
                     ])

                    
        #Get Data from SQL to PDF

        try:        
            self.SLName = str(self.results[10])

               
            while self.results:        
                   
                    self.Address = Paragraph(str(self.results[5])+' ZIP- '+str(self.results[6]),self.styleBH)
                    
                    self.data.append([str(self.results[0]),
                          str(self.results[1]),
                          str(self.results[2]),
                          str(self.results[3]),
                          str(self.results[4]),
                          self.Address,                                              
                          str(self.results[7]),             
                          str(self.results[8]),
                          str(self.results[9])])

                    self.results=self.cursor.fetchone()
     
                         
            self.t=Table(self.data,colWidths=(24*mm,52*mm,12*mm,16*mm,20*mm,104*mm,20*mm,14*mm,9*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                                
                 
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.Firstpage, onLaterPages=self.addPageNumber)

            self.connection.close()

                
                #Open PDF
            
            
        except:
               self.errorMsg()
        os.startfile('sql.pdf')
#-------------------------------------------------------------------------------------------    

    def addPageNumber(self,canvas, doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

#--------------------------------------------------------------------------------------------
    def Firstpage(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')
        text2 = "*** AVON ANGELES BRANCH  *  MASTER LIST REPORT  *  As of "+time.strftime("%m/%d/%Y")+" ***" 
        canvas.drawRightString(202*mm, 206*mm, text2)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]

        text3 = "SL Name : "+self.SLName
        text_width3 = stringWidth(text3,'Helvetica',8)
        canvas.drawRightString(PAGE_WIDTH - text_width3, 200*mm, text3)
#--------------------------------------------------------------------------------------------------------------------
    def NewFDTrackingMenu(self):
        
        self.app7 = QtGui.QMainWindow(self)
        self.app7.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app7.setWindowTitle(self.tr('New FD Tracking Report'))

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app7.move((resolution.width() / 2) - (self.frameSize().width() / 1.60),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
                
        self.BtnGenerateNewFDPerSL = QtGui.QPushButton('Per Sales Leader')
        self.BtnGenerateNewFDAll = QtGui.QPushButton('All SL')

        #Command that will display the widget's
        self.CentralWidget = QtGui.QWidget()
        self.CentralWidgetLayout = QtGui.QHBoxLayout()
        
        self.CentralWidgetLayout.addWidget(self.BtnGenerateNewFDPerSL)
        self.CentralWidgetLayout.addWidget(self.BtnGenerateNewFDAll)
        
        self.CentralWidget.setLayout(self.CentralWidgetLayout)        
        self.app7.setCentralWidget(self.CentralWidget)

        #Button Event
        self.BtnGenerateNewFDPerSL.clicked.connect(self.NewFDTracking)
        self.BtnGenerateNewFDAll.clicked.connect(self.NewFDTrackingALL) 

        
        self.app7.setFixedSize(350,70)        
        self.app7.show()

#----------------------------------------------------------------------------------------------
    def NewFDTracking(self):

        self.app3 = QtGui.QMainWindow(self)
        self.app3.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app3.setWindowTitle(self.tr('New FD Tracking'))
        #self.app3.move(220,360)
        self.app3.setFixedSize(1010,100)

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app3.move((resolution.width() / 2) - (self.frameSize().width() / 1),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
        

        self.LabelSLAcct2 = QtGui.QLabel("SL Account : ")    
        self.EntrySLAcct2 = QtGui.QLineEdit(self)
        self.LabelYearMonth2 = QtGui.QLabel("Month:(YYYYMM)")
        self.EntryYearMonth2 = QtGui.QLineEdit(self)
        self.LabelDateFrom2 = QtGui.QLabel("From:(YYYY-MM-DD)")
        self.EntryDateFrom2 = QtGui.QLineEdit(self)
        self.BttnDateFrom2 = QtGui.QPushButton("...")
        self.LabelDateTo2 = QtGui.QLabel("To:(YYYY-MM-DD)")
        self.EntryDateTo2 = QtGui.QLineEdit(self)
        self.BttnDateTo2 = QtGui.QPushButton("...")
        self.BttnGenerate2 = QtGui.QPushButton("Generate")
        self.BttnExport2 = QtGui.QPushButton("Export")
                             
        #Command that will display the widget's
        self.CentralWidget2 = QtGui.QWidget()
        self.CentralWidgetLayout2 = QtGui.QHBoxLayout()
        self.CentralWidgetLayout2.addStretch()

        self.CentralWidgetLayout2.addWidget(self.LabelSLAcct2,1)
        self.CentralWidgetLayout2.addWidget(self.EntrySLAcct2,2)
        self.CentralWidgetLayout2.addWidget(self.LabelYearMonth2,1)
        self.CentralWidgetLayout2.addWidget(self.EntryYearMonth2,1)
        self.CentralWidgetLayout2.addWidget(self.LabelDateFrom2,1)
        self.CentralWidgetLayout2.addWidget(self.EntryDateFrom2,1)
        self.CentralWidgetLayout2.addWidget(self.BttnDateFrom2,1)        
        self.CentralWidgetLayout2.addWidget(self.LabelDateTo2,1)
        self.CentralWidgetLayout2.addWidget(self.EntryDateTo2,1)
        self.CentralWidgetLayout2.addWidget(self.BttnDateTo2,1)
        self.CentralWidgetLayout2.addWidget(self.BttnGenerate2,1)
        self.CentralWidgetLayout2.addWidget(self.BttnExport2,1)
        
        
        self.CentralWidget2.setLayout(self.CentralWidgetLayout2)
        self.app3.setCentralWidget(self.CentralWidget2)

        self.calendar = QtGui.QCalendarWidget(self)
        self.datecal = self.calendar.selectedDate()

        self.EntryYearMonth2.setText(self.datecal.toString('yyyyMM'))

        #Command click date pick buttons
        self.BttnDateFrom2.clicked.connect(self.DatePickFrom)
        self.BttnDateTo2.clicked.connect(self.DatePickTo)
        self.BttnGenerate2.clicked.connect(self.NewFDReportGen)
        self.BttnExport2.clicked.connect(self.ExportCSVNewFD)
                       
        self.app3.show()
       
#----------------------------------------------------------------------------------------
    def DatePickFrom(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(760,430)
        self.datepick.setFixedSize(320,300)

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))  
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateFrom)
     
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateFrom2.setText(self.datepick.date.toString('yyyy-MM-dd'))          
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()
              
    def showDateFrom(self, date):
        
        self.EntryDateFrom2.setText(date.toString('yyyy-MM-dd'))
   
#----------------------------------------------------------------------------------------
    def DatePickTo(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(1000,430)
        self.datepick.setFixedSize(320,300)

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateTo)
        
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateTo2.setText(self.datepick.date.toString('yyyy-MM-dd'))        
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()      

      
    def showDateTo(self, date):
        
        self.EntryDateTo2.setText(date.toString('yyyy-MM-dd'))

#--------------------------------------------------------------------------------------------------------------
    def NewFDReportGen(self):


        self.ConnectToSQL()
        self.SQLCommand=(
            """select
                A.managerid		    SL_No
                ,(Select Name From Dealer where DistID = A.ManagerID)   SL_Name
                ,A.Distid  
                ,A.Name
                ,A.Street Address
                ,A.Zip
                ,A.MobilePhoneNo
                ,CONVERT(VARCHAR(10),A.AppointDate,10) AppointDate
                ,SUM(B.CTDDiscRateBase) as TOTALDPS
                ,A.PastDueAmt
                from Dealer A with (NOLOCK)
                Inner Join discctdsales B
                on A. DistID=B.DistID, enum c
                where ManagerID='"""+self.EntrySLAcct2.text()+"""'
                and B.campaign >= '""" +self.EntryYearMonth2.text()+"""'--- change campaign
                and A.AppointDate >='"""+self.EntryDateFrom2.text()+"""' and A.AppointDate <='"""+self.EntryDateTo2.text()+"""' --- change appdate
                and A.Status in ('A','I')
                and A.TitleID ='01'
                and B.ProductDiscountGroupID in ('CFT','HERBALCARE','HOMESTYLE','NCFT')
                and a.starlevel = c.enumvalue and c.tablename ='dealer'
                and c.columnname ='StarLevel'
                ---and b.CTDDiscRateBase>='0'
                ---and A.PastDueAmt>='0'
                group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.AppointDate,A.PastdueAmt
                order by A.distid
             """)

        self.cursor.execute(self.SQLCommand)
        self.results=self.cursor.fetchone()

        #PDF format

        self.styleP = getSampleStyleSheet()
        self.styleBH = self.styleP["Normal"]
        self.styleBH.alignment = TA_CENTER
        self.styleBH.fontSize=7


        self.doc = SimpleDocTemplate("sqlnewfd.pdf", pagesize = letter, rightMargin=10,lefMargin=10, topMargin=45,bottomMargin=12)
        self.doc.pagesize=landscape(letter)
        self.elements=[]
        self.data=[]

        self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                        ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                        ('BOX',(0,0),(-1,-1),0.50,colors.red),
                        ('FONTSIZE',(0,0),(-1,-1),9),
                        ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                        ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                        ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                        ])

        self.data.append(['ACCOUNT NO.',
                     'NAME',
                     'ADDRESS', 
                     'CONTACT #',
                     'APPT DATE',
                     'TOTAL DPS',
                     'PASTDUE',
                     ])

                    
        #Get Data from SQL to PDF

        try:        
            self.SLName = str(self.results[1])      
            self.TotaDPSNewFD = 0
            self.CountFD = 0
            Pastdue = 0
        
            while self.results:        
                   
                    self.Address = Paragraph(str(self.results[4])+' ZIP- '+str(self.results[5]),self.styleBH)
                    
                    self.data.append([str(self.results[2]),
                          str(self.results[3]),
                          self.Address,
                          str(self.results[6]),                                                                        
                          str(self.results[7]),             
                          str(self.results[8]),
                          str(self.results[9])
                                      ])
                       
                    #compute total DPSO
                    self.TotaDPSNewFD +=float(self.results[8])
                    self.CountFD += 1
                    Pastdue += float(self.results[9])
                                     
                    self.results=self.cursor.fetchone()

            self.data.append(['TOTAL RECRUIT :',
                         self.CountFD,
                         '', 
                         '',
                         'TOTAL DPS :',
                         round(self.TotaDPSNewFD,2),
                         round(Pastdue,2)   
                         ])

                         
            self.t=Table(self.data,colWidths=(28*mm,60*mm,93*mm,25*mm,20*mm,24*mm,20*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageNewFD, onLaterPages=self.addPageNumber)

            self.connection.close()
            

                
        except:
               self.errorMsg()
                    #Open PDF
        os.startfile('sqlnewfd.pdf')

#-----------------------------------------------------------------------------------------------------------------

    def FirstpageNewFD(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')
        text2 = "*** AVON ANGELES BRANCH  *  New FD TRACKING  *  As of "+time.strftime("%m/%d/%Y")+" ***" 
        canvas.drawRightString(202*mm, 206*mm, text2)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]

        text3 = "SL Name : "+self.SLName
        text_width3 = stringWidth(text3,'Helvetica',8)
        canvas.drawRightString(PAGE_WIDTH - text_width3, 200*mm, text3)

#----------------------------------------------------------------------------------------------
    def DueDateUpdate(self):
        
        self.app4 = QtGui.QMainWindow(self)
        self.app4.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app4.setWindowTitle(self.tr('CURRENT DUE REPORT'))
        #self.app4.move(550,360)
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app4.move((resolution.width() / 2) - (self.frameSize().width() / 1.6),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

        self.LabelSLAcctDue = QtGui.QLabel("SL Account : ")
        self.EntrySLAcctDue = QtGui.QLineEdit(self)        
        
        self.BtnGenerateDue = QtGui.QPushButton('Generate')
        self.BtnGenerateDueAll = QtGui.QPushButton('Generate All')
        self.BtnExportDueAll = QtGui.QPushButton('Export')

        #Command that will display the widget's
        self.CentralWidget4 = QtGui.QWidget()
        self.CentralWidgetLayout4 = QtGui.QHBoxLayout()
        
        self.CentralWidgetLayout4.addWidget(self.LabelSLAcctDue)
        self.CentralWidgetLayout4.addWidget(self.EntrySLAcctDue)
        self.CentralWidgetLayout4.addWidget(self.BtnGenerateDue)
        self.CentralWidgetLayout4.addWidget(self.BtnGenerateDueAll)
        self.CentralWidgetLayout4.addWidget(self.BtnExportDueAll)
        
        self.CentralWidget4.setLayout(self.CentralWidgetLayout4)        
        self.app4.setCentralWidget(self.CentralWidget4)

        #Button Event
        self.BtnGenerateDue.clicked.connect(self.DueDateUpdateClick)
        self.BtnGenerateDueAll.clicked.connect(self.DueDateUpdateClickAll)
        self.BtnExportDueAll.clicked.connect(self.exportCurrentDue)
        
        self.app4.setFixedSize(430,100)        
        self.app4.show()

#-------------------------------------------------------------------------------------
    def DueDateUpdateClick(self):
    
        self.ConnectToSQL()
        
        self.SQLCommand=(
            """
            select  
            (Select Name From Dealer where DistID = A.ManagerID)SL_Name 
            ,A.Distid   
            ,A.Name 
            ,A.Mobilephoneno 
            ,A.CurrDueAMT ,
            c.dispvalue as Segment from Dealer A 
            with (NOLOCK) Inner Join discctdsales B on A. 
            DistID=B.DistID,   enum c 
            where ManagerID='"""+self.EntrySLAcctDue.text()+"""' 
            and A.Status in ('A','I','R') 
            and A.TitleID in ('01','02','03','04','05') 
            and A.currdueamt > '0.01' 
            and a.starlevel = c.enumvalue 
            and c.tablename ='dealer' 
            and c.columnname ='StarLevel' 
            group by a.managerid
            , A.Distid 
            ,A.Name
            ,A.TITLEID
            ,A.Status
            ,A.Address
            ,A.District
            ,A.ZIP
            ,A.Dayphoneno
            ,A.Mobilephoneno
            ,A.Fixedcredit
            ,A.ARBalanceAMT
            ,A.CurrDueAmt
            ,A.PaytermGroupID
            ,A.PastDueAmt
            ,A.StarLevel
            ,A.AppointDate
            ,c.dispvalue 
            order by A.CurrDueAMT DESC
             """)

        self.cursor.execute(self.SQLCommand)
        self.results=self.cursor.fetchone()

        #PDF format

        self.styleP = getSampleStyleSheet()
        self.styleBH = self.styleP["Normal"]
        self.styleBH.alignment = TA_CENTER
        self.styleBH.fontSize=7


        self.doc = SimpleDocTemplate("sqldue.pdf", pagesize = letter, rightMargin=10,lefMargin=5, topMargin=45,bottomMargin=12)
        #self.doc.pagesize=landscape(letter)
        self.elements=[]
        self.data=[]

        self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                        ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                        ('BOX',(0,0),(-1,-1),0.50,colors.red),
                        ('FONTSIZE',(0,0),(-1,-1),8),
                        ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                        ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                        ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                        ])

        self.data.append(['SL NAME',
                     'FD ACCOUNT',
                     'FD NAME', 
                     'CONTACT #',
                     'CURRENT DUE',
                     'SEGMENT',
                     ])

                    
        #Get Data from SQL to PDF

        try:        
            self.SLName = str(self.results[1])      
            self.TotaLDue = 0
            self.CountFDDue = 0
        
            while self.results:        
                   
                                       
                    self.data.append([str(self.results[0]),
                          str(self.results[1]),                     
                          str(self.results[2]),                                                                        
                          str(self.results[3]),
                          str(self.results[4]),
                          str(self.results[5])])
                       
                    #compute total DPS
                    self.TotaLDue +=float(self.results[4])
                    self.CountFDDue += 1
                                     
                    self.results=self.cursor.fetchone()

            self.data.append(['TOTAL FD w/Due :',
                         self.CountFDDue,
                         '', 
                         'TOTAL :',
                         round(self.TotaLDue,2),
                         '',
                         ])

                         
            self.t=Table(self.data,colWidths=(52*mm,24*mm,58*mm,22*mm,23*mm,25*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageDuedate, onLaterPages=self.addPageNumber)

            self.connection.close()
            
            #Open PDF

                
        except:
               self.errorMsg()

        os.startfile('sqldue.pdf')
#------------------------------------------------------------------------------------------------------------
    def DueDateUpdateClickAll(self):
    
        self.ConnectToSQL()
        
        self.SQLCommand=(
            """
            select  
            (Select Name From Dealer where DistID = A.ManagerID)SL_Name 
            ,A.Distid   
            ,A.Name 
            ,A.Mobilephoneno 
            ,A.CurrDueAMT ,
            c.dispvalue as Segment from Dealer A 
            with (NOLOCK) Inner Join discctdsales B on A. 
            DistID=B.DistID,   enum c              
            where A.Status in ('A','I','R') 
            and A.TitleID in ('01','02','03','04','05') 
            and A.currdueamt > '0.01' 
            and a.starlevel = c.enumvalue 
            and c.tablename ='dealer' 
            and c.columnname ='StarLevel' 
            group by a.managerid
            , A.Distid 
            ,A.Name
            ,A.TITLEID
            ,A.Status
            ,A.Address
            ,A.District
            ,A.ZIP
            ,A.Dayphoneno
            ,A.Mobilephoneno
            ,A.Fixedcredit
            ,A.ARBalanceAMT
            ,A.CurrDueAmt
            ,A.PaytermGroupID
            ,A.PastDueAmt
            ,A.StarLevel
            ,A.AppointDate
            ,c.dispvalue 
            order by A.CurrDueAMT DESC
             """)

        self.cursor.execute(self.SQLCommand)
        self.results=self.cursor.fetchone()

        #PDF format

        self.styleP = getSampleStyleSheet()
        self.styleBH = self.styleP["Normal"]
        self.styleBH.alignment = TA_CENTER
        self.styleBH.fontSize=7


        self.doc = SimpleDocTemplate("sqldue.pdf", pagesize = letter, rightMargin=10,lefMargin=5, topMargin=45,bottomMargin=12)
        #self.doc.pagesize=landscape(letter)
        self.elements=[]
        self.data=[]

        self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                        ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                        ('BOX',(0,0),(-1,-1),0.50,colors.red),
                        ('FONTSIZE',(0,0),(-1,-1),8),
                        ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                        ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                        ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                        ])

        self.data.append(['SL NAME',
                     'FD ACCOUNT',
                     'FD NAME', 
                     'CONTACT #',
                     'CURRENT DUE',
                     'SEGMENT',
                     ])

                    
        #Get Data from SQL to PDF

        try:        
            self.SLName = str(self.results[1])      
            self.TotaLDue = 0
            self.CountFDDue = 0
        
            while self.results:        
                   
                                       
                    self.data.append([str(self.results[0]),
                          str(self.results[1]),                     
                          str(self.results[2]),                                                                        
                          str(self.results[3]),
                          str(self.results[4]),
                          str(self.results[5])])
                       
                    #compute total DPS
                    self.TotaLDue +=float(self.results[4])
                    self.CountFDDue += 1
                                     
                    self.results=self.cursor.fetchone()

            self.data.append(['TOTAL FD w/Due :',
                         self.CountFDDue,
                         '', 
                         'TOTAL PDA :',
                         round(self.TotaLDue,2),
                         '',
                         ])

                         
            self.t=Table(self.data,colWidths=(52*mm,24*mm,58*mm,22*mm,23*mm,25*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageDuedate, onLaterPages=self.addPageNumber)

            self.connection.close()
            

                
        except:
               self.errorMsg()
                #Open PDF
        os.startfile('sqldue.pdf')

#------------------------------------------------------------------------------------------------------------------
    def FirstpageDuedate(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')
        text2 = "*** AVON ANGELES BRANCH  *  CURRENT DUE DATE  *  As of "+time.strftime("%m/%d/%Y")+" ***" 
        canvas.drawRightString(170*mm, 265*mm, text2)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]

#----------------------------------------------------------------------------------------------
    def KPIUpdate(self):
        
        self.app5 = QtGui.QMainWindow(self)
        self.app5.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app5.setWindowTitle(self.tr('SL KPI UPDATE'))
        #self.app5.move(550,360)
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app5.move((resolution.width() / 2) - (self.frameSize().width() / 1.6),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

        self.LabelCampYear = QtGui.QLabel("YEAR : (YYYY)")
        self.EntryCampYear = QtGui.QLineEdit(self)
        self.LabelCampMonth = QtGui.QLabel("MONTH :(MM)")
        self.EntryCampMonth = QtGui.QLineEdit(self)    
        
        self.BtnGenerateKPI = QtGui.QPushButton('Generate')
        self.BtnExportKPI = QtGui.QPushButton('Export')
        

        #Command that will display the widget's
        self.CentralWidget5 = QtGui.QWidget()
        self.CentralWidgetLayout5 = QtGui.QHBoxLayout()
        
        self.CentralWidgetLayout5.addWidget(self.LabelCampYear,1)
        self.CentralWidgetLayout5.addWidget(self.EntryCampYear,2)
        self.CentralWidgetLayout5.addWidget(self.LabelCampMonth,1)
        self.CentralWidgetLayout5.addWidget(self.EntryCampMonth,2)
        self.CentralWidgetLayout5.addWidget(self.BtnGenerateKPI,1)
        self.CentralWidgetLayout5.addWidget(self.BtnExportKPI,2)
        

        
        self.CentralWidget5.setLayout(self.CentralWidgetLayout5)        
        self.app5.setCentralWidget(self.CentralWidget5)

        #defaultinput of date campaign
        self.calendarKPI = QtGui.QCalendarWidget(self)
        self.datecalKPI = self.calendarKPI.selectedDate()


        self.EntryCampYear.setText(self.datecalKPI.toString('yyyy'))
        self.EntryCampMonth.setText(self.datecalKPI.toString('MM'))


        #Button Event
        self.BtnGenerateKPI.clicked.connect(self.SLKPIUpdateCheck)
        self.BtnExportKPI.clicked.connect(self.ExportSLKPIUpdateCheck)


             
        self.app5.setFixedSize(420,100)        
        self.app5.show()

     
#------------------------------------------------------------------------------------------------------------------------
    def SLKPIUpdateCheck(self):

        if self.EntryCampYear.text() > self.datecalKPI.toString('yyyy'):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")
        elif self.EntryCampYear.text() < str(2009):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")

        else:

            if self.EntryCampMonth.text() == '02':
                self.DayofMonth = '28'
                self.SLKPIUpdateClick()

            elif self.EntryCampMonth.text() == '04':
                self.DayofMonth = '30'
                self.SLKPIUpdateClick()
            elif self.EntryCampMonth.text() == '06':
                self.DayofMonth = '30'
                self.SLKPIUpdateClick()
            elif self.EntryCampMonth.text() == '09':
                self.DayofMonth = '30'
                self.SLKPIUpdateClick()                
            elif self.EntryCampMonth.text() == '11':
                self.DayofMonth = '30'
                self.SLKPIUpdateClick()
            elif self.EntryCampMonth.text() == '01' or self.EntryCampMonth.text() == '03' or self.EntryCampMonth.text() == '05' or self.EntryCampMonth.text() == '07' or \
                 self.EntryCampMonth.text() == '08' or self.EntryCampMonth.text() == '10' or self.EntryCampMonth.text() == '12':
                self.DayofMonth = '31'
                self.SLKPIUpdateClick()
            else :
                QtGui.QMessageBox.information(self, 'Invalid Month input',
                         "Sorry, the Month you input is invalid. Please try again.")



#-----------------------------------------------------------------------------------------------------------------------------------------        
    def SLKPIUpdateClick(self):

        
        self.ConnectToSQL()


        try:
            self.SQLCommand=(
                """
                    SELECT
                Name SLName
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID) AS Personal
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) FirstGen
                              --, TotalDPS,
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                     WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                                       AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                     GROUP BY Distid		
                                    HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND TitleID <= '03'
                     AND AppointDate BETWEEN '"""+self.EntryCampMonth.text()+"""/01/"""+self.EntryCampYear.text()+"""' 
                                     AND '"""+self.EntryCampMonth.text()+"""/"""+self.DayofMonth+"""/"""+self.EntryCampYear.text()+"""'
                  ) Appts
               FROM Dealer DL WITH(NOLOCK)
                    WHERE TitleID IN ('04','05')
                   AND (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) !=0

                  order by dl.FirstGen DESC


                 """)

            self.cursor.execute(self.SQLCommand)
            self.results=self.cursor.fetchone()

            #PDF format

            self.styleP = getSampleStyleSheet()
            self.styleBH = self.styleP["Normal"]
            self.styleBH.alignment = TA_CENTER
            self.styleBH.fontSize=12


            self.doc = SimpleDocTemplate("sqlKPI.pdf", pagesize = letter, rightMargin=10,lefMargin=5, topMargin=60,bottomMargin=12)
            #self.doc.pagesize=landscape(letter)
            self.elements=[]
            self.data=[]

            self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                            ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                            ('BOX',(0,0),(-1,-1),0.50,colors.red),
                            ('FONTSIZE',(0,0),(-1,-1),11),
                            ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                            ])

            self.data.append([
                         'RANK',
                         'SL Name',         
                         'Personal',
                         'FirstGen',
                         'Total',
                         'Active',
                         'Appoint',
                         ])

                        
            #Get Data from SQL to PDF

          
            self.TotaLSalesSLPersonal = 0
            self.TotalSalesSLFirst = 0
            self.TotalSalesSL = 0
            self.TotalSalesActive = 0
            self.TotalSalesRec = 0
            self.CountSL = 1
            locale.setlocale( locale.LC_ALL, '' ),
        
            while self.results:        

                                       
                    self.data.append([self.CountSL,
                          str(self.results[0]),
                         # str(self.results[1]),                     
                                     
                          locale.format("%d",float(self.results[1]),grouping=True),
                          locale.format("%d",float(self.results[2]),grouping=True),
                          locale.format("%d",float(self.results[1])+float(self.results[2]),grouping=True),
                          str(self.results[3]),
                          str(self.results[4])
          
                                      ])
                       
                    #compute total DPS
                    self.CountSL +=1
                    self.TotaLSalesSLPersonal +=int(self.results[1])
                    self.TotalSalesSLFirst +=int(self.results[2])                              
                    self.TotalSalesActive +=int(self.results[3])
                    self.TotalSalesRec +=int(self.results[4])

                    self.results=self.cursor.fetchone()
                    
            #Last Table total
            self.data.append([
                             '',
                             '',                         
                             locale.format("%d",float(self.TotaLSalesSLPersonal),grouping=True),
                             locale.format("%d",float(self.TotalSalesSLFirst),grouping=True),
                             locale.format("%d",float((self.TotalSalesSLFirst)+(self.TotaLSalesSLPersonal)),grouping=True),
                             locale.format("%d",float(self.TotalSalesActive),grouping=True),
                             self.TotalSalesRec,
                             ])

                         
            self.t=Table(self.data,colWidths=(12*mm,85*mm,20*mm,26*mm,28*mm,15*mm,16*mm,),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageKPI, onLaterPages=self.addPageNumber)

            self.connection.close()

        except:
               self.errorMsg()
        
        #Open PDF
        os.startfile('sqlKPI.pdf')

                

#------------------------------------------------------------------------------------------------------------------
    def FirstpageKPI(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')

        if self.EntryCampMonth.text() == '01':
            monthpdf = 'JANUARY'
        elif self.EntryCampMonth.text() == '02':
            monthpdf = 'FEBRUARY'
        elif self.EntryCampMonth.text() == '03':
            monthpdf = 'MARCH'
        elif self.EntryCampMonth.text() == '04':
            monthpdf = 'APRIL'
        elif self.EntryCampMonth.text() == '05':
            monthpdf = 'MAY'
        elif self.EntryCampMonth.text() == '06':
            monthpdf = 'JUNE'
        elif self.EntryCampMonth.text() == '07':
            monthpdf = 'JULY'
        elif self.EntryCampMonth.text() == '08':
            monthpdf = 'AUGUST'
        elif self.EntryCampMonth.text() == '09':
            monthpdf = 'SEPTEMBER'
        elif self.EntryCampMonth.text() == '10':
            monthpdf = 'OCTOBER'
        elif self.EntryCampMonth.text() == '11':
            monthpdf = 'NOVEMBER'
        elif self.EntryCampMonth.text() == '12':
            monthpdf = 'DECEMBER'
        
        
        text2c = "AVON ANGELES BRANCH"
        canvas.drawRightString(135*mm, 267*mm, text2c)
        
        text3c = monthpdf+" "+self.EntryCampYear.text()+" SL KPI REPORT"
        canvas.drawRightString(142*mm, 262*mm, text3c)

        text4c = "Print date : "+time.strftime("%m/%d/%Y") 
        canvas.drawRightString(130*mm, 257*mm, text4c)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]

#----------------------------------------------------------------------------------------------
    def KPIUpdateCritical(self):
        
        self.app5 = QtGui.QMainWindow(self)
        self.app5.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app5.setWindowTitle(self.tr('SL CRITICAL UPDATE'))
        #self.app5.move(550,360)
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app5.move((resolution.width() / 2) - (self.frameSize().width() / 1.6),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

        self.LabelCampYear = QtGui.QLabel("YEAR : (YYYY)")
        self.EntryCampYearC = QtGui.QLineEdit(self)
        self.LabelCampMonth = QtGui.QLabel("MONTH :(MM)")
        self.EntryCampMonthC = QtGui.QLineEdit(self)    
        
        self.BtnGenerateKPI = QtGui.QPushButton('Generate')
        self.BtnGenerateKPIExport = QtGui.QPushButton('Export')
        

        #Command that will display the widget's
        self.CentralWidget5 = QtGui.QWidget()
        self.CentralWidgetLayout5 = QtGui.QHBoxLayout()
        
        self.CentralWidgetLayout5.addWidget(self.LabelCampYear,1)
        self.CentralWidgetLayout5.addWidget(self.EntryCampYearC,2)
        self.CentralWidgetLayout5.addWidget(self.LabelCampMonth,1)
        self.CentralWidgetLayout5.addWidget(self.EntryCampMonthC,2)
        
        self.CentralWidgetLayout5.addWidget(self.BtnGenerateKPI,1)
        self.CentralWidgetLayout5.addWidget(self.BtnGenerateKPIExport,1)
        
        self.CentralWidget5.setLayout(self.CentralWidgetLayout5)        
        self.app5.setCentralWidget(self.CentralWidget5)

        #defaultinput of date campaign
        self.calendarKPI = QtGui.QCalendarWidget(self)
        self.datecalKPI = self.calendarKPI.selectedDate()


        self.EntryCampYearC.setText(self.datecalKPI.toString('yyyy'))
        self.EntryCampMonthC.setText(self.datecalKPI.toString('MM'))


        #Button Event
        self.BtnGenerateKPI.clicked.connect(self.SLKPIUpdateCheckCritical)
        self.BtnGenerateKPIExport.clicked.connect(self.ExportSLKPIUpdateCheckCritical)

             
        self.app5.setFixedSize(400,100)        
        self.app5.show()
#------------------------------------------------------------------------------------------------------------------------
    def SLKPIUpdateCheckCritical(self):

        if self.EntryCampYearC.text() > self.datecalKPI.toString('yyyy'):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")
        elif self.EntryCampYearC.text() < str(2009):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")

        else:

            if self.EntryCampMonthC.text() == '02':
                self.DayofMonthC = '28'
                self.SLKPIUpdateClickCritical()

            elif self.EntryCampMonthC.text() == '04':
                self.DayofMonthC = '30'
                self.SLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '06':
                self.DayofMonthC = '30'
                self.SLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '09':
                self.DayofMonthC = '30'
                self.SLKPIUpdateClickCritical()                
            elif self.EntryCampMonthC.text() == '11':
                self.DayofMonthC = '30'
                self.SLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '01' or self.EntryCampMonthC.text() == '03' or self.EntryCampMonthC.text() == '05' or self.EntryCampMonthC.text() == '07' or \
                 self.EntryCampMonthC.text() == '08' or self.EntryCampMonthC.text() == '10' or self.EntryCampMonthC.text() == '12':
                self.DayofMonthC = '31'
                self.SLKPIUpdateClickCritical()
            else :
                QtGui.QMessageBox.information(self, 'Invalid Month input',
                         "Sorry, the Month you input is invalid. Please try again.")



#-----------------------------------------------------------------------------------------------------------------------------------------        
    def SLKPIUpdateClickCritical(self):
        
        self.ConnectToSQL()
        
        try:
       
            self.SQLCommand=(
                """
                    SELECT 
                 Name SLName
                ,(SELECT Name FROM Dealer WITH(NOLOCK)
                   WHERE DistID = DL.ManagerID) UplineName
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID) AS Personal
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) FirstGen
                              --, TotalDPS,
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                     WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                                       AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                     GROUP BY Distid		
                                    HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND TitleID <= '03'
                     AND AppointDate BETWEEN '"""+self.EntryCampMonthC.text()+"""/01/"""+self.EntryCampYearC.text()+"""' 
                                     AND '"""+self.EntryCampMonthC.text()+"""/"""+self.DayofMonthC+"""/"""+self.EntryCampYearC.text()+"""'
                  ) Appts
               FROM Dealer DL WITH(NOLOCK)
                    WHERE TitleID IN ('04','05')
                    AND (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID)< 2000
                    AND ((SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID))+ (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID))>=30000
                                     
                  order by dl.FirstGen DESC


                 """)

            self.cursor.execute(self.SQLCommand)
            self.results=self.cursor.fetchone()

            #PDF format

            self.styleP = getSampleStyleSheet()
            self.styleBH = self.styleP["Normal"]
            self.styleBH.alignment = TA_CENTER
            self.styleBH.fontSize=10


            self.doc = SimpleDocTemplate("sqlKPICritical.pdf", pagesize = letter, rightMargin=0,lefMargin=5, topMargin=60,bottomMargin=12)
            #self.doc.pagesize=landscape(letter)
            self.elements=[]
            self.data=[]

            self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                            ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                            ('BOX',(0,0),(-1,-1),0.50,colors.red),
                            ('FONTSIZE',(0,0),(-1,-1),10),
                            ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                            ])

            self.data.append([
                         'SL Name',
                         'UplineName',
                         'Personal',
                         'FirstGen',
                         'Total',
                         'Active',
                         'Appoint',
                         ])

                        
            #Get Data from SQL to PDF

          
            self.TotaLSalesSLPersonal = 0
            self.TotalSalesSLFirst = 0
            self.TotalSalesSL = 0
            self.TotalSalesActive = 0
            self.TotalSalesRec = 0
            self.CountSL = 0
            locale.setlocale( locale.LC_ALL, '' ),
        
            while self.results:        

                                       
                    self.data.append([str(self.results[0]),
                          str(self.results[1]),                    
                                     
                          locale.format("%d",float(self.results[2]),grouping=True),
                          locale.format("%d",float(self.results[3]),grouping=True),
                          locale.format("%d",float(self.results[2])+float(self.results[3]),grouping=True),
                          str(self.results[4]),
                          str(self.results[5])
          
                                      ])
                       
                    #compute total DPS
                    self.CountSL +=1
                    self.TotaLSalesSLPersonal +=int(self.results[2])
                    self.TotalSalesSLFirst +=int(self.results[3])                              
                    self.TotalSalesActive +=int(self.results[4])
                    self.TotalSalesRec +=int(self.results[5])

                    self.results=self.cursor.fetchone()
                    
            #Last Table total
            self.data.append(['TOTAL SL :',
                             self.CountSL,             
                             locale.format("%d",float(self.TotaLSalesSLPersonal),grouping=True),
                             locale.format("%d",float(self.TotalSalesSLFirst),grouping=True),
                             locale.format("%d",float((self.TotalSalesSLFirst)+(self.TotaLSalesSLPersonal)),grouping=True),
                             locale.format("%d",float(self.TotalSalesActive),grouping=True),
                             self.TotalSalesRec,
                             ])

                         
            self.t=Table(self.data,colWidths=(60*mm,60*mm,17*mm,23*mm,25*mm,12*mm,13*mm,),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageKPICritical, onLaterPages=self.addPageNumber)

            self.connection.close()

        except:
               self.errorMsg()
        
        #Open PDF
        os.startfile('sqlKPICritical.pdf')
        
       
       

#------------------------------------------------------------------------------------------------------------------------
    def FirstpageKPICritical(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')

        if self.EntryCampMonthC.text() == '01':
            monthpdf = 'JANUARY'
        elif self.EntryCampMonthC.text() == '02':
            monthpdf = 'FEBRUARY'
        elif self.EntryCampMonthC.text() == '03':
            monthpdf = 'MARCH'
        elif self.EntryCampMonthC.text() == '04':
            monthpdf = 'APRIL'
        elif self.EntryCampMonthC.text() == '05':
            monthpdf = 'MAY'
        elif self.EntryCampMonthC.text() == '06':
            monthpdf = 'JUNE'
        elif self.EntryCampMonthC.text() == '07':
            monthpdf = 'JULY'
        elif self.EntryCampMonthC.text() == '08':
            monthpdf = 'AUGUST'
        elif self.EntryCampMonthC.text() == '09':
            monthpdf = 'SEPTEMBER'
        elif self.EntryCampMonthC.text() == '10':
            monthpdf = 'OCTOBER'
        elif self.EntryCampMonthC.text() == '11':
            monthpdf = 'NOVEMBER'
        elif self.EntryCampMonthC.text() == '12':
            monthpdf = 'DECEMBER'
        
        
        text2c = "AVON ANGELES BRANCH"
        canvas.drawRightString(135*mm, 267*mm, text2c)
        
        text3c = monthpdf+" "+self.EntryCampYearC.text()+" SL CRITICAL REPORT"
        canvas.drawRightString(148*mm, 262*mm, text3c)

        text4c = "Print date : "+time.strftime("%m/%d/%Y") 
        canvas.drawRightString(130*mm, 257*mm, text4c)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]
#-----------------------------------------------------------------------------------------------------------------------------------------     

    def NewSLCTracking(self):

        self.app6 = QtGui.QMainWindow(self)
        self.app6.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app6.setWindowTitle(self.tr('New SLC Tracking'))
        #self.app6.move(430,360)
        self.app6.setFixedSize(690,100)
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app6.move((resolution.width() / 2) - (self.frameSize().width() / 1.5),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

        self.LabelDateFrom2SL = QtGui.QLabel("From:(YYYY-MM-DD)")
        self.EntryDateFrom2SL = QtGui.QLineEdit(self)
        self.BttnDateFrom2SL = QtGui.QPushButton("...")
        self.LabelDateTo2SL = QtGui.QLabel("To:(YYYY-MM-DD)")
        self.EntryDateTo2SL = QtGui.QLineEdit(self)
        self.BttnDateTo2SL = QtGui.QPushButton("...")
        self.BttnGenerate2SL = QtGui.QPushButton("Generate")
        self.BttnGenerate2SLExport = QtGui.QPushButton("Export")
                             
        #Command that will display the widget's
        self.CentralWidget2SL = QtGui.QWidget()
        self.CentralWidgetLayout2SL = QtGui.QHBoxLayout()
        self.CentralWidgetLayout2SL.addStretch()
 
        self.CentralWidgetLayout2SL.addWidget(self.LabelDateFrom2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.EntryDateFrom2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.BttnDateFrom2SL,1)        
        self.CentralWidgetLayout2SL.addWidget(self.LabelDateTo2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.EntryDateTo2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.BttnDateTo2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.BttnGenerate2SL,1)
        self.CentralWidgetLayout2SL.addWidget(self.BttnGenerate2SLExport,1)
        
        self.CentralWidget2SL.setLayout(self.CentralWidgetLayout2SL)
        self.app6.setCentralWidget(self.CentralWidget2SL)

        self.calendarSL = QtGui.QCalendarWidget(self)
        self.datecalSL = self.calendarSL.selectedDate()

        #Command click date pick buttons
        self.BttnDateFrom2SL.clicked.connect(self.DatePickFromSL)
        self.BttnDateTo2SL.clicked.connect(self.DatePickToSL)
        self.BttnGenerate2SL.clicked.connect(self.NewSLReportGen)
        self.BttnGenerate2SLExport.clicked.connect(self.ExportNewSLReportGen)
                       
        self.app6.show()
       
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
    def DatePickFromSL(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(760,430)
        self.datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateFromSL)
     
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateFrom2SL.setText(self.datepick.date.toString('yyyy-MM-dd'))          
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()
              
    def showDateFromSL(self, date):
        
        self.EntryDateFrom2SL.setText(date.toString('yyyy-MM-dd'))
   
#----------------------------------------------------------------------------------------
    def DatePickToSL(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(1000,430)
        self.datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateToSL)
        
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateTo2SL.setText(self.datepick.date.toString('yyyy-MM-dd'))        
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()      

      
    def showDateToSL(self, date):
        
        self.EntryDateTo2SL.setText(date.toString('yyyy-MM-dd'))


#--------------------------------------------------------------------------------------------------------------
    def NewSLReportGen(self):


        self.ConnectToSQL()

        try:
            self.SQLCommand=(
                """select
                     A.Distid  
                    ,A.Name
                    ,(Select Name From Dealer where DistID = A.ManagerID)   SL_Name

                    ,A.Street Address
                    ,A.Zip
                    ,A.MobilePhoneNo
                    ,CONVERT(VARCHAR(10),A.TitleDate,10) AppointDate

                    from Dealer A with (NOLOCK)
                    Inner Join discctdsales B
                    on A. DistID=B.DistID, enum c
                    where A.TitleDate >='"""+self.EntryDateFrom2SL.text()+"""' and A.TitleDate <='"""+self.EntryDateTo2SL.text()+"""' --- change appdate
                    and A.Status in ('A','I')
                    and A.TitleID in ('04','05')
                    group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.TitleDate                
                    
                    order by A.TitleDate

                 """)

            self.cursor.execute(self.SQLCommand)
            self.results=self.cursor.fetchone()

            #PDF format

            self.styleP = getSampleStyleSheet()
            self.styleBH = self.styleP["Normal"]
            self.styleBH.alignment = TA_CENTER
            self.styleBH.fontSize=7


            self.doc = SimpleDocTemplate("sqlnewSL.pdf", pagesize = letter, rightMargin=10,lefMargin=10, topMargin=45,bottomMargin=12)
            self.doc.pagesize=landscape(letter)
            self.elements=[]
            self.data=[]

            self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                            ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                            ('BOX',(0,0),(-1,-1),0.50,colors.red),
                            ('FONTSIZE',(0,0),(-1,-1),8),
                            ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                            ])

            self.data.append(['SL Acct.',
                         'SL NAME',
                         'UPLINE NAME',
                         'ADDRESS', 
                         'CONTACT #',
                         'TITLE DATE',

                         ])

                        
            #Get Data from SQL to PDF

                    

            self.CountFD = 0
        
            while self.results:        
                   
                    self.Address = Paragraph(str(self.results[3])+' ZIP- '+str(self.results[4]),self.styleBH)
                    
                    self.data.append([str(self.results[0]),
                          str(self.results[1]),
                          str(self.results[2]),
                          self.Address,
                          str(self.results[5]),                                                                        
                          str(self.results[6])
                                      ])

                       
                    #compute total DPS

                    self.CountFD += 1
                                     
                    self.results=self.cursor.fetchone()

            self.data.append(['TOTAL SLC :',
                         self.CountFD,
                         '', 
                         '',
                         '',
                         ''

                         ])

                         
            self.t=Table(self.data,colWidths=(28*mm,62*mm,60*mm,80*mm,20*mm,20*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageNewSL, onLaterPages=self.addPageNumber)

            self.connection.close()
            
        except:
               self.errorMsg()
            
            #Open PDF
        os.startfile('sqlnewSL.pdf')
                

#-----------------------------------------------------------------------------------------------------------------
    def FirstpageNewSL(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')
        text2 = "*** AVON ANGELES BRANCH  *  New SL TRACKING  *  As of "+time.strftime("%m/%d/%Y")+" ***" 
        canvas.drawRightString(202*mm, 206*mm, text2)


#-----------------------------------------------------------------------------------------------------------------
    def ExportCSVMaster(self):

            
        if self.ComboStatus.currentText()==' ':
            statusMaster=("A','I")
        else:
            statusMaster=self.ComboStatus.currentText()

        
        self.ConnectToSQL()

          #SQL Scripts
        SQLCommand=(
            """select
            (Select Name From Dealer where DistID = A.ManagerID) SL_Name,
            A.DistId AccountNo, 
            A.Name,
            A.Status,
            A.AppointDate,
            (A.DealerCreditAmt-A.ARBalanceAmt) AvailableCL,
            A.Street Address,
            A.Zip,
            A.MobilePhoneNo,
            A.PastdueAmt,
            A.LOA            
            from Dealer A          
            where ManagerID='"""+self.EntryFDAcct.text()+"""'
            and status in ('"""+statusMaster+"""')
            Order by NAME"""            
                )

        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()

        try:
            with open('masterlist.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow(['SL NAME','ACCOUNT','NAME','STATUS','APPT DATE','AVAILABLE CL','ADDRESS','ZIP','CONTACT','PASTDUE','LOA'])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()

        except:
               self.errorMsg()
               
        os.startfile('masterlist.csv')
        #csvfile.closed()

        
#-----------------------------------------------------------------------------
    def ExportCSVNewFD(self):

        self.ConnectToSQL()
        
        SQLCommand=(
                """select
                A.managerid		    SL_No
                ,(Select Name From Dealer where DistID = A.ManagerID)   SL_Name
                ,A.Distid  
                ,A.Name
                ,A.Street Address
                ,A.Zip
                ,A.MobilePhoneNo
                ,CONVERT(VARCHAR(10),A.AppointDate,10) AppointDate
                ,SUM(B.CTDDiscRateBase) as TOTALDPS
                ,A.PastDueAmt
                from Dealer A with (NOLOCK)
                Inner Join discctdsales B
                on A. DistID=B.DistID, enum c
                where ManagerID='"""+self.EntrySLAcct2.text()+"""'
                and B.campaign >= '""" +self.EntryYearMonth2.text()+"""'--- change campaign
                and A.AppointDate >='"""+self.EntryDateFrom2.text()+"""' and A.AppointDate <='"""+self.EntryDateTo2.text()+"""' --- change appdate
                and A.Status in ('A','I')
                and A.TitleID ='01'
                and B.ProductDiscountGroupID in ('CFT','HERBALCARE','HOMESTYLE','NCFT')
                and a.starlevel = c.enumvalue and c.tablename ='dealer'
                and c.columnname ='StarLevel'
                ---and b.CTDDiscRateBase>='0'
                ---and A.PastDueAmt>='0'
                group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.AppointDate,A.PastdueAmt
                order by A.distid
             """)

        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()

        try:
            with open('newfd.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow(['SL NO.','SL NAME','ACCT NO','FD NAME','ADDRESS','ZIP','CONTACT','APPT DATE','TOTAL DPS','PASTDUE'])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
                
        except:
               self.errorMsg()        
            
        os.startfile('newfd.csv')
        csvfile.closed()


#---------------------------------------------------------------------------------------------------------
    def exportCurrentDue(self):

        self.ConnectToSQL()
        
        SQLCommand=(
            """
            select  
            (Select Name From Dealer where DistID = A.ManagerID)SL_Name 
            ,A.Distid   
            ,A.Name 
            ,A.Mobilephoneno 
            ,A.CurrDueAMT ,
            c.dispvalue as Segment from Dealer A 
            with (NOLOCK) Inner Join discctdsales B on A. 
            DistID=B.DistID,   enum c              
            where A.Status in ('A','I','R') 
            and A.TitleID in ('01','02','03','04','05') 
            and A.currdueamt > '0.01' 
            and a.starlevel = c.enumvalue 
            and c.tablename ='dealer' 
            and c.columnname ='StarLevel' 
            group by a.managerid
            , A.Distid 
            ,A.Name
            ,A.TITLEID
            ,A.Status
            ,A.Address
            ,A.District
            ,A.ZIP
            ,A.Dayphoneno
            ,A.Mobilephoneno
            ,A.Fixedcredit
            ,A.ARBalanceAMT
            ,A.CurrDueAmt
            ,A.PaytermGroupID
            ,A.PastDueAmt
            ,A.StarLevel
            ,A.AppointDate
            ,c.dispvalue 
            order by A.CurrDueAMT DESC
             """)
        
        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()
        
        try:
            with open('currentdue.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow(['SL NAME','ACCT NO','FD NAME','CONTACT','CURRENT DUE','SEGMENT'])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
        except:
               self.errorMsg()
                
        os.startfile('currentdue.csv')

#---------------------------------------------------------------------------------------------------------------------------------
    def ExportSLKPIUpdateCheck(self):

        if self.EntryCampYear.text() > self.datecalKPI.toString('yyyy'):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")
        elif self.EntryCampYear.text() < str(2009):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")

        else:

            if self.EntryCampMonth.text() == '02':
                self.DayofMonth = '28'
                self.ExportNewSLKPI()

            elif self.EntryCampMonth.text() == '04':
                self.DayofMonth = '30'
                self.ExportNewSLKPI()
            elif self.EntryCampMonth.text() == '06':
                self.DayofMonth = '30'
                self.ExportNewSLKPI()
            elif self.EntryCampMonth.text() == '09':
                self.DayofMonth = '30'
                self.ExportNewSLKPI()          
            elif self.EntryCampMonth.text() == '11':
                self.DayofMonth = '30'
                self.ExportNewSLKPI()
            elif self.EntryCampMonth.text() == '01' or self.EntryCampMonth.text() == '03' or self.EntryCampMonth.text() == '05' or self.EntryCampMonth.text() == '07' or \
                 self.EntryCampMonth.text() == '08' or self.EntryCampMonth.text() == '10' or self.EntryCampMonth.text() == '12':
                self.DayofMonth = '31'
                self.ExportNewSLKPI()
            else :
                QtGui.QMessageBox.information(self, 'Invalid Month input',
                         "Sorry, the Month you input is invalid. Please try again.")
#---------------------------------------------------------------------------------------------------------------------------------
    def ExportNewSLKPI(self):

        self.ConnectToSQL()

        try:        
            SQLCommand=(
                """
                    SELECT
                DISTID,
                Name SLName
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID) AS Personal
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) FirstGen
                              --, TotalDPS,
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                     WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                                       AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                     GROUP BY Distid		
                                    HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND TitleID <= '03'
                     AND AppointDate BETWEEN '"""+self.EntryCampMonth.text()+"""/01/"""+self.EntryCampYear.text()+"""' 
                                     AND '"""+self.EntryCampMonth.text()+"""/"""+self.DayofMonth+"""/"""+self.EntryCampYear.text()+"""'
                  ) Appts
               FROM Dealer DL WITH(NOLOCK)
                    WHERE TitleID IN ('04','05')
                   AND (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYear.text()+self.EntryCampMonth.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) !=0

                  order by dl.FirstGen DESC


                 """)

            self.cursor.execute(SQLCommand)
            results=self.cursor.fetchone()
            
            with open('slkpi.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow([
                         'SL ACCOUNT',
                         'SL Name',         
                         'Personal',
                         'FirstGen',
                         'Active',
                         'Appoint',])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
        except:
               self.errorMsg()
            
        os.startfile('slkpi.csv')

#---------------------------------------------------------------------------------------------------------------------------------------------------
    def ExportSLKPIUpdateCheckCritical(self):

        if self.EntryCampYearC.text() > self.datecalKPI.toString('yyyy'):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")
        elif self.EntryCampYearC.text() < str(2009):
            QtGui.QMessageBox.information(self, 'Invalid Year input',
                     "Sorry, the Year you input is invalid. Please try again.")

        else:

            if self.EntryCampMonthC.text() == '02':
                self.DayofMonthC = '28'
                self.ExportSLKPIUpdateClickCritical()

            elif self.EntryCampMonthC.text() == '04':
                self.DayofMonthC = '30'
                self.ExportSLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '06':
                self.DayofMonthC = '30'
                self.ExportSLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '09':
                self.DayofMonthC = '30'
                self.ExportSLKPIUpdateClickCritical()                
            elif self.EntryCampMonthC.text() == '11':
                self.DayofMonthC = '30'
                self.ExportSLKPIUpdateClickCritical()
            elif self.EntryCampMonthC.text() == '01' or self.EntryCampMonthC.text() == '03' or self.EntryCampMonthC.text() == '05' or self.EntryCampMonthC.text() == '07' or \
                 self.EntryCampMonthC.text() == '08' or self.EntryCampMonthC.text() == '10' or self.EntryCampMonthC.text() == '12':
                self.DayofMonthC = '31'
                self.ExportSLKPIUpdateClickCritical()
            else :
                QtGui.QMessageBox.information(self, 'Invalid Month input',
                         "Sorry, the Month you input is invalid. Please try again.")

#----------------------------------------------------------------------------------------------------------------------------------
    def ExportSLKPIUpdateClickCritical(self):
        
        self.ConnectToSQL()
       
        try:
            SQLCommand=(
                """
                    SELECT 
                 Name SLName
                ,(SELECT Name FROM Dealer WITH(NOLOCK)
                   WHERE DistID = DL.ManagerID) UplineName
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID) AS Personal
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) FirstGen
                              --, TotalDPS,
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                     WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                                       AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                     GROUP BY Distid		
                                    HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND TitleID <= '03'
                     AND AppointDate BETWEEN '"""+self.EntryCampMonthC.text()+"""/01/"""+self.EntryCampYearC.text()+"""' 
                                     AND '"""+self.EntryCampMonthC.text()+"""/"""+self.DayofMonthC+"""/"""+self.EntryCampYearC.text()+"""'
                  ) Appts
               FROM Dealer DL WITH(NOLOCK)
                    WHERE TitleID IN ('04','05')
                    AND (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID)< 2000
                    AND ((SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID))+ (SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryCampYearC.text()+self.EntryCampMonthC.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID))>=30000
                                     
                  order by dl.FirstGen DESC


                 """)

            self.cursor.execute(SQLCommand)
            results=self.cursor.fetchone()

            with open('slkpicritical.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow([
                         'SL Name',
                         'UplineName',
                         'Personal',
                         'FirstGen',
                         'Total',
                         'Active',
                         ])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()

        except:
               self.errorMsg()
            
        os.startfile('slkpicritical.csv')

#--------------------------------------------------------------------------------------------------------------------------------------

    def ExportNewSLReportGen(self):


        self.ConnectToSQL()


        SQLCommand=(
            """select
                 A.Distid  
                ,A.Name
                ,(Select Name From Dealer where DistID = A.ManagerID)   SL_Name

                ,A.Street Address
                ,A.Zip
                ,A.MobilePhoneNo
                ,CONVERT(VARCHAR(10),A.TitleDate,10) AppointDate

                from Dealer A with (NOLOCK)
                Inner Join discctdsales B
                on A. DistID=B.DistID, enum c
                where A.TitleDate >='"""+self.EntryDateFrom2SL.text()+"""' and A.TitleDate <='"""+self.EntryDateTo2SL.text()+"""' --- change appdate
                and A.Status in ('A','I')
                and A.TitleID in ('04','05')
                group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.TitleDate                
                
                order by A.TitleDate

             """)
        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()

        try:
            with open('newslcreport.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow([
                         'SL ACCOUNT',
                         'SL NAME',
                         'UPLINE',
                         'ADDRESS',
                         'ZIP',
                         'MOBILE NO.',
                         'APPT DATE',
                         ])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
        except:
               self.errorMsg()
        os.startfile('newslcreport.csv')

#----------------------------------------------------------------------------------------
    def DatePickFromFDALL(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(760,430)
        self.datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateFromFDALL)
     
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateFrom2FDALL.setText(self.datepick.date.toString('yyyy-MM-dd'))          
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()
              
    def showDateFromFDALL(self, date):
        
        self.EntryDateFrom2FDALL.setText(date.toString('yyyy-MM-dd'))
   
#----------------------------------------------------------------------------------------
    def DatePickToFDALL(self):

        self.datepick = QtGui.QMainWindow(self)
        self.datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(1000,430)
        self.datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        self.datepick.cal = QtGui.QCalendarWidget(self)
        self.datepick.cal.setGridVisible(True)
        self.datepick.cal.move(20, 20)
        self.datepick.cal.clicked[QtCore.QDate].connect(self.showDateToFDALL)
        
        self.datepick.date = self.datepick.cal.selectedDate()
        self.EntryDateTo2FDALL.setText(self.datepick.date.toString('yyyy-MM-dd'))        
  
        self.datepick.CentralWidgetC = QtGui.QWidget()
        self.datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        self.datepick.CentralWidgetLayoutC.addStretch()

        self.datepick.CentralWidgetLayoutC.addWidget(self.datepick.cal)
        
        self.datepick.CentralWidgetC.setLayout(self.datepick.CentralWidgetLayoutC)
        self.datepick.setCentralWidget(self.datepick.CentralWidgetC)
        
        self.datepick.show()      

      
    def showDateToFDALL(self, date):
        
        self.EntryDateTo2FDALL.setText(date.toString('yyyy-MM-dd'))

#----------------------------------------------------------------------------------------------------------------------------------------
    def NewFDTrackingALL(self):

        self.app8 = QtGui.QMainWindow(self)
        self.app8.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.app8.setWindowTitle(self.tr('New FD Tracking ALL'))
        self.app8.setFixedSize(850,100)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        self.app8.move((resolution.width() / 2) - (self.frameSize().width() / 1.2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

        self.LabelYearMonth2FDALL = QtGui.QLabel("Month:(YYYYMM)")
        self.EntryYearMonth2FDALL = QtGui.QLineEdit(self)
        self.LabelDateFrom2FDALL = QtGui.QLabel("From:(YYYY-MM-DD)")
        self.EntryDateFrom2FDALL = QtGui.QLineEdit(self)
        self.BttnDateFrom2FDALL = QtGui.QPushButton("...")
        self.LabelDateTo2FDALL = QtGui.QLabel("To:(YYYY-MM-DD)")
        self.EntryDateTo2FDALL = QtGui.QLineEdit(self)
        self.BttnDateTo2FDALL = QtGui.QPushButton("...")
        self.BttnGenerate2FDALL = QtGui.QPushButton("Generate")
        self.BttnGenerate2SLExportFDALL = QtGui.QPushButton("Export")
                             
        #Command that will display the widget's
        self.CentralWidget2FDALL = QtGui.QWidget()
        self.CentralWidgetLayout2FDALL = QtGui.QHBoxLayout()
        self.CentralWidgetLayout2FDALL.addStretch()

        self.CentralWidgetLayout2FDALL.addWidget(self.LabelYearMonth2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.EntryYearMonth2FDALL,1) 
        self.CentralWidgetLayout2FDALL.addWidget(self.LabelDateFrom2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.EntryDateFrom2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.BttnDateFrom2FDALL,1)        
        self.CentralWidgetLayout2FDALL.addWidget(self.LabelDateTo2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.EntryDateTo2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.BttnDateTo2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.BttnGenerate2FDALL,1)
        self.CentralWidgetLayout2FDALL.addWidget(self.BttnGenerate2SLExportFDALL,1)
        
        self.CentralWidget2FDALL.setLayout(self.CentralWidgetLayout2FDALL)
        self.app8.setCentralWidget(self.CentralWidget2FDALL)

        self.calendarFDALL = QtGui.QCalendarWidget(self)
        self.datecalFDALL = self.calendarFDALL.selectedDate()
        
        #default date on month
        calendar = QtGui.QCalendarWidget(self)
        datecal = calendar.selectedDate()

        self.EntryYearMonth2FDALL.setText(datecal.toString('yyyyMM'))

        #Command click date pick buttons
        self.BttnDateFrom2FDALL.clicked.connect(self.DatePickFromFDALL)
        self.BttnDateTo2FDALL.clicked.connect(self.DatePickToFDALL)
        self.BttnGenerate2FDALL.clicked.connect(self.NewFDReportGenFDALL)
        self.BttnGenerate2SLExportFDALL.clicked.connect(self.ExportNewFDALL)
        
                       
        self.app8.show()

#------------------------------------------------------------------------------------------------------------
    def FirstpageNewFDALL(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',12)
        canvas.setFillColor('Red')
        text2 = "*** AVON ANGELES BRANCH  *  New FD TRACKING REPORT  *  As of "+time.strftime("%m/%d/%Y")+" ***" 
        canvas.drawRightString(215*mm, 203*mm, text2)

#-----------------------------------------------------------------------------------------------------------------------------------------          

    def NewFDReportGenFDALL(self):


        self.ConnectToSQL()
        
        try:
            SQLCommand=(
                """select
                    A.managerid		    SL_No
                    ,(Select Name From Dealer where DistID = A.ManagerID)   SL_Name
                    ,A.Distid  
                    ,A.Name
                    ,A.Street Address
                    ,A.Zip
                    ,A.MobilePhoneNo
                    ,CONVERT(VARCHAR(10),A.AppointDate,10) AppointDate
                    ,SUM(B.CTDDiscRateBase) as TOTALDPS
                    ,A.PastDueAmt
                    from Dealer A with (NOLOCK)
                    Inner Join discctdsales B
                    on A. DistID=B.DistID, enum c
                    where B.campaign >= '""" +self.EntryYearMonth2FDALL.text()+"""'--- change campaign
                    and A.AppointDate >='"""+self.EntryDateFrom2FDALL.text()+"""' and A.AppointDate <='"""+self.EntryDateTo2FDALL.text()+"""' --- change appdate
                    and A.Status in ('A','I')
                    and A.TitleID ='01'
                    and B.ProductDiscountGroupID in ('CFT','HERBALCARE','HOMESTYLE','NCFT')
                    and a.starlevel = c.enumvalue and c.tablename ='dealer'
                    and c.columnname ='StarLevel'
                    ---and b.CTDDiscRateBase>='0'
                    ---and A.PastDueAmt>='0'
                    group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.AppointDate,A.PastdueAmt
                    order by SUM(B.CTDDiscRateBase) DESC
                 """)

            self.cursor.execute(SQLCommand)
            self.results=self.cursor.fetchone()

            #PDF format

            self.styleP = getSampleStyleSheet()
            self.styleBH = self.styleP["Normal"]
            self.styleBH.alignment = TA_CENTER
            self.styleBH.fontSize=7


            self.doc = SimpleDocTemplate("sqlnewfdall.pdf", pagesize = letter, rightMargin=10,lefMargin=10, topMargin=45,bottomMargin=12)
            self.doc.pagesize=landscape(letter)
            self.elements=[]
            self.data=[]

            self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                            ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                            ('BOX',(0,0),(-1,-1),0.50,colors.red),
                            ('FONTSIZE',(0,0),(-1,-1),9),
                            ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                            ])

            self.data.append([
                         'SL NAME',
                         'ACCOUNT NO.',
                         'NAME',                     
                         'CONTACT #',
                         'APPT DATE',
                         'TOTAL DPS',
                         'PASTDUE',
                         ])

                        
            #Get Data from SQL to PDF
      
            self.TotaDPSNewFD = 0
            self.CountFD = 0
            Pastdue = 0
        
            while self.results:        
                   
                   
                    self.data.append([str(self.results[1]),
                          str(self.results[2]),
                          str(self.results[3]),
                          str(self.results[6]),                                                                        
                          str(self.results[7]),             
                          str(self.results[8]),
                          str(self.results[9])
                                      ])
                       
                    #compute total DPSO
                    self.TotaDPSNewFD +=float(self.results[8])
                    self.CountFD += 1
                    Pastdue += float(self.results[9])
                                     
                    self.results=self.cursor.fetchone()

            self.data.append(['TOTAL RECRUIT :',
                         self.CountFD,
                         '',                          
                         '',
                         'TOTAL DPS :',
                         round(self.TotaDPSNewFD,2),
                         round(Pastdue,2)   
                         ])

                         
            self.t=Table(self.data,colWidths=(70*mm,30*mm,70*mm,25*mm,20*mm,28*mm,23*mm),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageNewFDALL, onLaterPages=self.addPageNumber)

            self.connection.close()

        except:
               self.errorMsg()
        
        #Open PDF
        os.startfile('sqlnewfdall.pdf')
                
#--------------------------------------------------------------------------------------------------------------------------------------

    def ExportNewFDALL(self):


        self.ConnectToSQL()


        SQLCommand=(
            """select
               -- A.managerid		    SL_No
                (Select Name From Dealer where DistID = A.ManagerID)   SL_Name
                ,A.Distid  
                ,A.Name
                --A.Street Address
                --A.Zip
                ,A.MobilePhoneNo
                ,CONVERT(VARCHAR(10),A.AppointDate,10) AppointDate
                ,SUM(B.CTDDiscRateBase) as TOTALDPS
                ,A.PastDueAmt
                from Dealer A with (NOLOCK)
                Inner Join discctdsales B
                on A. DistID=B.DistID, enum c
                where B.campaign >= '""" +self.EntryYearMonth2FDALL.text()+"""'--- change campaign
                and A.AppointDate >='"""+self.EntryDateFrom2FDALL.text()+"""' and A.AppointDate <='"""+self.EntryDateTo2FDALL.text()+"""' --- change appdate
                and A.Status in ('A','I')
                and A.TitleID ='01'
                and B.ProductDiscountGroupID in ('CFT','HERBALCARE','HOMESTYLE','NCFT')
                and a.starlevel = c.enumvalue and c.tablename ='dealer'
                and c.columnname ='StarLevel'
                ---and b.CTDDiscRateBase>='0'
                ---and A.PastDueAmt>='0'
                group by a.managerid, A.Distid , A.Name ,A.Street, A.Zip, A.MobilePhoneNo,A.AppointDate,A.PastdueAmt
                order by SUM(B.CTDDiscRateBase) DESC
             """)
        
        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()

        try:
            with open('newfdall.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow([
                         'SL NAME',
                         'ACCOUNT NO.',
                         'NAME',                     
                         'CONTACT #',
                         'APPT DATE',
                         'TOTAL DPS',
                         'PASTDUE',
                         ])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
        except:
               self.errorMsg()
            
        os.startfile('newfdall.csv')
#----------------------------------------------------------------------------------------------------------------------------------------
    def NewSLTrackingKPI(self):

        app3 = QtGui.QMainWindow(self)
        app3.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        app3.setWindowTitle(self.tr('New SL Tracking KPI'))        
        app3.setFixedSize(1010,100)

        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        app3.move((resolution.width() / 2) - (self.frameSize().width() / 1),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
        


        LabelYearMonth2 = QtGui.QLabel("Month:(YYYYMM)")
        self.EntryYearMonth2SLNEWSLKPI = QtGui.QLineEdit(self)
        LabelDateFrom2 = QtGui.QLabel("From:(YYYY-MM-DD)")
        self.EntryDateFromSLNEWSLKPI = QtGui.QLineEdit(self)
        BttnDateFrom2 = QtGui.QPushButton("...")
        LabelDateTo2 = QtGui.QLabel("To:(YYYY-MM-DD)")
        self.EntryDateTo2SLNEWSLKPI = QtGui.QLineEdit(self)
        BttnDateTo2 = QtGui.QPushButton("...")
        BttnGenerate2 = QtGui.QPushButton("Generate")
        BttnExport2 = QtGui.QPushButton("Export")
                             
        #Command that will display the widget's
        CentralWidget2 = QtGui.QWidget()
        CentralWidgetLayout2 = QtGui.QHBoxLayout()
        CentralWidgetLayout2.addStretch()


        CentralWidgetLayout2.addWidget(LabelYearMonth2,1)
        CentralWidgetLayout2.addWidget(self.EntryYearMonth2SLNEWSLKPI,1)
        CentralWidgetLayout2.addWidget(LabelDateFrom2,1)
        CentralWidgetLayout2.addWidget(self.EntryDateFromSLNEWSLKPI,1)
        CentralWidgetLayout2.addWidget(BttnDateFrom2,1)        
        CentralWidgetLayout2.addWidget(LabelDateTo2,1)
        CentralWidgetLayout2.addWidget(self.EntryDateTo2SLNEWSLKPI,1)
        CentralWidgetLayout2.addWidget(BttnDateTo2,1)
        CentralWidgetLayout2.addWidget(BttnGenerate2,1)
        CentralWidgetLayout2.addWidget(BttnExport2,1)
        
        
        CentralWidget2.setLayout(CentralWidgetLayout2)
        app3.setCentralWidget(CentralWidget2)

        calendar = QtGui.QCalendarWidget(self)
        datecal = calendar.selectedDate()

        self.EntryYearMonth2SLNEWSLKPI.setText(datecal.toString('yyyyMM'))

        #Command click date pick buttons
        BttnDateFrom2.clicked.connect(self.DatePickFromNEWSLKPI)
        BttnDateTo2.clicked.connect(self.DatePickToNEWSLKPI)
        BttnGenerate2.clicked.connect(self.NewSLKPIUpdateClick)
        BttnExport2.clicked.connect(self.ExportNewSLKPIUpdateClick)
                       
        app3.show()

#----------------------------------------------------------------------------------------
    def DatePickFromNEWSLKPI(self):

        datepick = QtGui.QMainWindow(self)
        datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(760,430)
        datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        datepick.cal = QtGui.QCalendarWidget(self)
        datepick.cal.setGridVisible(True)
        datepick.cal.move(20, 20)
        datepick.cal.clicked[QtCore.QDate].connect(self.showDateFromSLNEWSLKPI)
     
        datepick.date = datepick.cal.selectedDate()
        self.EntryDateFromSLNEWSLKPI.setText(datepick.date.toString('MM/dd/yyyy'))            
        datepick.CentralWidgetC = QtGui.QWidget()
        datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        datepick.CentralWidgetLayoutC.addStretch()

        datepick.CentralWidgetLayoutC.addWidget(datepick.cal)
        
        datepick.CentralWidgetC.setLayout(datepick.CentralWidgetLayoutC)
        datepick.setCentralWidget(datepick.CentralWidgetC)
        
        datepick.show()
              
    def showDateFromSLNEWSLKPI(self, date):
        
        self.EntryDateFromSLNEWSLKPI.setText(date.toString('MM/dd/yyyy'))
   
#----------------------------------------------------------------------------------------
    def DatePickToNEWSLKPI(self):

        datepick = QtGui.QMainWindow(self)
        datepick.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        datepick.setWindowTitle(self.tr('Calendar'))
        #self.datepick.move(1000,430)
        datepick.setFixedSize(320,300)
        
        #position of window
        resolution = QtGui.QDesktopWidget().screenGeometry()
        datepick.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))
   
        datepick.cal = QtGui.QCalendarWidget(self)
        datepick.cal.setGridVisible(True)
        datepick.cal.move(20, 20)
        datepick.cal.clicked[QtCore.QDate].connect(self.showDateToSL)
        
        datepick.date = datepick.cal.selectedDate()
        self.EntryDateTo2SLNEWSLKPI.setText(datepick.date.toString('MM/dd/yyyy'))        
  
        datepick.CentralWidgetC = QtGui.QWidget()
        datepick.CentralWidgetLayoutC = QtGui.QHBoxLayout()
        datepick.CentralWidgetLayoutC.addStretch()

        datepick.CentralWidgetLayoutC.addWidget(datepick.cal)
        
        datepick.CentralWidgetC.setLayout(datepick.CentralWidgetLayoutC)
        datepick.setCentralWidget(datepick.CentralWidgetC)
        
        datepick.show()      

      
    def showDateToSL(self, date):
        
        self.EntryDateTo2SLNEWSLKPI.setText(date.toString('MM/dd/yyyy'))


#---------------------------------------------------------------------
    def NewSLKPIUpdateClick(self):

        
        self.ConnectToSQL()

        try:
            self.SQLCommand=(
                """
                    SELECT
                (SELECT Name FROM Dealer WITH(NOLOCK) WHERE DistID = DL.ManagerID) UplineName  	
                ,Name SLName

                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID = DL.DistID) AS Personal
                ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
                   WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                     AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                     AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                     WHERE ManagerID = DL.DistID)) FirstGen
                              --, TotalDPS,
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                     WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                                       AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                     GROUP BY Distid		
                                    HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
                ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
                   WHERE ManagerID = DL.DistID
                     AND TitleID <= '03'
                     AND AppointDate BETWEEN '"""+self.EntryDateFromSLNEWSLKPI.text()+"""' 
                                     AND '"""+self.EntryDateTo2SLNEWSLKPI.text()+"""'
                  ) Appts
               FROM Dealer DL WITH(NOLOCK)          
                    WHERE TitleID IN ('04','05')
                   AND TITLEDATE BETWEEN '"""+self.EntryDateFromSLNEWSLKPI.text()+"""' 
                                     AND '"""+self.EntryDateTo2SLNEWSLKPI.text()+"""'


                  order by dl.FirstGen DESC


                 """)

            self.cursor.execute(self.SQLCommand)
            self.results=self.cursor.fetchone()

            #PDF format

            self.styleP = getSampleStyleSheet()
            self.styleBH = self.styleP["Normal"]
            self.styleBH.alignment = TA_CENTER
            self.styleBH.fontSize=12


            self.doc = SimpleDocTemplate("sqlKPINEWSLC.pdf", pagesize = letter, rightMargin=3,lefMargin=5, topMargin=60,bottomMargin=12)
            #self.doc.pagesize=landscape(letter)
            self.elements=[]
            self.data=[]

            self.styleT=TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                            ('INNERGRID',(0,0),(-1,-1),0.1,colors.black),
                            ('BOX',(0,0),(-1,-1),0.50,colors.red),
                            ('FONTSIZE',(0,0),(-1,-1),9),
                            ('TEXTFONT',(0,0),(-1,-1),'Arial-Narrow'),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.green),
                            ('TEXTCOLOR',(0,-1),(-1,-1),colors.blue)
                            ])

            self.data.append([
                         '',
                         'UPLINE',
                         'SL Name',         
                         'Personal',
                         'FirstGen',
                         'Total',
                         'Active',
                         'Appoint',
                         ])

                        
            #Get Data from SQL to PDF

          
            self.TotaLSalesSLPersonal = 0
            self.TotalSalesSLFirst = 0
            self.TotalSalesSL = 0
            self.TotalSalesActive = 0
            self.TotalSalesRec = 0
            self.CountSL = 1
            locale.setlocale( locale.LC_ALL, '' ),
        
            while self.results:        

                                       
                    self.data.append([self.CountSL,
                          str(self.results[0]),
                          str(self.results[1]),                     
                                     
                          locale.format("%d",float(self.results[2]),grouping=True),
                          locale.format("%d",float(self.results[3]),grouping=True),
                          locale.format("%d",float(self.results[2])+float(self.results[3]),grouping=True),
                          str(self.results[4]),
                          str(self.results[5])
          
                                      ])
                       
                    #compute total DPS
                    self.CountSL +=1
                    self.TotaLSalesSLPersonal +=int(self.results[2])
                    self.TotalSalesSLFirst +=int(self.results[3])                              
                    self.TotalSalesActive +=int(self.results[4])
                    self.TotalSalesRec +=int(self.results[5])

                    self.results=self.cursor.fetchone()
                    
            #Last Table total
            self.data.append([
                             '',
                             '',
                             '',                         
                             locale.format("%d",float(self.TotaLSalesSLPersonal),grouping=True),
                             locale.format("%d",float(self.TotalSalesSLFirst),grouping=True),
                             locale.format("%d",float((self.TotalSalesSLFirst)+(self.TotaLSalesSLPersonal)),grouping=True),
                             locale.format("%d",float(self.TotalSalesActive),grouping=True),
                             self.TotalSalesRec,
                             ])

                         
            self.t=Table(self.data,colWidths=(6*mm,55*mm,60*mm,18*mm,20*mm,23*mm,14*mm,14*mm,),
                         repeatRows=1,repeatCols=0,hAlign = 'RIGHT')

                
            self.t.setStyle(self.styleT)
            self.elements.append(self.t)      
            self.doc.build(self.elements, onFirstPage=self.FirstpageNEWSLKPI, onLaterPages=self.addPageNumber)

            self.connection.close()

        except:
            
            self.errorMsg()
        
        #Open PDF
        os.startfile('sqlKPINEWSLC.pdf')

#-----------------------------------------------------------------------------
    def FirstpageNEWSLKPI(self,canvas,doc):
        """
        Add the page number
        """
        page_num = canvas.getPageNumber()
        text = "Page %s" % page_num
        canvas.setFont('Helvetica',8)
        canvas.setFillColor('Blue')
        canvas.drawRightString(150*mm, 2*mm, text)

        canvas.setFont('Helvetica',10)
        canvas.setFillColor('Red')
              
        
        text2c = "AVON ANGELES BRANCH"
        canvas.drawRightString(135*mm, 267*mm, text2c)
        
        text3c = "Campaign: "+self.EntryYearMonth2SLNEWSLKPI.text()+" NEW SL KPI REPORT"
        canvas.drawRightString(142*mm, 262*mm, text3c)

        text4c = "Print date : "+time.strftime("%m/%d/%Y") 
        canvas.drawRightString(130*mm, 257*mm, text4c)

        PAGE_WIDTH  = defaultPageSize[0]
        PAGE_HEIGHT = defaultPageSize[1]

#---------------------------------------------------------------------
    def ExportNewSLKPIUpdateClick(self):

        
        self.ConnectToSQL()

        
        SQLCommand=(
            """
                SELECT
            (SELECT Name FROM Dealer WITH(NOLOCK) WHERE DistID = DL.ManagerID) UplineName  	
            ,Name SLName

            ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
               WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                 AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                 AND DistID = DL.DistID) AS Personal
            ,(SELECT isnull(SUM(CtdDiscRateBase),0) FROM DISCCTDSALES WITH(NOLOCK)
               WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                 AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                 AND DistID IN (SELECT DistID FROM Dealer WITH(NOLOCK)
                                 WHERE ManagerID = DL.DistID)) FirstGen
                          --, TotalDPS,
            ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
               WHERE ManagerID = DL.DistID
                 AND DistID IN (SELECT Distid FROM DISCCTDSALES WITH(NOLOCK)
                                 WHERE Campaign = '"""+self.EntryYearMonth2SLNEWSLKPI.text()+"""'
                                   AND ProductDiscountGroupID IN ('CFT','NCFT','HERBALCARE','HOMESTYLE')
                                 GROUP BY Distid		
                                HAVING SUM(CtdDiscRateBase) > 0)) ActiveCnt
            ,(SELECT COUNT(DistID) FROM Dealer WITH(NOLOCK)
               WHERE ManagerID = DL.DistID
                 AND TitleID <= '03'
                 AND AppointDate BETWEEN '"""+self.EntryDateFromSLNEWSLKPI.text()+"""' 
                                 AND '"""+self.EntryDateTo2SLNEWSLKPI.text()+"""'
              ) Appts
           FROM Dealer DL WITH(NOLOCK)          
                WHERE TitleID IN ('04','05')
               AND TITLEDATE BETWEEN '"""+self.EntryDateFromSLNEWSLKPI.text()+"""' 
                                 AND '"""+self.EntryDateTo2SLNEWSLKPI.text()+"""'

              order by dl.FirstGen DESC


             """)

        self.cursor.execute(SQLCommand)
        results=self.cursor.fetchone()

        try:
            with open('newSLKPIall.csv','w',newline='') as csvfile:
                writer = csv.writer(csvfile)

                writer.writerow([   
                         'UPLINE',
                         'SL Name',         
                         'Personal',
                         'FirstGen',
                         'Total',
                         'Active',
                         'Appoint',
                         ])

                while results:
                    writer.writerow(results)
                    results=self.cursor.fetchone()


                self.connection.close()
        except:
               self.errorMsg()
        
            
        os.startfile('newSLKPIall.csv')

#----------------------------------------------------------------------------------------------        
def main():
    app = QtGui.QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
