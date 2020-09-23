VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Hospital Management"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10425
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu main 
      Caption         =   "&Main"
      Begin VB.Menu meddets 
         Caption         =   "&Medicinal Details"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu exdets 
         Caption         =   "&Exam Details"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu docdetails 
         Caption         =   "&Add Doctors"
      End
      Begin VB.Menu mnu14 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_pat 
      Caption         =   "&Patients"
      Begin VB.Menu mnuinp 
         Caption         =   "InPatients"
         Begin VB.Menu mnuaddinp 
            Caption         =   "Add &In Patients"
         End
         Begin VB.Menu mnu3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuipmeds 
            Caption         =   "Add &Medical Details"
         End
         Begin VB.Menu mnu4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuipexms 
            Caption         =   "Add &Exam Details"
         End
         Begin VB.Menu mnu88 
            Caption         =   "-"
         End
         Begin VB.Menu mnudischrge 
            Caption         =   "Discharge"
         End
         Begin VB.Menu mnu99 
            Caption         =   "-"
         End
         Begin VB.Menu mnuadmlst 
            Caption         =   "Admission List"
         End
      End
      Begin VB.Menu outpat 
         Caption         =   "&Out Patients"
         Begin VB.Menu mnuaddout 
            Caption         =   "Add &Out Patients"
         End
         Begin VB.Menu mnu5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuopmeds 
            Caption         =   "Add &Medical Details"
         End
         Begin VB.Menu mnu6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuopexm 
            Caption         =   "Add &Exam Details"
         End
      End
   End
   Begin VB.Menu BllPaymnts 
      Caption         =   "&Bill Payments"
      Begin VB.Menu inpbillpay 
         Caption         =   "&In Patient Bill Payments"
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu opbillPay 
         Caption         =   "&Out Patient Bill Payments"
      End
   End
   Begin VB.Menu reps 
      Caption         =   "&Reports"
      Begin VB.Menu inp 
         Caption         =   "In Patients"
         Begin VB.Menu mnubillrep 
            Caption         =   "&Bill Report"
         End
         Begin VB.Menu mnu8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuipdaily 
            Caption         =   "&Daily Report"
         End
         Begin VB.Menu mnu9 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuipmonrep 
            Caption         =   "&Monthly Report"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu33 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuyrlyiprep 
            Caption         =   "&Yearly Report"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu575 
         Caption         =   "-"
      End
      Begin VB.Menu op 
         Caption         =   "&Out Patients"
         Begin VB.Menu mnuopbllrep 
            Caption         =   "&Bill Report"
         End
         Begin VB.Menu mnu11 
            Caption         =   "-"
         End
         Begin VB.Menu mnudaiioprep 
            Caption         =   "&Daily Report"
         End
         Begin VB.Menu mnu13 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnumonoprep 
            Caption         =   "&Monthly Report"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu12 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuyrlyoprep 
            Caption         =   "&Yearly Report"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub docdetails_Click()
frmdoctors.Show
End Sub
Private Sub exdets_Click()
frmlabtest.Show
End Sub
Private Sub inpbillpay_Click()
frmIPBillPayments.Show
End Sub
Private Sub inpbillrep_Click()
frmIPBillPaymentReport.Show
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Confirm to Quit Application ?", vbQuestion + vbYesNo) = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub
Private Sub mnuAboutMe_Click()
frmMain.Show
End Sub
Private Sub mnuAddMore_Click()
frmAddMore.Show
End Sub
Private Sub mnuAddMoreCustomer_Click()
frmAddMoreToBill.Show
End Sub

Private Sub mnuAllBillPayments_Click()
frmCustomerAllBillPayments.Show
End Sub
Private Sub mnuBillNo_Click()
frmCustomerBillNo.Show
End Sub
Private Sub mnuBillUpdate_Click()
frmUpdateBill.Show
End Sub

Private Sub meddets_Click()
frmMedicine.Show
End Sub

Private Sub mnuabtus_Click()
frmAboutUs.Show
End Sub
Private Sub mnuaddinp_Click()
frminpat.Show
frminpat.WindowState = 2
End Sub
Private Sub mnuaddout_Click()
frmopdet.Show
frmopdet.WindowState = 2
End Sub

Private Sub mnuadmlst_Click()
frmadmlst.Show
End Sub
Private Sub mnubillrep_Click()
frmIPBillPaymentReport.Show
End Sub
Private Sub mnuCalc_Click()
Shell "c:\windows\calc.exe", vbNormalFocus
End Sub
Private Sub mnuChangePassword_Click()
frmPasswordChange.Show
End Sub
Private Sub mnuCompInfo_Click()
frmManufacture.Show
End Sub
Private Sub mnuCurrentStock_Click()
rptCurrentStockReport.Show
End Sub
Private Sub mnuCustInfo_Click()
frmCustomerInfo.Show
End Sub
Private Sub mnuCustomerBill_Click()
frmCustomerBillReport.Show
End Sub
Private Sub mnuCustomerBillPayments_Click()
frmBillPayments.Show
End Sub
Private Sub mnuCustomerStockReturns_Click()
frmCustStockReturns.Show
End Sub

Private Sub mnuCustPurRep_Click()
frmCustomerPurchaseReport.Show
End Sub
Private Sub mnuDailyPayments_Click()
frmDailyPayments.Show
End Sub
Private Sub mnuDistInfo_Click()
frmDistDetails.Show
End Sub
Private Sub mnuDueReports_Click()
frmDueReport.Show
End Sub
Private Sub mnuExcludeTaxValue_Click()
frmCustomerBillTax.Show
End Sub
Private Sub mnuExit_Click()
If MsgBox("Confirm to Quit Application ?", vbQuestion + vbYesNo) = vbYes Then
    End
End If
End Sub
Private Sub mnuExpiredStoc_Click()
frmExpiredStock.Show
End Sub
Private Sub mnuExpiredStock_Click()
frmExpiredStockReport.Show
End Sub
Private Sub mnuIndividualCustomerPayments_Click()
frmCustomerSingleBillPayments.Show
End Sub
Private Sub mnuInvoiceNumber_Click()
frmInvoiceNo.Show
End Sub
Private Sub mnuInvoicePayments_Click()
frmDistPayments.Show
End Sub
Private Sub mnuInvoicePaymentsReports_Click()
frmInvoicePayments.Show
End Sub
Private Sub mnuInvoiceReport_Click()
frmInvoiceReport.Show
End Sub
Private Sub mnuMedicineInfo_Click()
frmMedicine.Show
End Sub
Private Sub mnuMonthlyPayments_Click()
frmMonthlyPayments.Show
End Sub
Private Sub mnuNarration_Click()
frmNarration.Show
End Sub
Private Sub mnuNewBillEntry_Click()
frmStockBilling.Show
End Sub
Private Sub mnuNewInvoiceEntry_Click()
frmStockEntry.Show
End Sub

Private Sub mnudischrge_Click()
frmdischarge.Show
End Sub
Private Sub mnuipdaily_Click()
frmIPDailyPayments.Show
End Sub
Private Sub mnuipexms_Click()
frmipexm.Show
frmipexm.WindowState = 2
End Sub
Private Sub mnuipmeds_Click()
frmipmed.Show
frmipmed.WindowState = 2
End Sub
Private Sub mnuNotepad_Click()
Shell "c:\windows\notepad.exe", vbNormalFocus
End Sub
Private Sub mnuPackingName_Click()
frmPacking.Show
End Sub
Private Sub mnuSalesReport_Click()
frmSales.Show
End Sub
Private Sub mnuStockReturns_Click()
frmStockReturns.Show
End Sub
Private Sub mnuUpdateInvoiceEntry_Click()
frmUpdateInvoice.Show
End Sub
Private Sub mnuopbllrep_Click()
frmOPBillPayments.Show
End Sub
Private Sub mnuopexm_Click()
frmopexm.Show
frmopexm.WindowState = 2
End Sub
Private Sub mnuopmeds_Click()
frmopmeds.Show
frmopmeds.WindowState = 2
End Sub
Private Sub opbillPay_Click()
frmOPBillPayments.Show
End Sub
Private Sub opbillreps_Click()
frmOPBillPaymentReport.Show
End Sub
