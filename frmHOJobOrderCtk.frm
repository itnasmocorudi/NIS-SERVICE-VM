VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmHOJobOrderCtk 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cetak Job Order"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
End
Attribute VB_Name = "frmHOJobOrderCtk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New HOJobOrder

Private Sub Form_Load()
Dim oRs As Object
Screen.MousePointer = vbHourglass

   sSQL = "SELECT * FROM jobOrder WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY JOB, FS_ITEM"
   Set oRs = objO.FetchData(sSQL)
   Report.Database.SetDataSource oRs

    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Zoom 100
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
    asxBars1.Width = CRViewer1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub
