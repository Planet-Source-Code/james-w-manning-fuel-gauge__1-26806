VERSION 5.00
Object = "*\AFuelGauge.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin FuelProgress.FuelGauge FuelGauge1 
      Height          =   1515
      Left            =   1530
      TabIndex        =   0
      Top             =   480
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   2672
   End
   Begin VB.Timer Timer1 
      Interval        =   125
      Left            =   1500
      Top             =   2580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
    i = 0
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    If i = FuelGauge1.Max Then
        Timer1.Enabled = False
        Exit Sub
    End If
    FuelGauge1.Value = i
End Sub
