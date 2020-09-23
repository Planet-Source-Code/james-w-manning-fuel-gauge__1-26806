VERSION 5.00
Begin VB.UserControl FuelGauge 
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   ScaleHeight     =   1515
   ScaleWidth      =   1605
   ToolboxBitmap   =   "FuelGauge.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   1260
      Width           =   1545
   End
   Begin VB.Line Needle 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1350
      X2              =   810
      Y1              =   690
      Y2              =   990
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   0
      Picture         =   "FuelGauge.ctx":0312
      Top             =   0
      Width           =   1605
   End
End
Attribute VB_Name = "FuelGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_ShowValue = 0
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
'Property Variables:
Dim m_ShowValue As Boolean
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_Value As Integer
Const Pi = 3.14159

Private Sub UserControl_Initialize()
    Width = Image1.Width
    Height = Image1.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
    Needle.X1 = 1350
    Needle.Y1 = 690
    Needle.X2 = Needle.X1 + Sin(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Needle.Y2 = Needle.Y1 - Cos(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Label1.Visible = IIf(m_ShowValue = True, True, False)
    Label1.Caption = Trim(Str(Round(m_Value / (m_Max - m_Min) * 100))) & "%"
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    Needle.X1 = 1350
    Needle.Y1 = 690
    Needle.X2 = Needle.X1 + Sin(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Needle.Y2 = Needle.Y1 - Cos(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Label1.Visible = IIf(m_ShowValue = True, True, False)
    Label1.Caption = Trim(Str(Round(m_Value / (m_Max - m_Min) * 100))) & "%"
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    Needle.X1 = 1350
    Needle.Y1 = 690
    Needle.X2 = Needle.X1 + Sin(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Needle.Y2 = Needle.Y1 - Cos(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Label1.Visible = IIf(m_ShowValue = True, True, False)
    Label1.Caption = Trim(Str(Round(m_Value / (m_Max - m_Min) * 100))) & "%"
    PropertyChanged "Value"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_ShowValue = m_def_ShowValue
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_ShowValue = PropBag.ReadProperty("ShowValue", m_def_ShowValue)
End Sub

Private Sub UserControl_Resize()
    Width = Image1.Width
    Height = Image1.Height
End Sub

Private Sub UserControl_Show()
    Needle.X1 = 1350
    Needle.Y1 = 690
    Needle.X2 = Needle.X1 + Sin(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Needle.Y2 = Needle.Y1 - Cos(Pi * (Round(m_Value / (m_Max - m_Min) * 76) - 119) / 180) * 600
    Label1.Visible = IIf(m_ShowValue = True, True, False)
    Label1.Caption = Trim(Str(Round(m_Value / (m_Max - m_Min) * 100))) & "%"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ShowValue", m_ShowValue, m_def_ShowValue)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowValue() As Boolean
    ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal New_ShowValue As Boolean)
    m_ShowValue = New_ShowValue
    PropertyChanged "ShowValue"
End Property

