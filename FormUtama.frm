VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SK-Arrays-To-Columns-in-Text-Box__VB6"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib _
"user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

Private Const EM_SETTABSTOPS = &HCB
Private Sub Form_Load()
    Dim ColumnArray() As String, Row As Integer, Col As Integer
    ReDim ColumnArray(6, 2) 'for 6 lines and 2 columns
                            'Change based on your needs
                    
    'Add Data to Array
    For Row = 0 To UBound(ColumnArray, 1)
        For Col = 0 To UBound(ColumnArray, 2)
            ColumnArray(Row, Col) = Row & Col
        Next
        Text1.Text = Text1.Text + vbCrLf 'New Line
    Next
    'then clear the textbox and Print the array
    Text1.Text = ""
    'Set tab stops to ensure data is evenly spaced.
    'values you pass depend on the length of the data
    
    'This is not really necessary in this example
    'because length of data in each column is the same
    'but may be necessary in other cases
    
    SetTBTabStops Text1, 40, 80, 120
    
    For Row = 0 To UBound(ColumnArray, 1)
        For Col = 0 To UBound(ColumnArray, 2)
            Text1.Text = Text1.Text & ColumnArray(Row, Col) & vbTab
        Next
        Text1.Text = Text1.Text + vbCrLf 'New Line
    Next
End Sub

Public Function SetTBTabStops(TB As Object, ParamArray TabStops()) As Boolean

Dim alTabStops() As Long
Dim lCtr As Long
Dim lColumns As Long
Dim lRet As Long

On Error GoTo errorhandler:

ReDim alTabStops(UBound(TabStops)) As Long

For lCtr = 0 To UBound(TabStops)
    alTabStops(lCtr) = TabStops(lCtr)
Next

lColumns = UBound(alTabStops) + 1


lRet = SendMessage(TB.hwnd, EM_SETTABSTOPS, lColumns, alTabStops(0))

SetTBTabStops = (lRet = 0)
Exit Function

errorhandler:
SetTBTabStops = False

End Function

