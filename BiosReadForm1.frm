VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bios Data"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "BiosReadForm1.frx":0000
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'øº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`ºøº`º
'(oigres P) Email: oigres@postmaster.co.uk
'Get the Bios number and serial details?
'Code adapted from ww.freeVBcode.com by Arkadiy Olovyannikov
'Functions in visual basic virtual machine (runtime dlls)
Private Declare Sub GetMem1 Lib "msvbvm50.dll" (ByVal _
   MemAddress As Long, var As Byte)

'You can read Integer (2 bytes), Long and LongInteger variables
 'using GetMem2, GetMem4 and GetMem8 functions
'Private Declare Sub GetMem2 Lib "msvbvm50.dll" (ByVal _
' MemAddress As Long, var As Integer)
'Private Declare Sub GetMem4 Lib "msvbvm50.dll" (ByVal _
' MemAddress As Long, var As Long)
'API has LongInteger var type 8 bytes long (FileTime is a _
' sample)
'----------------------
Private Function GetBIOSDate() As String
  Dim p As Byte, MemAddr As Long, sBios As String
  Dim i As Integer
  'start of bios serial number ?&HFE0C0
  MemAddr = &HFE000
  For i = 0 To 331
      Call GetMem1(MemAddr + i, p)
      'get printable characters
      If p > 31 And p <= 128 Then
      sBios = sBios & Chr$(p)
    End If
  Next i
  GetBIOSDate = sBios
End Function
'Using
'Text1.Text = GetBiosDate

'Private Declare Sub GetMem8 Lib "msvbvm50.dll" (ByVal _
' MemAddress As Long, var As LongInteger)

'You can also write data derectly into memory using the same
'PutMem1 - PutMem8 functions


Private Sub Form_Load()
Text1.Text = GetBIOSDate
End Sub
