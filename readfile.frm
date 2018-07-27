VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6660
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   2940
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "讀檔"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fileArray() As String
Dim tmpline As String
Dim sum As Variant
Open App.Path & "\hw3.txt" For Input As #1
Do While Not EOF(1)
          Line Input #1, tmpline
          'List1.AddItem tmpline
          fileArray = Split(tmpline, ",")
          sum = CInt(fileArray(2)) * 4800 + CInt(fileArray(3)) * 1600
          List1.AddItem fileArray(0) + " " + fileArray(1) + ":" + CStr(sum) + "元"
Loop
Close #1

End Sub



