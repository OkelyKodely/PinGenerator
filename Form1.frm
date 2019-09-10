VERSION 5.00
Begin VB.Form PINGEN 
   Caption         =   "PIN Generator"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Length"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Prefix"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount per PIN:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of PIN:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "PINGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents RS As Recordset
Attribute RS.VB_VarHelpID = -1
Public WithEvents RS1 As Recordset
Attribute RS1.VB_VarHelpID = -1
Public valid As Boolean
Public db As Connection

Private Sub Command1_Click()
   valid = False
   ValidNumber Text1(0)
   If valid = True Then
      valid = False
      ValidNumber Text1(1)
   End If
   If valid = True Then
      valid = False
      ValidNumber Text1(2)
   End If
   If valid = True Then
      valid = False
      ValidNumber Text1(3)
   End If
   If IsNumeric(Text1(3)) Then
      If Text1(3) <> 12 Then
         Text1(3).SetFocus
         Exit Sub
      End If
   End If
   If valid = True Then
      GenPIN
   End If
End Sub

Sub ValidNumber(txtbox As TextBox)
   If Not IsNumeric(txtbox) Then txtbox = 0

   If IsNumeric(txtbox) And txtbox > 0 Then
      valid = True
   Else
      txtbox.SetFocus
      Exit Sub
   End If

   If IsNumeric(txtbox) And txtbox > 0 Then
      valid = True
   Else
      txtbox.SetFocus
      Exit Sub
   End If
End Sub

Sub GenPIN()
   Randomize Timer
   sqlstr = "delete from temppin"
   db.Execute sqlstr
   Dim tmppin As Double
   Dim mPin As String
   Dim rpin As String
   For i = 1 To Val(Text1(0).Text)
      tmppin = Rnd(100)
      tmppin = Str(tmppin)
      mPin = ""
      For n = 1 To Len(tmppin)
         If Mid(tmppin, n, 1) <> "." Then
            mPin = mPin + Mid(tmppin, n, 1)
         End If
      Next n
      tmppin = Rnd(100)
      tmppin = Str(tmppin)
      For n = 1 To Len(tmppin)
         If Mid(tmppin, n, 1) <> "." Then
            mPin = mPin + Mid(tmppin, n, 1)
         End If
      Next n
      rpin = Str(Text1(2).Text) & mPin
      rpin = Left(rpin, Text1(3))
      sqlstr = "select * from master where pin='" & rpin & "'"
      Set RS = New Recordset
      RS.Open sqlstr, db, adOpenStatic, adLockOptimistic

      If RS.RecordCount > 0 Then
         MsgBox "Invalid pin"
         End
      End If

      sqlstr = "select * from temppin where pin='" & rpin & "'"
      Set RS1 = New Recordset
      RS1.Open sqlstr, db, adOpenStatic, adLockOptimistic

      If RS1.RecordCount > 0 Then
         MsgBox "Invalid pin"
         End
      Else
         With RS1
            .AddNew
            .Fields("Pin") = rpin
            .Fields("amount") = Text1(1)
            .Update
         End With
      End If
      Label3.Caption = "Generating... " & i & " of " & Val(Text1(0).Text)

      DoEvents
   Next

   Set RS1 = New Recordset
   RS1.Open "select Serial,Pin, Amount from temppin", db, adOpenStatic, adLockOptimistic

   If RS1.RecordCount > 0 Then
      RS1.MoveFirst
      Do Until RS1.EOF
         With RS
            .AddNew
            .Fields("PIN") = RS1.Fields("PIN")
            .Fields("Amount") = RS1.Fields("Amount")
            .Update
         End With
         Label3.Caption = "Inserting... " & RS.AbsolutePosition & " of " & Val(Text1(0).Text)
         RS1.MoveNext
      Loop
   End If
End Sub



Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Set db = New Connection
   db.CursorLocation = adUseClient
   db.Open "PROVIDER=MSDASQL;dsn=PIN;uid=;pwd=;"
   Set RS = New Recordset
   Set RS1 = New Recordset
   RS.Open "select Serial,Pin,Amount from Master", db, adOpenStatic, adLockOptimistic
   RS1.Open "select Serial,Pin, Amount from temppin", db, adOpenStatic, adLockOptimistic
End Sub

