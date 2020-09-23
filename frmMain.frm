VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MsgBox Editor"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrieview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Code it!"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbStyles 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   "OK"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Code:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Style:"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code As String
Dim Style As String

Private Sub cmdPrieview_Click()
Select Case cmbStyles.Text
    Case "OK"
        MsgBox txtText.Text, vbOKOnly, txtTitle.Text
    Case "OK/Cancel"
        MsgBox txtText.Text, vbOKCancel, txtTitle.Text
    Case "Critical"
        MsgBox txtText.Text, vbCritical, txtTitle.Text
    Case "Abort/Retry/Ignore"
        MsgBox txtText.Text, vbAbortRetryIgnore, txtTitle.Text
    Case "Exclamation"
        MsgBox txtText.Text, vbExclamation, txtTitle.Text
    Case "Information"
        MsgBox txtText.Text, vbInformation, txtTitle.Text
    Case "Question"
        MsgBox txtText.Text, vbQuestion, txtTitle.Text
    Case "Yes/No"
        MsgBox txtText.Text, vbYesNo, txtTitle.Text
    Case "Yes/No/Cancel"
        MsgBox txtText.Text, vbYesNoCancel, txtTitle.Text
End Select
End Sub

Private Sub Command1_Click()
Select Case cmbStyles.Text
    Case "OK"
        Style = "vbOKOnly"
    Case "OK/Cancel"
        Style = "vbOKCancel"
    Case "Critical"
        Style = "vbCritical"
    Case "Abort/Retry/Ignore"
        Style = "vbAbortRetryIgnore"
    Case "Exclamation"
        Style = "vbExclamation"
    Case "Information"
        Style = "vbInformation"
    Case "Question"
        Style = "vbQuestion"
    Case "Yes/No"
        Style = "vbYesNo"
    Case "Yes/No/Cancel"
        Style = "vbYesNoCancel"
End Select
Code = "MsgBox """ & txtText.Text & """, " & Style & ", """ & txtTitle.Text & """"
txtCode.Text = Code
End Sub

Private Sub Form_Load()
cmbStyles.AddItem "OK"
cmbStyles.AddItem "OK/Cancel"
cmbStyles.AddItem "Critical"
cmbStyles.AddItem "Abort/Retry/Ignore"
cmbStyles.AddItem "Exclamation"
cmbStyles.AddItem "Information"
cmbStyles.AddItem "Question"
cmbStyles.AddItem "Yes/No"
cmbStyles.AddItem "Yes/No/Cancel"
End Sub
