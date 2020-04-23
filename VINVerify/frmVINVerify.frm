VERSION 5.00
Begin VB.Form frmVINVerify 
   Caption         =   "  Vehicle VIN Verification"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   Icon            =   "frmVINVerify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "&Verify VIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtVIN 
      Height          =   285
      Left            =   1800
      MaxLength       =   17
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblVIN 
      Caption         =   "VIN Number: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmVINVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnGoodVIN As Boolean

Private Sub cmdVerify_Click()

'** Verify that the VIN # entered is Valid
'** Richard Knechtel 11/01/2000
blnGoodVIN = VINCD(Trim(UCase(txtVIN.Text)))

If blnGoodVIN = False Then
    MsgBox "Vehicle VIN Number entered is not valid -" & vbCrLf & " Check Digit is Invalid."
    txtVIN.SetFocus
    txtVIN.SelStart = 0
    txtVIN.SelLength = Len(txtVIN.Text)
Else
    MsgBox "Vehicle VIN Number entered is valid."
End If


End Sub

Private Sub cmdExit_Click()

End

End Sub
