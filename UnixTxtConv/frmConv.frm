VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConv 
   Caption         =   "  Unix/Amiga<-->PC Text Converter"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "frmConv.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmConv.frx":0442
   ScaleHeight     =   7185
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConvert 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Convert to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
      Begin VB.OptionButton optPC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optUnix 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unix/Amiga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Convert It!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00FFC0C0&
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Label lblDrives 
      BackStyle       =   0  'Transparent
      Caption         =   "Select File To Convert:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'**
'** Program: Unix/Amiga Text File Converter
'**
'** Description: This program will read a(n) Unix/Amiga text file and convert
'**              it to PC format or read a pc file and convert it to
'**              Unix/Amiga text file format.
'**
'** Author: Richard Knechtel
'** Email: krs3@qwest.net
'** Date:   6/20/2001
'**
'** Note: This program and any of the source may not be sold, sub-licensed
'**       or sold for profit without the express written permisson of the
'**       the author (ME). Author holds all exclusive copyrights.
'**
'*********************************************************************************
'*********  LICENSE  *************************************/
'/*                                                      */
'/* This program is free software; you can redistribute  */
'/* it and/or modify it under the terms of the GNU       */
'/* General Public License as published by  the Free     */
'/* Software Foundation; either version 2 of the License,*/
'/* or (at your option) any later version.               */
'/*                                                      */
'/* This program is distributed in the hope that it will */
'/* be useful, but WITHOUT ANY WARRANTY; without even    */
'/* the implied warranty of MERCHANTABILITY or FITNESS   */
'/* FOR A PARTICULAR PURPOSE.  See the GNU General       */
'/* Public License for more details.                     */
'/*                                                      */
'/* You should have received a copy of the GNU General   */
'/* Public License along with this program; if not,      */
'/* write to the                                         */
'/* Free Software                                        */
'/* Foundation, Inc., 675 Mass Ave,                      */
'/* Cambridge, MA 02139, USA.                            */
'/*                                                      */
'/********************************************************/

Option Explicit
Dim ErrString As String
Dim Done As Boolean

Private Sub cmdSelect_Click()

    cd.DialogTitle = "Open Unix/Amiga Text File"
    cd.Filter = "All Files| *.*"
    cd.InitDir = "C:\"
    cd.ShowOpen
    txtFile.Text = cd.FileName

End Sub

Private Sub cmdConvert_Click()

    Done = ReadTheFile(Trim(txtFile.Text), ErrString)
    
    lblMsg.Visible = True
    If Done = True Then
        If optUnix.Value = True Then
            lblMsg.Caption = "Conversion of " & Trim(txtFile.Text) & " to Unix/Amiga complete."
        End If
        If optPC.Value = True Then
            lblMsg.Caption = "Conversion of " & Trim(txtFile.Text) & " to PC complete."
        End If
    Else
        lblMsg.Caption = "There was an error converting " & Trim(txtFile.Text) & "."
    End If
    
    lblMsg.Refresh

End Sub

Private Sub mnuAbout_Click()
    
    Load frmAbout
    frmAbout.Show

End Sub

Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

