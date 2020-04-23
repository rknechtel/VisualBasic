VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C000C0&
   Caption         =   "About"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00FFC0C0&
      Caption         =   "                   Unix/Amiga Text File Converter "
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub Form_Load()

lblAbout.Caption = "                  Unix/Amiga Text File Converter " & vbCrLf & _
                   "                   Author: Richard Knechtel " & vbCrLf & _
                   "                   Date Coded: 6/20/2001 " & vbCrLf & _
                   "                   Email: rknech@pcisys.net " & vbCrLf & _
                   "                   Description: This program will read a(n) " & vbCrLf & _
                   "                                         Unix/Amiga<-->PC text file and " & vbCrLf & _
                   "                                         convert it to PC format or read " & vbCrLf & _
                   "                                         a PC file and convert it to " & vbCrLf & _
                   "                                         Unix/Amiga text file format. " & vbCrLf

End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub
