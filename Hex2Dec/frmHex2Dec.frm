VERSION 5.00
Begin VB.Form frmHex2Dec 
   Caption         =   "    Convert HEX to DEC"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmHex2Dec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDEC 
      Height          =   405
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtHEX 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblDEC 
      Caption         =   "DEC:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblHEX 
      Caption         =   "HEX:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmHex2Dec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'**
'** Program: Hex To Decimal
'**
'** Description: This program will Convert A Hexidecimal to a Decimal
'**
'** Author: Richard Knechtel
'** Email: rknech@pcisys.net
'** Date:   11/17/2001
'**
'** Note: This program and any of the source may not be sold, sub-licensed
'**       or sold for profit without the express written permisson of the
'**       the author (ME). Author holds all exclusive copyrights.
'**       This program is released under the GPL (see below license) and the
'**       Copying.txt file accompanying the source.
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

Private Function HexToDec(HexName) As String

    Dim sTmp, table, Base, sRet, sTmp1, sTmp2

    table = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
    Base = 1
    sRet = 0
    HexName = Replace(UCase(HexName), ".BMP", "")
    sTmp = Len(HexName)
     For sTmp1 = sTmp To 1 Step -1
         sTmp2 = Ascan(table, Mid$(HexName, sTmp1, 1)) + 1
         
        If sTmp2 > 0 Then
            sRet = sRet + (sTmp2 - 1) * Base
            Base = Base * 16
        Else
            sRet = 0
            Exit For
        End If
    Next sTmp1
    HexToDec = sRet

End Function

Private Sub cmdConvert_Click()
txtDEC = HexToDec(Trim(txtHEX.Text))

End Sub

Private Sub cmdExit_Click()

Unload Me
End

End Sub
Public Function Ascan(ArrWhere, WhatValue) As String
    ' Position "WhatValue" in the Array "ArrWhere"
    Dim I11   As Integer
    Dim Ret As Integer
    Ret = -1
    For I11 = 0 To UBound(ArrWhere)
        If ArrWhere(I11) = WhatValue Then
            Ret = I11
            Exit For
        End If
    Next I11
    Ascan = Ret
End Function
