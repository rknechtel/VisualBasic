VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBuildDD 
   Caption         =   " Build Data Dictionary"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   Icon            =   "frmBuildDD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuildDD 
      Caption         =   "&Build Data Dictionary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtDBName 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblMsg 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   8775
   End
   Begin VB.Label lblDBName 
      Caption         =   "Database Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Build Data Dictionary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "This Program Builds A Data Dictionary from an Access97 DataBase"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
End
Attribute VB_Name = "frmBuildDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'**
'** Program: Build Data Dictionary
'**
'** Description: This program will read an Access Database and buld a
'**              datadictionary table for it. If the Data Dictionary table does not
'**              Exist, it will create the table then create the data dictionry
'**              table.
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

Dim db As Database '** Access97 Database
Dim x As Integer '** Recordset index
Dim DDFound As Boolean '** Data Dictionary Found

Private Sub cmdBuildDD_Click()
    
Dim td As TableDef
Dim f As Field

    
    '** Set conections to Database and Data Dictionary
    Set db = OpenDatabase(Trim(txtDBName.Text))
    
    '** Initialize Data Dictionary Found check variable to False
    DDFound = False
    
    '** See if Data Dictionary Table exists
    For x = 1 To db.TableDefs.Count - 1
        If db.TableDefs(x).Name = "DataDict" Then
            DDFound = True
            Exit For
        End If
    Next x
    
    '** Create Data Dictionary Table if not found
    If DDFound = False Then
    
        Set td = New TableDef
        Set f = td.CreateField("Table", dbText, 200)
        td.Fields.Append f
        
        Set f = Nothing
        Set f = td.CreateField("Field", dbText, 200)
        td.Fields.Append f

        Set f = Nothing
        Set f = td.CreateField("DataType", dbText, 30)
        td.Fields.Append f
        
        Set f = Nothing
        Set f = td.CreateField("Display", dbText, 200)
        td.Fields.Append f
        
        Set f = Nothing
        Set f = td.CreateField("Encapsulator", dbText, 6)
        td.Fields.Append f
        
        Set f = Nothing
        Set f = td.CreateField("Description", dbMemo, 200)
        td.Fields.Append f
        
        Set f = Nothing
        Set f = td.CreateField("LookupSQL", dbText, 255)
        td.Fields.Append f
        
        td.Name = "DataDict"
        db.TableDefs.Append td
                
        '** Cleanup
        Set f = Nothing
        Set td = Nothing
        
    End If
    
    '** Build Data Dictionary for the tables
    For x = 0 To db.TableDefs.Count - 1
        '** We don't want the Data Dictionary Table and the MS system tables
        If db.TableDefs(x).Name <> "DataDict" And Left(db.TableDefs(x).Name, 4) <> "MSys" Then
            GenerateDataItemsForTable db.TableDefs(x).Name, "DataDict", db
        End If
    Next x
        
'    GenerateDataItemsForTable "Employee", "DataDict", db
'    GenerateDataItemsForTable "CheckListDetail", "DataDict", db
'    GenerateDataItemsForTable "CheckListHistory", "DataDict", db
'    GenerateDataItemsForTable "TrailerFleet", "DataDict", db
'    GenerateDataItemsForTable "TruckFleet", "DataDict", db
'    GenerateDataItemsForTable "Maintenance", "DataDict", db
'    GenerateDataItemsForTable "OutofService", "DataDict", db
    
    '** Cleanup
    Set db = Nothing
    
    '** Display Done Message
    lblMsg.Caption = "Finished Building Data Dictionary for file: " + Trim(txtDBName.Text)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdSelect_Click()

    CD.DialogTitle = "Open Access97 Database"
    CD.Filter = "Access Database| *.mdb"
    CD.InitDir = "C:\"
    CD.ShowOpen
    txtDBName.Text = CD.FileName
    
End Sub

Private Sub Form_Load()

    CenterWindow frmBuildDD
    
End Sub
