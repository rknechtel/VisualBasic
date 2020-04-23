Attribute VB_Name = "ModFunctions"
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

'*********************************************************************************
'** NOTE:
'** The Original "GenerateDataItemsForTable" Subroutine is
'** "Copyright ©1999 Paragon Corporation   (May 22, 1999)"
'** BUT:
'*********************************************************************************************************************
'** Reworked By Richard Knechtrel 11/17/2001 ***
'** All Changes copyright ME!
'** Original code didn't work - SURPRISE!
'** There table definitions were wrong for their code
'** There code didn't work - missing a database reference - added as a parameter
'** added some items in the "Switch" function.
'** The "GenerateDataItemsForTable" Subroutine in this program is a Derivitive of a
'** publicly released Work.
'*********************************************************************************************************************


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

Sub CenterWindow(frmForm As Form)
'*******************************************************************************
'****  CenterWindow
'****    centers the window over the screen
'*******************************************************************************
    
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    
End Sub

Public Sub GenerateDataItemsForTable(aTableName As String, aDataDictionaryTable As String, CurrentDB As Database)
'**********************************************************************************
'***       Usage: GenerateDataItemsForTable "MyTable", "MyDataDictionaryTable"  ***
'**********************************************************************************
'***-- aTableName -  Table to grab field info from
'***-- aDataDictionaryTable - Data dictionary to put items in most confirm
'***--                        to a particular structure
'**********************************************************************************

Dim tdf As TableDef, fldCur As Field, colTdf As TableDefs
Dim rstDatadict As Recordset
Dim i As Integer, j As Integer, k As Integer
                  
Set rstDatadict = CurrentDB.OpenRecordset(aDataDictionaryTable)
                  
                  
Set colTdf = CurrentDB.TableDefs
                  
Set tdf = colTdf(aTableName)

'*** Put data items in table
For i = 0 To tdf.Fields.Count - 1
                  
    Set fldCur = tdf.Fields(i)
                              
    rstDatadict.AddNew
    rstDatadict![Table] = tdf.Name
    rstDatadict![Field] = "[" & tdf.Name & "]" & ".[" & fldCur.Name & "]"
    rstDatadict![Display] = fldCur.Name
                                  
    '*** --SQL encapsulator for a date type 8 is # and string type 10 is single quote
    '*** --all remaining numbers are nothing
    rstDatadict![encapsulator] = Switch(fldCur.Type = 10, "'", fldCur.Type = 8, "#", True, "")
    
    If rstDatadict!encapsulator = "" Then
        rstDatadict!encapsulator = " "
    End If

'** Debug Code
' Debug.Print fldCur.Name; fldCur.Type

    rstDatadict!DataType = Switch(fldCur.Type = 1, "True/False", fldCur.Type = 3, "Integer", _
                           fldCur.Type = 4, "Long Integer", fldCur.Type = 10, "Text", _
                           fldCur.Type = 8, "Date", True, "Number")

                                                      
    For j = 0 To tdf.Fields(i).Properties.Count - 1
        If fldCur.Properties(j).Name = "Description" Then
           rstDatadict![Description] = fldCur.Properties(j).Value
        End If

        If fldCur.Properties(j).Name = "Caption" Then
           rstDatadict![Display] = fldCur.Properties(j).Value
        End If

        If fldCur.Properties(j).Name = "Rowsource" Then
           rstDatadict![LookupSQL] = fldCur.Properties(j).Value
        End If
    Next j

    rstDatadict.Update
                  
Next i
                      
End Sub



