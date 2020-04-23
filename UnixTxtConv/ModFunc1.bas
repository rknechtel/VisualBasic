Attribute VB_Name = "ModFunc1"
'*********************************************************************************
'** Note:
'**      The "ReadTheFile" & "Lf2Crlf" functions I got off of the web - I do not know who the
'**      original author was so if you are the one - this is the "credits"
'**      section. - Credit for these functions go to the original author
'**      whomever you are and where ever you are! Their copyright is fully
'**      retained by them.
'**      I just made some modifications to their code. :^)
'**
'** Richard Knechtel 6/20/2001
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

Public Function Lf2Crlf(InputString As String, StringLength As Integer) As String
'** InputString is the string to convert.
'** StringLength is the lengt of the string

Dim InstrPos As Long
Dim StartPos As Integer
       
StartPos = 1
InstrPos = InStr(StartPos, InputString, vbLf)

Do Until InstrPos = 0
    '** insert the CRLF
    InputString = Left(InputString, InstrPos - 1) & vbCrLf & Right(InputString, Len(InputString) - InstrPos)
    StartPos = InstrPos + 2
    InstrPos = InStr(StartPos, InputString, vbLf)
Loop

Lf2Crlf = InputString
       
End Function

Public Function Crlf2Lf(InputString As String, StringLength As Integer) As String
'** I copied the Lf2Crlf and flipped it so we can convert from PC 2 Unix :^)
'** Richard Knechtel 6/20/2001

'** InputString is the string to convert.
'** StringLength is the lengt of the string

Dim InstrPos As Long
Dim StartPos As Integer
       
StartPos = 1
InstrPos = InStr(StartPos, InputString, vbCrLf)

Do Until InstrPos = 0
    '** insert the LF
    InputString = Left(InputString, InstrPos - 1) & vbLf & Right(InputString, Len(InputString) - InstrPos)
    StartPos = InstrPos + 2
    InstrPos = InStr(StartPos, InputString, vbCrLf)
Loop

Crlf2Lf = InputString
       
End Function

Public Function ReadTheFile(FileName As String, ErrorString As String) As Boolean
'** This function reads a UNIX file and converts it to PC format

Dim TempFile As String
Dim TempNum As Integer
Dim NrOfTimes As Long
Dim Remain As Long
Dim ReadNum As Integer
Dim WriteNum As Integer
Dim C As Integer
Dim ReadString As String * 4096 'This is the recordlength
Dim WriteString As String
Dim RecordLength As Integer
       
On Error GoTo errhandler
RecordLength = 4096

'** First copy the file to a tempory file and use this file to read from
'** Change the extension to tst

TempNum = InStr(Len(FileName) - 4, FileName, ".")
TempFile = Left(FileName, TempNum) & "tst"
FileCopy FileName, TempFile
Kill (FileName)

'** Calculate the nr of records who need to be read And the remaining bytes
NrOfTimes = FileLen(TempFile) \ RecordLength
Remain = FileLen(TempFile) Mod RecordLength

'** Prepare the filemanipulations
ReadNum = FreeFile
Open TempFile For Random As #ReadNum Len = RecordLength
WriteNum = FreeFile
Open FileName For Output As #WriteNum

'** Read the file and do the conversion
C = 1

Do Until NrOfTimes = 0
    Get #ReadNum, C, ReadString
    '** What Format does the user want to conver to:
    '** Richard Knechtel 6/20/2001
    If frmConv.optPC.Value = True Then '** RSK01 6/20/2001
        WriteString = Lf2Crlf(ReadString, RecordLength)
    End If '** RSK01 6/20/2001
    If frmConv.optUnix.Value = True Then '** RSK01 6/20/2001
        WriteString = Crlf2Lf(ReadString, RecordLength) '** RSK01 6/20/2001
    End If '** RSK01 6/20/2001
    Print #WriteNum, WriteString;
    NrOfTimes = NrOfTimes - 1
    C = C + 1
Loop

'** Read the remainings of the file

If Remain <> 0 Then
    ReadString = ""
    Get #ReadNum, C, ReadString
    '** remove the unwanted null chars from the string
    '** What Format does the user want to conver to:
    '** Richard Knechtel 6/20/2001
    WriteString = Left(ReadString, Remain)
    If frmConv.optPC.Value = True Then '** RSK01 6/20/2001
        WriteString = Lf2Crlf(WriteString, RecordLength)
    End If '** RSK01 6/20/2001
    If frmConv.optUnix.Value = True Then '** RSK01 6/20/2001
        WriteString = Crlf2Lf(WriteString, RecordLength) '** RSK01 6/20/2001
    End If '** RSK01 6/20/2001
    Print #WriteNum, WriteString
End If

'** Close the 2 files
Close WriteNum
Close #ReadNum

'** Remove the tempfile
Kill (TempFile)
ReadTheFile = True

Exit Function

errhandler:
       Debug.Print Err.Number & " / " & Err.Description
       ErrorString = "Error : " & Err.Number & " : " & Err.Description
       ReadTheFile = False
       
End Function

