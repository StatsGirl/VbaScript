Attribute VB_Name = "ProtocolTable"
Option Compare Database

Sub ProtocolRecs()

'Date created: 2.6.2017
'Date Modified: 6/21/2017
'Purpose: The purpose of this standard module is to create new protocol encounters based on the subject enrollment # (from enrollment table),
'pulling timepoint information (i.e., TP date) and TP number from Patients table, and
'instruments collected at each timepoint.
Dim NewTab As New ProRecords
Dim NewDBS As New Demographics
Dim Name1 As String
Dim Name2 As String
Dim Name3 As String
Dim enterName1 As String
Dim enterPath As String
enterName1 = InputBox("What is the name of this protocol")
enterPath = InputBox("Copy and Paste the path and file for this protocol")
Name1 = enterName1 'Always the protocol you are trying to referencing
Name2 = "Enrollment" 'Always the enrollment table
Name3 = "ProtocolRecords" 'Always the protocol records table
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, enterName1, enterPath, True 'create access table for protocol excel spreadsheet

NewDBS.SetDBS Name1, Name2, Name3
NewTab.ChangeNums (Name1)
NewTab.Recode1 (Name1)
NewTab.ImportEMRN (Name2)
NewTab.CompareMRN Name1, Name2
NewTab.GatherInfo Name1, Name2
NewTab.UpdateProTable Name1, Name3
End Sub
