Attribute VB_Name = "ImportDemographics"
Option Compare Database

Sub ImportDemographics_()
'---------------------------------------------
'Date created: 1.19.17
'Date modified: 6/20/2017
'Created by Breya Walker
'Purpose: The purpose of this standard module is to import patient demographics from a protocol spreadsheet into the patients table in the DB.

Dim NewDBS As New Demographics 'Call Class module Demographics and all of its functions
Dim Name1 As String
Dim Name2 As String
Dim enterName1 As String
Dim enterPath As String
enterName1 = InputBox("What is the name of this protocol")
enterPath = InputBox("Copy and Paste the path and file for this protocol")
Name1 = enterName1 'Be sure to change name of protocol when loading
Name2 = "Patients" 'Change output location if necessary
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, enterName1, enterPath, True 'Change directory of protocol and length of ranges in sheet before starting


'------------------------------------Import Patient Demographics from protocol sheet to patients information sheet--------------------------------------
NewDBS.SetDBS Name1, Name2 'Set NEWDBS as setDBS function in demographics and call Protocol sheet and patient sheet in
NewDBS.ImportMRN (Name1)
NewDBS.ImportDOB (Name1)
NewDBS.ImportName (Name1)
NewDBS.ImportAge (Name1)
NewDBS.gender (Name1)
NewDBS.Race (Name1)
NewDBS.FieldNames (Name1)
NewDBS.FillTable Name1, Name2

End Sub


