Attribute VB_Name = "EnrollmentTable"
Function Contains(MRN, eMRN) As Boolean
End Function

Sub UpdateEnrollmentTable()
'Date created: 1.24.17
'Date modified: 6/20/2017
'Created by Breya Walker
'Purpose: The purpose of this standard module is to update the enrollment table everytime a recordset is added to the patients table. This macro should be run everytime a new patient is enrolled on a protocol
'if their information mistakenly is added into the excel spreadsheet
Dim NewEnr As New Enrollment
Dim NewDBS As New Demographics
Dim Name1 As String
Dim Name2 As String
Name1 = "Patients"
Name2 = "Enrollment"
NewDBS.SetDBS Name1, Name2 'Set new DBS

NewEnr.ProMRN (Name1)
NewEnr.ProName (Name1)
NewEnr.EmptyTable (Name2) 'Function fills an empty table with enrollments if for whatever reason the enrollment table is empty
NewEnr.RefillTable (Name2) 'This function adds new enrollments to an already existing enrollment table

End Sub
