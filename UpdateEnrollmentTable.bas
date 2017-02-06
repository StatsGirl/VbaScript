Attribute VB_Name = "UpdateEnrollmentTable"
Function Contains(MRN, eMRN) As Boolean
End Function

Sub UpdateEnrollmentTable()
'purpose of this macro is to update the enrollment table everytime a recordset is added to the patients table
'Breya Walker
'1.24.17
Dim dbs As DAO.Database
Set dbs = CurrentDb
Dim p As DAO.Recordset
Set p = dbs.OpenRecordset("Patients")
Dim e As DAO.Recordset
Set e = dbs.OpenRecordset("Enrollment")
Dim i As Integer
Dim j As Variant
Dim jj As Variant
Dim t As Integer
Dim m As Variant
Dim k As Integer
Dim b As Integer
Dim CurrentMRNs As Integer
Dim NewIndexMRN As Integer
Dim NewIndex() As Variant
Dim NewReferencePoint As Integer
Dim num As Variant
Dim MRN() As Variant
Dim eMRN() As Variant
Dim rv As Boolean
Dim Protocol() As Variant
Dim ArrayList() As Variant
Dim ArrayList2() As Variant
NewReferencePoint = 0
CurrentMRNs = DCount("[MRN]", "Patients") 'Get length of MRNs including the MRNs that do not have values in each field
'.FieldSize
'Step1: Store MRN and protocol name into arrays from patients table
'p.MoveFirst
For i = 0 To CurrentMRNs - 1 ' RecordCount - 1 'Double check the number of null rows in recordset before starting. Have to adjust number to reflect number
      ReDim Preserve MRN(i) 'Redim the MRN array to account for changing dim of MRN
        MRN(i) = p!MRN
    p.MoveNext
Next

'loop through protocol name and storep

p.MoveFirst
For i = 0 To CurrentMRNs - 1
    ReDim Preserve Protocol(i)
    Protocol(i) = p!Protocol
    p.MoveNext
Next

i = 0

j = 0
If e.EOF Then 'If we are after the first record set then EOF = true
For Each j In MRN
    e.AddNew
    e!MRN = MRN(i)
    e!ProtocolName = Protocol(i)
    e!Enrollment = t
    e.Update
    i = i + 1
    t = t + 1
Next j

'Create a new eMRN variables based on new table generated
e.MoveFirst
For m = 0 To e.RecordCount - 1
    ReDim Preserve eMRN(m)
    eMRN(m) = e!MRN
    e.MoveNext
Next
End If

e.MoveFirst 'Start at first record in recordset
For m = 0 To e.RecordCount - 1
    ReDim Preserve eMRN(m)
    eMRN(m) = e!MRN 'Create array of enrollment MRNs based on the MRNs in enrollment table
    e.MoveNext
Next
'MRNs are not sorted in ascending order in array. Check the dimensionality of our MRN and eMRN arrays.
i = 0
jj = 0
'Currently p recordcount is starting at 1093 when we have 1122 patients check 2.2.17
k = UBound(MRN) 'UBound is similar to getting the length of an array
num = CurrentMRNs 'use patient MRN record count num var to create new enrollment number for patient
NewIndexMRN = k - UBound(eMRN) 'UBound(eMRN) 'The difference btw the two array sizes
NewReferencePoint = k - NewIndexMRN 'Value to index in MRN() when copying over info into eMRN table

ReDim NewIndex(NewIndexMRN) 'equal to the length of the value in NewIndexMRN
If e.EOF = True Then
For Each jj In NewIndex 'for each index in NewIndex array
'Start interation at NewReferencePoint in MRN() NewIndexMRN times NewReferencePoint=NewReferencePoint+1
    If UBound(MRN) <> UBound(eMRN) Then 'UBound(eMRN) Then 'If the array dimensions are different this means we have new patient entries
            e.AddNew 'Add new patient information to Enrollment table
            e!MRN = MRN(NewReferencePoint) 'e!MRN is equal to the newest patient information stored in MRN(k) with k being equal to the last index in the array
            e!ProtocolName = Protocol(NewReferencePoint) 'adds new enrollment based NewReferencePoint value
            e!Enrollment = NewReferencePoint + 1 'create new enrollment number
            e.Update
            NewReferencePoint = NewReferencePoint + 1
     End If
  Next jj
End If


End Sub
