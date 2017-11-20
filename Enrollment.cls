VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim MRN() As Variant
Dim CurrentMRNs As Integer
Dim Protocol() As Variant
Dim k As Integer
Dim eMRN() As Variant
    
Function ProMRN(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    CurrentMRNs = DCount("[MRN]", Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1 'Recordcount -1 to get length of number of recordsets in table
           ReDim Preserve MRN(L) 'Redim the MRNs
            MRN(L) = r!MRN
        r.MoveNext
    Next
End Function

Sub ProName(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
        For L = 0 To CurrentMRNs - 1
        ReDim Preserve Protocol(L)
        Protocol(L) = r!Protocol
        r.MoveNext
    Next
End Sub

Sub EmptyTable(Name2)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Dim j As Variant
    Dim i As Integer
    Dim t As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name2)
    If r.EOF Then 'If we are after the first record set then EOF = true
        For Each j In MRN
            r.AddNew
            r!MRN = MRN(i)
            r!ProtocolName = Protocol(i)
            r!Enrollment = t
            r.Update
            i = i + 1
            t = t + 1
        Next j

        'Create a new eMRN variables based on new table generated
        r.MoveFirst
        For L = 0 To r.RecordCount - 1
            ReDim Preserve eMRN(L)
            eMRN(L) = r!MRN
            r.MoveNext
        Next
        End If
        
        r.MoveFirst 'Start at first record in recordset
        For L = 0 To r.RecordCount - 1
             ReDim Preserve eMRN(L)
                eMRN(L) = r!MRN 'Create array of enrollment MRNs based on the MRNs in enrollment table
                r.MoveNext
        Next
End Sub

Sub RefillTable(Name2)
'Create a new eMRN variables based on new table generated
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Dim j As Variant
    Dim i As Integer
    Dim t As Integer
    Dim num As Integer
    Dim NewIndexMRN As Integer
    Dim NewIndex() As Variant
    Dim NewReferencePoint As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name2) 'Enter data into enrollment table

    k = UBound(MRN) 'UBound is similar to getting the length of an array
    num = CurrentMRNs 'use patient MRN record count num var to create new enrollment number for patient
    NewIndexMRN = k - UBound(eMRN) 'UBound(eMRN) 'The difference btw the two array sizes
    NewReferencePoint = k - NewIndexMRN 'Value to index in MRN() when copying over info into eMRN table
    
    ReDim NewIndex(NewIndexMRN) 'equal to the length of the value in NewIndexMRN
    If r.EOF = True Then
    For Each j In NewIndex 'for each index in NewIndex array
    'Start interation at NewReferencePoint in MRN() NewIndexMRN times NewReferencePoint=NewReferencePoint+1
        If UBound(MRN) <> UBound(eMRN) Then 'UBound(eMRN) Then 'If the array dimensions are different this means we have new patient entries
                r.AddNew 'Add new patient information to Enrollment table
                r!MRN = MRN(NewReferencePoint) 'e!MRN is equal to the newest patient information stored in MRN(k) with k being equal to the last index in the array
                r!ProtocolName = Protocol(NewReferencePoint) 'adds new enrollment based NewReferencePoint value
                r!Enrollment = NewReferencePoint + 1 'create new enrollment number
                r.Update
                NewReferencePoint = NewReferencePoint + 1
         End If
      Next j
    End If

End Sub
