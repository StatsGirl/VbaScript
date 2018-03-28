VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Demographics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public r As DAO.Recordset
Public n As DAO.Recordset
Dim MRN() As Variant
Dim DOB() As Variant
Dim FirstName() As Variant
Dim LastName() As Variant
Dim Ages() As Variant
Dim genders() As Variant
Dim Races() As Variant
Dim Stratum() As Variant
Dim Site() As Variant
Dim UPN() As Variant
Dim DX() As Variant
Dim Types() As Variant
Dim Cohort() As Variant
Dim FieldNamesarr() As Variant

'The following is a property of the SetDBS objects.
Public Function SetDBS(Name1 As String, Name2 As String, Optional Name3)
    Dim DBS As DAO.Database
    Dim r As DAO.Recordset
    Dim n As DAO.Recordset
    Set DBS = CurrentDb 'MsgBox "your db has been set" & DBS.Name
    Set r = DBS.OpenRecordset(Name1)
    Set n = DBS.OpenRecordset(Name2)
End Function

'The following is a method of the ImportMRN objects.

Public Function ImportMRN(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1 'Recordcount -1 to get length of number of recordsets in table
           ReDim Preserve MRN(L) 'Redim the MRNs
            MRN(L) = r!MRN
        r.MoveNext
    Next
    'MRN1
End Function

'Sub MRN1()
  '  MRN1() = MRN
 'Call ImportMRN.MRN1
'End Sub

'The following is a method of the ImportDOB objects.

Sub ImportDOB(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1 'Recordcount -1 to get length of number of recordsets in table
        ReDim Preserve DOB(L) 'Redim DOB()
        DOB(L) = r!DOB
        r.MoveNext
    Next
End Sub

'The following is a method of the ImportName objects.

Sub ImportName(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)

    Dim Anyname() As Variant
    Dim Names() As Variant
    Dim L As Integer
    Dim t As Integer
    Dim j As Variant
    r.MoveFirst
    For L = 0 To r.RecordCount - 1 'Recordcount -1 to get length of number of recordsets in table
        ReDim Preserve Names(L) 'Redim Names()
        Names(L) = r!Name
        r.MoveNext
    Next

    Anyname = Names
    L = 0
    'Split First and last name and save as two elements in a two 2d array called AnyName
    For Each j In Names 'Double array
        ReDim Preserve Anyname(L)
        Anyname(L) = Split(j, " ", -1)
        L = L + 1
    Next j
    
    t = 0
    'Save First and last name separately
    For Each j In Anyname
       ReDim Preserve FirstName(t)
       ReDim Preserve LastName(t)
        FirstName(t) = Anyname(t)(0) 'Save First name
        LastName(t) = Anyname(t)(1) 'Save last name
        t = t + 1
    Next j

End Sub

'The following is a method of the ImportAge objects.

Sub ImportAge(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1 'Recordcount -1 to get length of number of recordsets in table
        ReDim Preserve Ages(L) 'Redim Age
        Ages(L) = r!Age
        r.MoveNext
    Next
End Sub

'The following is a method of the ImportGender objects.
Sub gender(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1
        ReDim Preserve genders(L) 'Recordcount -1 to get length of number of recordsets in table
        genders(L) = r!gender 'Redim Gender()
        r.MoveNext
    Next
End Sub


'The following is a method of the ImportRace objects.
Sub Race(Name1)
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim L As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    r.MoveFirst
    For L = 0 To r.RecordCount - 1
        ReDim Preserve Races(L) 'Recordcount -1 to get length of number of recordsets in table
        Races(L) = r!Race 'Redim Race()
        r.MoveNext
    Next
End Sub

'The following is a method in FieldsNames Objects.
Sub FieldNames(Name1)
    Dim Fld As Field
    FieldNamesarr = Array("Stratum", "Site", "UPN", "DX", "Type", "Cohort")
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Set DBS = CurrentDb
    
    Dim L As Integer
    Dim i As Integer

    Set r = DBS.OpenRecordset(Name1)
    Dim UniqueField() As Variant
    i = 0
    
        For Each Fld In r.Fields 'Be sure to set a dim object Field in module when using variable
            If Fld.Name = FieldNamesarr(0) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve Stratum(L)
                Stratum(L) = r!Stratum 'get name of field from Recordset
                r.MoveNext
            Next
            End If
            
            If Fld.Name = FieldNamesarr(1) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve Site(L)
                Site(L) = r!Site 'get name of field from Recordset
                r.MoveNext
            Next
            End If
            
            If Fld.Name = FieldNamesarr(2) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve UPN(L)
                UPN(L) = r!UPN 'get name of field from Recordset
                r.MoveNext
            Next
            End If
            
            If Fld.Name = FieldNamesarr(3) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve DX(L)
                DX(L) = r!DX 'get name of field from Recordset
                r.MoveNext
            Next
            End If
            
            If Fld.Name = FieldNamesarr(4) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve Types(L)
                Types(L) = r!Type 'get name of field from Recordset.
                r.MoveNext
            Next
            End If
            
            If Fld.Name = FieldNamesarr(5) Then
            r.MoveFirst
            For L = 0 To r.RecordCount - 1
                ReDim Preserve Cohort(L)
                Cohort(L) = r!Cohort 'get name of field from Recordset.
                r.MoveNext
            Next
            End If
        Next
End Sub

'The following is a method in FillTable objects.
Function FillTable(Name1, Name2)
Dim n As DAO.Recordset
Dim r As DAO.Recordset
Dim DBS As DAO.Database
Dim L As Integer
Dim Fld As Field
Set DBS = CurrentDb
Set n = DBS.OpenRecordset(Name2)
Set r = DBS.OpenRecordset(Name1)
Dim i As Integer
Dim h As Variant

Dim Protocol As String
Protocol = Name1
i = 0
For Each h In MRN
        n.AddNew
        n!FirstName = FirstName(i)
        n!LastName = LastName(i)
        n!MRN = MRN(i)
        n!Age = Ages(i)
        n!Race = Races(i)
        n!gender = genders(i)
        n!DOB = DOB(i)
        n!Protocol = Protocol
        For Each Fld In r.Fields
            If Fld.Name = FieldNamesarr(0) Then
                n!Stratum = Stratum(i)
            ElseIf Fld.Name = FieldNamesarr(1) Then
                n!Site = Site(i)
            ElseIf Fld.Name = FieldNamesarr(2) Then
                n!UPN = UPN(i)
            ElseIf Fld.Name = FieldNamesarr(3) Then
                n!DX = DX(i)
            ElseIf Fld.Name = FieldNamesarr(4) Then
               n!Types = Types(i)
            ElseIf Fld.Name = FieldNamesarr(5) Then
               n!Cohort = Cohort(i)
            Else
            End If
         Next
        i = i + 1
        n.Update
Next h
End Function
        


