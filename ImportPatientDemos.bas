Attribute VB_Name = "ImportPatientDemos"

Sub ImportPatientDemos()


'Created by Breya Walker
'Date:1.19.17
'Purpose: To import patient demographic information from any protocol spreadsheet and save into Patients Table in db


Dim Range As String
Dim i As Integer
Dim t As Integer
Dim j As Variant
Dim k As Integer
Dim h As Variant
Dim L As Variant
Dim p As Integer
Dim s As Integer
Dim f As Integer
Dim o As Integer
Dim u As Integer
Dim d As Integer
Dim y As Integer
Dim Name As Variant
Dim xls As Object
Set xls = CreateObject("Excel.Application")
Dim wkb  As Object
Dim NewFile As String
'Set wkb = GetObject("T:\Study Files\RERTEP\Data\RERTEPDatacollectedCopy.xlsx")
Dim wks As Object
'Set wks = wkb.Worksheets(1)
Dim dbs As DAO.Database
Dim r As DAO.Recordset
Dim n As DAO.Recordset
Set dbs = CurrentDb
Dim Protocol As String
Dim Fields As Field
Dim FieldNames As String
Dim FieldNames1 As String
Dim FieldNames2 As String
Dim FieldNames3 As String
Dim FieldNames4 As String
Dim FieldNames5 As String
Protocol = "SCCC" 'BE SURE TO CHANGE PROTOCOL NAME BEFORE STARTING

'Transfer in sheets from excel into table
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "NEWTABLE", "Filepath", True, "A:H"

'FieldNames = DoCmd.TransferSpreadsheet(8)


'Set Record to imported table

Set r = dbs.OpenRecordset("NEWTABLE")
FieldNames = "Stratum" 'Be sure to update fieldnames if additional fields to be added are seen
FieldNames1 = "Site"
FieldNames2 = "UPN"
FieldNames3 = "DX"
FieldNames4 = "Type"
FieldNames5 = "Cohort"
'Set n to record where copying should be going to
Set n = dbs.OpenRecordset("RecordSetReference")
Dim Names() As Variant
Dim FirstName() As Variant
Dim LastName() As Variant
Dim AnyName() As Variant
Dim MRN() As Variant
Dim DOB() As Variant
Dim Age() As Variant
Dim Race() As String
Dim Gender() As Variant
Dim Stratum() As Variant
Dim Inst() As Variant
Dim UPN() As Variant
Dim Dx() As Variant
Dim TYPEs() As Variant
Dim Cohort() As Variant


'Loop through the Names in the Names table and assign them to Names array
'r.MoveFirst
For i = 0 To r.RecordCount - 1 'Double check the number of null rows in recordset before starting. Have to adjust number to reflect number
       ReDim Preserve Names(i) 'Redim the Names array to account for changing dim of Names
        Names(i) = r!Name
    r.MoveNext
Next

'loop through DOB and store into DOB
r.MoveFirst
For p = 0 To r.RecordCount - 1
    ReDim Preserve DOB(p)
    DOB(p) = r!DOB
    r.MoveNext
Next

p = 0
'Loop through Age and store into Age
r.MoveFirst
For p = 0 To r.RecordCount - 1
    ReDim Preserve Age(p)
    Age(p) = r!Age
    r.MoveNext
Next

p = 0
'Loop through Gender and store into Gender
r.MoveFirst
For p = 0 To r.RecordCount - 1
    ReDim Preserve Gender(p)
    Gender(p) = r!Gender
    r.MoveNext
Next

f = 0
'Check and see if Site column exists in excel sheet
'If does create a site array
For Each Field In r.Fields
    If Field.Name = FieldNames1 Then
    r.MoveFirst
    For f = 0 To r.RecordCount - 1
        ReDim Preserve Inst(f)
        Inst(f) = r!Site
        r.MoveNext
    Next
    End If
Next

s = 0
'Check and see if Stratum column exists in excel sheet
'If does create a Stratum array
For Each Field In r.Fields
    If Field.Name = FieldNames Then
    r.MoveFirst
    For s = 0 To r.RecordCount - 1
        ReDim Preserve Stratum(s)
        Stratum(s) = r!Stratum
        r.MoveNext
    Next
    End If
Next

u = 0
'Check and see if UPN column exists in excel sheet
'If does create a UPN array
For Each Field In r.Fields
    If Field.Name = FieldNames2 Then
    r.MoveFirst
    For u = 0 To r.RecordCount - 1
        ReDim Preserve UPN(u)
        UPN(u) = r!MRN
        r.MoveNext
    Next
    End If
Next

d = 0
'Check and see if Dx column exists in excel sheet
'If does create a Dx array
For Each Field In r.Fields
    If Field.Name = FieldNames3 Then
    r.MoveFirst
    For d = 0 To r.RecordCount - 1
        ReDim Preserve Dx(d)
        Dx(d) = r!Dx
        r.MoveNext
    Next
    End If
Next

y = 0
'Check and see if Type of Dx column exists in excel sheet
'If does create a Type of Dx array
For Each Field In r.Fields
    If Field.Name = FieldNames4 Then
    r.MoveFirst
    For y = 0 To r.RecordCount - 1
        ReDim Preserve TYPEs(y)
        TYPEs(y) = r!Type
        r.MoveNext
    Next
    End If
Next

y = 0
'Check and see if Cohort column exists in excel sheet
'If does create a Cohort array
For Each Field In r.Fields
    If Field.Name = FieldNames5 Then
    r.MoveFirst
    For y = 0 To r.RecordCount - 1
        ReDim Preserve Cohort(y)
         Cohort(y) = r!Cohort
        r.MoveNext
    Next
    End If
Next

'Loop through Race and store into Race
r.MoveFirst
For o = 0 To r.RecordCount - 1
    ReDim Preserve Race(o)
    Race(o) = r!Races
    r.MoveNext
Next

i = 0
AnyName = Names

'Split First and last name and save as two elements in a two 2d array called AnyName
For Each j In Names 'Double array
    ReDim Preserve AnyName(i)
    AnyName(i) = Split(j, " ", -1)
    i = i + 1
Next j

t = 0
'Save First Name
For Each j In AnyName
 
   ReDim Preserve FirstName(t)
    FirstName(t) = AnyName(t)(0)
    t = t + 1
   
Next j

k = 0
'Save Last Name
For Each j In AnyName
 '
   ReDim Preserve LastName(k)
    LastName(k) = AnyName(k)(1)
    k = k + 1
   ' Next
Next j


'n.MoveFirst
'For i = 0 To 110 'Find out how to reference length of array in order to not use #
'       'ReDim Preserve Names(i) 'Redim the Names array to account for changing dim of Names
'        n.AddNew
'        n!FirstName = FirstName(i) 'Names(i) = r!Names
'        n!MRN(i) = r!MRN(i)
''    n.Update
'Next

'Save First Name
r.MoveFirst
For L = 0 To r.RecordCount - 1 'Double check recordcount number before import
       ReDim Preserve MRN(L) 'Redim the Names array to account for changing dim of Names
        MRN(L) = r!MRN
        'L = L + 1
    r.MoveNext
Next

i = 0
For Each h In MRN 'Find out how to reference length of array in order to not use #
       'This works just have to get it to loop through rest of MRNs resume 1.19.17
        n.AddNew
        n!FirstName = FirstName(i)
        n!LastName = LastName(i)
        n!MRN = MRN(i)
        n!Age = Age(i)
        n!Race = Race(i)
        n!Gender = Gender(i)
        n!DOB = DOB(i)
        n!Protocol = Protocol
        For Each Field In r.Fields
            If Field.Name = FieldNames Then
                n!Stratum = Stratum(i)
            ElseIf Field.Name = FieldNames1 Then
                n!Site = Inst(i)
            ElseIf Field.Name = FieldNames2 Then
                n!UPN = UPN(i)
            ElseIf Field.Name = FieldNames3 Then
                n!Dx = Dx(i)
            ElseIf Field.Name = FieldNames4 Then
               n!Type = TYPEs(i)
            ElseIf Field.Name = FieldNames5 Then
               n!Cohort = Cohort(i)
            Else
            End If
         Next
        i = i + 1
        n.Update
Next h
        
        
        
        

Debug.Print "DONE GIRL!";


End Sub
