VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim comp() As Variant
Dim j() As Variant
Dim objFLD As DAO.Field
Dim i As Integer
Dim q As Integer
Dim w As Integer
Dim a As Integer
Dim b As Integer
Dim MRN_() As Variant
Dim eMRN() As Variant
Dim eEnr() As Variant
Dim lb As Long
Dim ub As Long
Dim MyArr() As Variant
Dim MRNLen As Integer
Dim Match() As Variant
Dim EMRNL As Integer
Dim FieldsL As Integer
Dim Instrument() As Variant


Sub ChangeNums(Name1) 'Change the code numbers assigned with each spreadsheet to increment them by 1
    Dim DBS As DAO.Database
    Dim r As DAO.Recordset
    Dim t As Integer
    Dim p As Integer
    Dim FieldName As String
    Dim NewValue As Double
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)

    r.MoveFirst 'access first record
    

    FieldsL = r.Fields.Count
    
    For Each objFLD In r.Fields
    i = 0
ReDim Preserve comp(t)
FieldName = r.Fields(t).Name 'Name of field currently working
r.MoveFirst

For p = 0 To r.RecordCount - 1
    ReDim Preserve j(i)
    FieldName = r.Fields(t).Name
    j(i) = objFLD.Value
    If IsNumeric(objFLD.Value) = True Then
    NewValue = CVar(objFLD.Value)
    If objFLD.Value >= 1 And objFLD.Value < 14 Then
        r.Edit
        'Debug.Print NewValue & FieldName 'Before Changing code by adding 1
        NewValue = NewValue + 1
        'Debug.Print NewValue & FieldName 'After changing code by adding 1
        objFLD.Value = CSng(NewValue) 'Save New code as string
        r.Update
    Else
    End If
    j(i) = objFLD.Value
    
    End If

    
     r.MoveNext
i = i + 1
Next

comp(t) = r.Fields(t).Name
t = t + 1
Next
MsgBox "We changed the values by 1 in the" & Name1 & "Table"
End Sub


Sub Recode1(Name1) ' Recode all yeses into 1 and No into 0 if present
Dim k As Integer
k = 0
Dim DBS As DAO.Database
Set DBS = CurrentDb
Dim r As DAO.Recordset
Set r = DBS.OpenRecordset(Name1)

For Each objFLD In r.Fields
    If objFLD.Type = dbText Then  'Find all fields in recordset that are of type text
        If InStr(objFLD.Name, "Comment") = 0 Then
            i = 0
                    
            r.MoveFirst

            For k = 0 To r.RecordCount - 1
                ReDim Preserve j(i)
                j(i) = objFLD.Value
                If InStr(objFLD.Value, "yes") = 1 Then
                    r.Edit
                    objFLD.Value = CSng(1)
                    r.Update
                ElseIf InStr(objFLD.Value, "No") = 1 Then
                    r.Edit
                    objFLD.Value = CSng(0)
                    r.Update
                End If
            j(i) = objFLD.Value
            'm = m + 1
            r.MoveNext
            i = i + 1
            Next
        Else
    End If
    Else
    End If
Next
End Sub

Sub ImportEMRN(Name2) 'Import enrollment mrns and enrollment numbers
    Dim DBS As DAO.Database
    Set DBS = CurrentDb
    Dim r As DAO.Recordset
    Set r = DBS.OpenRecordset(Name2)
    Dim q As Integer
    r.MoveFirst 'Start at first record in recordset
    For q = 0 To r.RecordCount - 1
        ReDim Preserve eMRN(q)
        ReDim Preserve eEnr(q)
        eMRN(q) = r!MRN 'Create array of enrollment MRNs based on the MRNs in enrollment table
        eEnr(q) = r!Enrollment
        r.MoveNext
    Next
End Sub

Function CompareMRN(Name1, Name2) 'Compare the MRNs in the patient table to the enrollment table
    Dim r As DAO.Recordset
    Dim n As DAO.Recordset
    Dim DBS As DAO.Database
    Dim NewMRN As New Demographics
    Dim NewEMRN As New ProRecords
    Dim q As Integer
    Dim w As Integer
    Dim a As Integer
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)
    Set n = DBS.OpenRecordset(Name2)
   
    NewMRN.ImportMRN (Name1) 'Import MRNs from protocol MRNS based on the MRNs from r.table
    NewEMRN.ImportEMRN (Name2) 'Import the enrollment MRNS from the enrollment table along with their enrollment numbers
  
    lb = LBound(eMRN)
    ub = UBound(eMRN)
    
    q = 0
    
    
    MRNLen = DCount("[MRN]", Name1) 'Number of patients in r.Table
    EMRNL = DCount("[MRN]", Name2)
    q = 0
    w = 0
    i = 0
    b = 0
    
    'Step4.3: Do a search and compare of MRNs in r.Table with those in enrollment Table
    
    r.MoveFirst
    For b = 0 To MRNLen - 1
    q = 0
        For q = 0 To EMRNL - 1
            If r!MRN = eMRN(q) Then
                ReDim Preserve Match(0 To 1, 0 To ub)
                Match(0, i) = eMRN(q)
                Match(1, i) = eEnr(q)
                i = i + 1
            Else
                'Do not add this MRN to the Match file
            End If
        Next q
        r.MoveNext

    Next b

End Function

Function GatherInfo(Name1, Name2) 'MRNs that are in from the records table are in the enrollment table then copy over all their data from their excel protocol and Copy it to Protocol Records Table
    Dim a As Integer
    Dim aa As Integer
    Dim b As Integer
    Dim w As Integer
    Dim z As Integer
    Dim x As Integer
    Dim dd As Integer
    Dim cc As Integer
    Dim bb As Integer
    Dim ii As Integer
    Dim aaa As Integer
    Dim upward As Integer
    Dim objFLD As DAO.Field
    Dim Comments() As Variant
    Dim y As Integer
    Dim r As DAO.Recordset
    Dim DBS As DAO.Database
    Dim NewMRN As New Demographics
    Dim NewEMRN As New ProRecords
    Set DBS = CurrentDb
    Set r = DBS.OpenRecordset(Name1)

    lb = LBound(eMRN) 'Lower bound of enrollment table
    ub = UBound(eMRN) 'upper bound of enrollment table
    
    MRNLen = DCount("[MRN]", Name1) ' Length of MRN recordset in table
    EMRNL = DCount("[MRN]", Name2) 'Length of enrollment table
    
    upward = r.Fields.Count - 1 'The length of the number of fields-1

    'MRNs in the Matches array; find their data in recordset table and save

cc = 0
dd = 0

'Works as of : 1.26.2018 DO NOT TOUCH

    For y = 0 To EMRNL - 1
    q = 0
    a = 0
    r.MoveFirst
        While q < r.RecordCount
        a = 0
            If r!MRN = Match(0, dd) Then
                For a = 0 To upward
                ReDim Preserve Instrument(0 To EMRNL - 1, 0 To 5, 0 To upward)
                Instrument(cc, 0, b) = Match(1, dd) 'copy over enrollment number from enrollment table
                Instrument(cc, 1, b) = comp(z)
                x = 1
                aa = 0
                aaa = 0
                i = 0

                For Each objFLD In r.Fields
                    If objFLD.Type = dbDate Then 'Find all timepoint dates
                        If InStr(1, objFLD.Name, "T") = 1 Then  'Like "T" Then
                            Instrument(cc, 2, aa) = "T" & x 'increment T
                            Instrument(cc, 3, aa) = r.Fields(objFLD.Name).Value 'Add date to instrument column
                            x = x + 1
                            aa = aa + 1
                        Else
                        End If
                        End If
                        If objFLD.Type = dbText Then 'find all comments in recordset
                            If InStr(1, objFLD.Name, "Comment") = 1 Then
                                Instrument(cc, 4, aaa) = objFLD.Value
                                aaa = aaa + 1
                            Else
                                Instrument(cc, 5, aaa) = objFLD.Value
                                aaa = aaa + 1
                            End If
                        End If
                    'End If
                Next

                b = b + 1
                z = z + 1
                Next a
            Else
        End If
        r.MoveNext
        b = 0
        q = q + 1
        Wend
        dd = dd + 1
        'b = 0
        cc = cc + 1
        z = 0

Next y
        

End Function

Function UpdateProTable(Name1, Name3) 'All information from the GatherInfo function will be used here to update protocol table
Dim a As Integer
Dim z As Integer
Dim x As Integer
Dim dd As Integer
Dim cc As Integer
Dim ii As Integer
Dim iii As Integer
Dim y As Integer
Dim p As Integer
Dim upward As Variant
Dim r As DAO.Recordset
Dim n As DAO.Recordset
Dim DBS As DAO.Database
Set DBS = CurrentDb
Set r = DBS.OpenRecordset(Name1)
Set n = DBS.OpenRecordset(Name3)

'Set integer values
dd = 0
cc = 0
x = 0
z = 0
y = 0
ii = 0
iii = 0
a = 0
p = 0
'get value of number of timepoints for this particular protocol
upward = InputBox("How many timepoints are on this protocol")

'Add new patients to empty protocol records table
'Check and see if instruments is empty at x index if not resume if yes exit function
If n.EOF = True Then
        While dd <= MRNLen
            For a = 0 To MRNLen - 1
            If IsEmpty(Instrument(p, 0, 0)) = False Then
            n.AddNew
                x = 0
                While x < upward
                        n.AddNew
                        n!Timepoint = Instrument(cc, 2, z)
                        n!Enrollment = Instrument(cc, 0, z)
                        n!Occurence = y
                       ' n!Comment = Instrument(cc, 4, z)
                        n!TPDate = Instrument(cc, 3, z)
                        n!Instrument = Instrument(cc, 1, ii)
                            If Not IsNull(Instrument(cc, 5, ii)) = True Then
                                If CStr(Instrument(cc, 5, ii)) = "1" Then
                                    n!InstruPres = 1
                                    n!Comment = Instrument(cc, 4, z)
                                 Else
                                  n!InstruPres = Val(Instrument(cc, 5, ii))
                                  n!Comment = Instrument(cc, 4, z)
                                End If
                             Else ': IsNull (Instrument(cc, 5, ii)) 'Instrument(cc, 3, ii) <> "1" 'Or Instrument(cc, 3, iii) <> 1
                                 'n!InstruPres = Val(Instrument(cc, 5, ii))
                             End If
                            y = y + 1
                            z = z + 1
                            x = x + 1
                            'ii = ii + 1
                            'iii = iii + 1
                        n.Update
                    Wend
                    z = 0
                    'ii = 0
                    'iii = 0
                    cc = cc + 1
                Else
                 Exit Function
                End If
                p = p + 1
                'ii = 0
            Next a
                dd = dd + 1
                p = 0
                cc = 0
                ii = ii + 1
                iii = iii + 1
            Wend
    Else
    n.MoveFirst
End If
'if protocol table is not empty and has records in it already
If n.EOF = False Then
        While dd <= MRNLen
            For a = 0 To MRNLen - 1
            If IsEmpty(Instrument(p, 0, 0)) = False Then
            n.AddNew
                x = 0
                While x < upward
                n.AddNew
                        n!Timepoint = Instrument(cc, 2, z)
                        n!Enrollment = Instrument(cc, 0, z)
                        n!Occurence = y
                       ' n!Comment = Instrument(cc, 4, z)
                        n!TPDate = Instrument(cc, 3, z)
                        n!Instrument = Instrument(cc, 1, ii)
                            If Not IsNull(Instrument(cc, 5, ii)) = True Then
                                If CStr(Instrument(cc, 5, ii)) = "1" Then
                                    n!InstruPres = 1
                                    n!Comment = Instrument(cc, 4, z)
                                 Else
                                  n!InstruPres = Val(Instrument(cc, 5, ii))
                                  n!Comment = Instrument(cc, 4, z)
                                End If
                             Else ': IsNull (Instrument(cc, 5, ii)) 'Instrument(cc, 3, ii) <> "1" 'Or Instrument(cc, 3, iii) <> 1
                                 'n!InstruPres = Val(Instrument(cc, 5, ii))
                             End If
                            y = y + 1
                            z = z + 1
                            x = x + 1
                            'ii = ii + 1
                            'iii = iii + 1
                        n.Update
                    Wend
                    z = 0
                    'ii = 0
                    'iii = 0
                    cc = cc + 1
                Else
                 Exit Function
                End If
                p = p + 1
                'ii = 0
            Next a
                dd = dd + 1
                p = 0
                cc = 0
                ii = ii + 1
                iii = iii + 1
            Wend

    n.MoveFirst
  End If


'If n.EOF = False Then
 '   While dd <= MRNLen
  '      For a = 0 To MRNLen
   '     If IsEmpty(Instrument(p, 0, 0)) = False Then
    '        n.AddNew
    '            x = 0

'                While x < upward
'                n.AddNew
 '                   n!Timepoint = Instrument(cc, 2, z)
  '                  n!Enrollment = Instrument(cc, 0, z)
   '                 n!Occurence = y
    '                n!Comment = Instrument(cc, 4, z)
     '               n!TPDate = Instrument(cc, 3, z)
      '              n!Instrument = Instrument(cc, 1, ii)
       '                     If Not IsNull(Instrument(cc, 3, ii)) = "1" Then ' And Instrument(cc, 3, iii) = 1 Then
        '                         n!InstruPres = 1
         '                    Else: IsNull (Instrument(cc, 3, ii)) <> "1" 'Or Instrument(cc, 3, iii) <> 1
          '                       n!InstruPres = 0
           '                  End If
            '                'If Not IsNull(Instrument(cc, 3, iii)) And Instrument(cc, 3, iii) = 1 Then
             '               '     n!InstruPres = 1
              '              ' Else: IsNull (Instrument(cc, 3, iii)) Or Instrument(cc, 3, iii) <> 1
               '             '     n!InstruPres = 0
                '            ' End If
                 '           y = y + 1
                  '          z = z + 1
                   '         x = x + 1
                    '        'ii = ii + 1
                     '       'iii = iii + 1
                      '  n.Update
                    'Wend
                    'z = 0
                    'ii = 0
                    'iii = 0
                   ' cc = cc + 1
           ' Else
            '    Exit Function
            'End If
            'p = p + 1
            'Next a
            '    dd = dd + 1
             '   p = 0
             '   cc = 0
              '  ii = ii + 1
              '  iii = iii + 1
            'Wend
'End If
   ' Else
   'n.MoveFirst
End Function

