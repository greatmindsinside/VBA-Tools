Attribute VB_Name = "SSNAlreadyInUse"
Option Explicit
Option Compare Text

Sub Start()

    SSNStringCompare ("SsnAlreadyInUseInSystem:socialSecurityNumber")
    SSNStringCompare ("SsnAlreadyInUseInSystem:dependentSocialSecurityNumber")

End Sub

Private Function SSNStringCompare(sTheText As String) As Range
    
    Dim rgFound As Variant
    Dim FirstRangeFound As String
    '------------------------------------
    'Column Numbers
    '------------------------------------
    Dim SSNErrorCX As Integer: SSNErrorCX = 104
    Dim SSNErrorCY As Integer: SSNErrorCY = 105
    Dim firstNameCol As Integer: firstNameCol = 5
    Dim lastNameCol As Integer: lastNameCol = 7
    Dim SSNColNum As Integer: SSNColNum = 4
    Dim dateOfBirthCol As Integer: dateOfBirthCol = 8
    Dim depSSNCol As Integer: depSSNCol = 20
    Dim depDOBCol As Integer: depDOBCol = 24
    Dim depFirstNameCol As Integer: depFirstNameCol = 21
    Dim depLastNameCol As Integer: depLastNameCol = 23
    Dim ClientMemCol As Integer: ClientMemCol = 49
    Dim DepClientMemCol As Integer: DepClientMemCol = 54
           
    Dim IsPrimary As Boolean: IsPrimary = False
    Dim IsDependent As Boolean: IsDependent = False
    Dim DOBMatch As Boolean: DOBMatch = False
    
    Dim CXErrorRowText As String
    Dim ErrorRowSSN As String
    Dim ErrorRowDOB As String
   
    Dim PriOnDepSameRecord As String: PriOnDepSameRecord = "The primary person is a dependent on another eligibility record with the same company."
    Dim DepIsPri As String: DepIsPri = "The dependent person is the primary or is associated with the primary on another eligibility record with the same company."
       
    Dim socialSecurityNumber As String
    Dim firstName As String
    Dim lastName As String
    Dim dateOfBirth As String
   
    Dim depFirstName As String
    Dim depLastName As String
    Dim depSSN As String
    Dim depDOB As String
          
    Dim sFirstLastName As String
    Dim sNameArray() As String
    Dim sNameArrayRow() As String
    
    
    With ActiveSheet.UsedRange
        Set rgFound = .Find(What:=sTheText, _
        After:=.Cells(.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
        
        If Not rgFound Is Nothing Then
            FirstRangeFound = rgFound.Address
            
            Do
                '===============================================
                ' Print the location that the data was found in
                '===============================================
                Debug.Print sTheText & " was found in Cell: " & rgFound.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False)
                    
                'Cells(rgFound.Row, SSNColNum).Interior.ColorIndex = 8
                
                '---------------------------------------------------
                'Get the values the client placed for the primary
                '---------------------------------------------------
                socialSecurityNumber = Cells(rgFound.Row, SSNColNum).Value
                firstName = Cells(rgFound.Row, firstNameCol).Value
                lastName = Cells(rgFound.Row, lastNameCol).Value
                dateOfBirth = Cells(rgFound.Row, dateOfBirthCol).Value
                depSSN = Cells(rgFound.Row, depSSNCol).Value
                depDOB = Cells(rgFound.Row, depDOBCol).Value
                depFirstName = Cells(rgFound.Row, depFirstNameCol).Value
                depLastName = Cells(rgFound.Row, depLastNameCol).Value
                
                             
                CXErrorRowText = Cells(rgFound.Row, SSNErrorCX).Value
                                                          
                sNameArrayRow = Split(CXErrorRowText, ".")
                Dim sName As Variant
                For Each sName In sNameArrayRow
                    Debug.Print ("sNameArrayRow: " + sName)
                    Dim sPosString As Integer
                    sPosString = InStr(1, sName, "DOB", 1)
                    If sPosString <> 0 Then
                        ErrorRowDOB = GetDOB(sName)
                           
                    End If
                                      
                Next
                              
                'ErrorRowDOB = GetDOB(sName)
                ErrorRowSSN = GetSSN(sNameArrayRow(0))
                              
                'sNameArrayRow is split based on the period so if the name has a period in it (Shepherd Jr.) then the DOB will be in a different array postion.
                              
                '----------------------------------
                ' Get Just The Name fom the string
                '----------------------------------
                CXErrorRowText = StripNonAlpha(sNameArrayRow(0))
                CXErrorRowText = Replace(CXErrorRowText, "SSN  is already assigned to ", "")
                CXErrorRowText = Replace(CXErrorRowText, "DOB", "")
                CXErrorRowText = Replace(CXErrorRowText, PriOnDepSameRecord, "")
                CXErrorRowText = Replace(CXErrorRowText, DepIsPri, "")
                Debug.Print ("The Error Row Text Before Array: " + CXErrorRowText)
                               
                sNameArray = Split(CXErrorRowText, " ", 2)
                                
                '------------------------------------------------------------
                'Check to see if the SSN Matches one of the client SSN Fields
                '-------------------------------------------------------------
                If StrComp(ErrorRowSSN, socialSecurityNumber, 1) = 0 And StrComp(ErrorRowDOB, dateOfBirth, 1) = 0 Then
                    Debug.Print ("Is The Primary and The DOB Matches...")
                    IsPrimary = True
                    DOBMatch = True
                                          
                    Call CheckPrimaryName(firstName, lastName, sNameArray, rgFound)
                    Cells(rgFound.Row, dateOfBirthCol).Interior.ColorIndex = 0
                ElseIf StrComp(ErrorRowSSN, socialSecurityNumber, 1) = 0 Then
                    Debug.Print ("Is The Primary and The DOB DOES NOT Match.")
                    IsPrimary = True
                    DOBMatch = False
                    
                    Call CheckPrimaryName(firstName, lastName, sNameArray, rgFound)
                    Cells(rgFound.Row, dateOfBirthCol).Interior.ColorIndex = 3
                    'Change The Date Of Birth
                    Cells(rgFound.Row, dateOfBirthCol).Value = ErrorRowDOB
                Else
                    If StrComp(ErrorRowSSN, depSSN, 1) = 0 And StrComp(ErrorRowDOB, depDOB, 1) = 0 Then
                        Debug.Print ("Is The Dependent and The DOB Matches...")
                        IsDependent = True
                        DOBMatch = True
                        Call CheckDependentName(depFirstName, depLastName, sNameArray, rgFound)
                        Cells(rgFound.Row, depDOBCol).Interior.ColorIndex = 0
                    ElseIf StrComp(ErrorRowSSN, depSSN, 1) = 0 Then
                        Debug.Print ("Is The Dependent and The DOB Does Not Match Matches...")
                        IsDependent = True
                        Call CheckDependentName(depFirstName, depLastName, sNameArray, rgFound)
                        Cells(rgFound.Row, depDOBCol).Interior.ColorIndex = 3
                        'Change The Date Of Birth
                        Cells(rgFound.Row, depDOBCol).Value = ErrorRowDOB
                    End If
                End If
                 
                Set rgFound = .FindNext(rgFound)
            Loop While Not rgFound Is Nothing And rgFound.Address <> FirstRangeFound
        End If
    End With
End Function

Sub CheckDependentName(firstName As String, lastName As String, ByRef sNameArray As Variant, rgFound As Variant)

    Dim depFirstNameCol As Integer: depFirstNameCol = 21
    Dim depLastNameCol As Integer: depLastNameCol = 23
      
    sNameArray(0) = Trim(sNameArray(0))
    sNameArray(1) = Trim(sNameArray(1))
     
    If firstName <> sNameArray(0) Then
        Cells(rgFound.Row, depFirstNameCol).Interior.ColorIndex = 3
    Else
        Cells(rgFound.Row, depFirstNameCol).Interior.ColorIndex = 0
    End If
            
    If lastName <> sNameArray(1) Then
          Cells(rgFound.Row, depLastNameCol).Interior.ColorIndex = 3
          'MsgBox "Last Name(Client): " + lastName
    Else
        Cells(rgFound.Row, depLastNameCol).Interior.ColorIndex = 0
    End If
    
    Debug.Print ("First Name(Client): " + firstName)
    Debug.Print ("Last Name(Client): " + lastName)
    Debug.Print ("First Name(Database): " + sNameArray(0))
    Debug.Print ("Last Name(Database): " + sNameArray(1))

End Sub

Sub CheckPrimaryName(firstName As String, lastName As String, ByRef sNameArray As Variant, rgFound As Variant)

    Dim firstNameCol As Integer: firstNameCol = 5
    Dim lastNameCol As Integer: lastNameCol = 7
      
    sNameArray(0) = Trim(sNameArray(0))
    sNameArray(1) = Trim(sNameArray(1))
    
    If firstName <> sNameArray(0) Then
        Cells(rgFound.Row, firstNameCol).Interior.ColorIndex = 3
        'MsgBox ("First Name(Database): " + sNameArray(0))
        'Cells(rgFound.Row, firstNameCol).Value = sNameArray(0)
    Else
        Cells(rgFound.Row, firstNameCol).Interior.ColorIndex = 0
    End If
            
    If lastName <> sNameArray(1) Then
          Cells(rgFound.Row, lastNameCol).Interior.ColorIndex = 3
          'Cells(rgFound.Row, firstNameCol).Value = sNameArray(1)
    Else
        Cells(rgFound.Row, lastNameCol).Interior.ColorIndex = 0
    End If
    
    Debug.Print ("First Name(Client): " + firstName)
    Debug.Print ("Last Name(Client): " + lastName)
    Debug.Print ("First Name(Database): " + sNameArray(0))
    Debug.Print ("Last Name(Database): " + sNameArray(1))
    
End Sub

Private Function GetSSN(sTheString) As String
    'need to rewrite this so it only returns the first 8 digits found
    'At this point it returns only numbers which includes the date of birth
    
    Dim ErrorRowSSN As String
    Dim ObjRegex
    
    Set ObjRegex = CreateObject("vbscript.regexp")
    With ObjRegex
        .Global = True
        .Pattern = "\D"
        
        ErrorRowSSN = .Replace(sTheString, vbNullString)
        ErrorRowSSN = Left(ErrorRowSSN, 9)
        
        GetSSN = ErrorRowSSN
    End With
    
End Function

Private Function GetDOB(sTheString) As String
    'need to rewrite this so it only returns the first 8 digits found
    'At this point it returns only numbers which includes the date of birth
    Debug.Print ("The DOB Before:" + sTheString)
    Dim ErrorRowDOB As String
    Dim ObjRegex
    
    Set ObjRegex = CreateObject("vbscript.regexp")
    With ObjRegex
        .Global = True
        .Pattern = "\D"
        
        ErrorRowDOB = .Replace(sTheString, vbNullString)
        ErrorRowDOB = Right(ErrorRowDOB, 8)
        Debug.Print ("The DOB:" + ErrorRowDOB)
        GetDOB = ErrorRowDOB
    End With
    
End Function

