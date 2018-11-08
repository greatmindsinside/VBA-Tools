Attribute VB_Name = "FindClientMemberID"
Option Explicit
Option Compare Text

Sub Start()

    FindClientMemberID ("FieldCannotBeUpdatedOnceSet:clientMemberId")
    FindClientMemberID ("FieldCannotBeUpdatedOnceSet:dependentClientMemberId")
    
End Sub

Private Function FindClientMemberID(sTheText As String) As Range
    
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
    
    Dim depClientMemID As String
    Dim MemID As Variant
    Dim clientMemberID As String
    Dim sPriMemIDArray() As String
    Dim PriMemID_CX As String: PriMemID_CX = "FieldCannotBeUpdatedOnceSet:clientMemberId"
    Dim ErrorRowMemID As String
    Dim CXErrorRowText As String
    
    Dim rgFound As Variant
    Dim FirstRangeFound As String
    Dim ColNumEligStartDate As Integer: ColNumEligStartDate = 10
    
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
                'Debug.Print ("The Row Number: " & rgFound.Row)
                
                clientMemberID = Cells(rgFound.Row, ClientMemCol).Value
                depClientMemID = Cells(rgFound.Row, DepClientMemCol).Value
                CXErrorRowText = Cells(rgFound.Row, SSNErrorCX).Value
                
                '-------------------------
                'Member ID Issues
                '-----------------------
                                  
                    CXErrorRowText = Trim(Cells(rgFound.Row, SSNErrorCX).Value)
                    'Primary Member ID Issue
                    'Split the text based on period and add to array
                    sPriMemIDArray = Split(CXErrorRowText, ".")
                    'Loop through array looking for the primary member ID
                    For Each MemID In sPriMemIDArray
                        If Not IsEmpty(MemID) And MemID <> "" Then
                             Debug.Print ("MemID: " & MemID)
                               'Instr returns 0 if not found
                                If InStr(1, MemID, "dependentClientMemberId", 1) <> 0 Then
                                    ErrorRowMemID = GetMemberID(MemID)
                                    'Debug.Print ("Dep Member ID: " + ErrorRowMemID)
                                    
                                    If StrComp(ErrorRowMemID, depClientMemID, 1) = 0 Then
                                        Cells(rgFound.Row, DepClientMemCol).Interior.ColorIndex = 0
                                    Else
                                        Cells(rgFound.Row, DepClientMemCol).Interior.ColorIndex = 3
                                        Cells(rgFound.Row, DepClientMemCol).Value = ErrorRowMemID
                                    End If
                                ElseIf InStr(1, MemID, "clientMemberId", 1) <> 0 Then
                                    'Primary Member ID
                                    ErrorRowMemID = GetMemberID(MemID)
                                    'Debug.Print ("Primary Member ID: " + ErrorRowMemID)
                                    
                                    If StrComp(ErrorRowMemID, clientMemberID, 1) = 0 Then
                                        Cells(rgFound.Row, ClientMemCol).Interior.ColorIndex = 0
                                    Else
                                        Cells(rgFound.Row, ClientMemCol).Interior.ColorIndex = 3
                                        Cells(rgFound.Row, ClientMemCol).Value = ErrorRowMemID
                                    End If
                                End If
                        End If
                        
                    Next
                                   
                Set rgFound = .FindNext(rgFound)
            Loop While Not rgFound Is Nothing And rgFound.Address <> FirstRangeFound
        End If
    End With
End Function

Function GetMemberID(sTheString)

    Dim ErrorRowMemID As String
    Dim ObjRegex
    Dim RealQ As String: RealQ = Chr(34)

    Set ObjRegex = CreateObject("vbscript.regexp")
    With ObjRegex
        .Global = True
        .Pattern = "\D"
        
        ErrorRowMemID = .Replace(sTheString, vbNullString)
        'Debug.Print ("The MemberID:" + ErrorRowMemID)
        GetMemberID = ErrorRowMemID
    End With

End Function


