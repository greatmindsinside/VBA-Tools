Attribute VB_Name = "LawsonsCode"
Option Explicit
'Make string compares case insensitive
Option Compare Text



Public Function IsFutureDate(dDate) As Boolean
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("IsFutureDate(dDate) As Boolean")
    Debug.Print ("----------------------------------------------------")
    
    If DateDiff("d", CDate(dDate), Date) < 0 Then
        '===========================================
        ' This is a Future Date
        '===========================================
        Debug.Print ("This is a future date: " & dDate)
        '======================================
        'Function Returns A Bool Value of True
        '======================================
        IsFutureDate = True
    Else
        '===========================================
        ' This is a Date in the Past
        '===========================================
        Debug.Print ("This is a date in the past: " & dDate)
        '======================================
        'Function Returns A Bool Value of False
        '======================================
        IsFutureDate = False
    End If
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("")
    
End Function

Private Function GetDay(sYYYYMMDD) As String
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("GetDay(sYYYYMMDD) As String")
    Debug.Print ("----------------------------------------------------")
    
    Dim sTheDay As String
    '-------------------------------------------------
    'Gets the day from a string in YYYYMMDD format
    '-------------------------------------------------
    If IsNumeric(sYYYYMMDD) And Len(sYYYYMMDD) = 8 Then
        sTheDay = Right(sYYYYMMDD, 2)
        GetDay = sTheDay
    End If
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("")
    
End Function

Private Function ConvertDate(sDate) As String
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("ConvertDate(sDate) As String")
    Debug.Print ("----------------------------------------------------")
    
    Dim CONVERT_YYYYMMDD_TO_DATE As String
    
    If IsNumeric(sDate) And Len(sDate) = 8 Then
        CONVERT_YYYYMMDD_TO_DATE = DateSerial(Left(sDate, 4), _
        Mid(sDate, 5, 2), _
        Right(sDate, 2))
        
        Debug.Print ("The Date Is: " & CONVERT_YYYYMMDD_TO_DATE)
        
        '===========================================
        'Function Returns String
        '===========================================
        ConvertDate = CONVERT_YYYYMMDD_TO_DATE
    Else
        Debug.Print ("Not A Valid Date")
        ConvertDate = "NULL"
    End If
    
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("")
    
    
End Function

Public Function IsActiveSheetName(sActiveSheetName As String) As Boolean
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("IsActiveSheetName(sActiveSheetName As String) As Boolean")
    Debug.Print ("----------------------------------------------------")
    
    Dim SheetName As String: SheetName = ActiveSheet.Name
    Dim WorkBookName As String: WorkBookName = ActiveWorkbook.Name
    
    If (sActiveSheetName = ActiveSheet.Name) Then
        Debug.Print ("The SheetName: " & SheetName)
        Debug.Print ("The Workbook: " & WorkBookName)
        
        IsActiveSheetName = True
    Else
        Debug.Print ("The SheetName: " & SheetName)
        Debug.Print ("The Workbook: " & WorkBookName)
        
        IsActiveSheetName = False
    End If
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("")
    
End Function

Private Function FindInActiveSheet(sTheText As String) As Range
    
    Dim rgFound As Variant
    Dim FirstRangeFound As String
    Dim ColNumEligStartDate As Integer: ColNumEligStartDate = 10
    Dim sConDate As String
    Dim sTheDay As String
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("FindInActiveSheet(sTheText As String) As Range")
    Debug.Print ("----------------------------------------------------")
    
        '========================================================
        ' Function returns the range that the text is found in.
        '========================================================
        'If eligiblity start date is blank then send it back.  If they had never sent an eligiblity start date then it would have loaded through with out falling out for GUID segment changes
        
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
            Debug.Print ("First Range Found: " & FirstRangeFound)
            Debug.Print ("Range Found: " & rgFound)
            
            Do
                '===============================================
                ' Print the location that the data was found in
                '===============================================
                Debug.Print sTheText & " was found in Cell: " & rgFound.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False)
                'Debug.Print ("The Row Number: " & rgFound.Row)
                
                If Not rgFound Is Nothing Then
                    Debug.Print ("The Row Number: " & rgFound.Row)
                    Debug.Print ("eligiblity date: " & Range(Cells(rgFound.Row, ColNumEligStartDate).Address).Value)
                    
                    '============================
                    ' Check if the date is valid
                    '============================
                    sConDate = ConvertDate(Range(Cells(rgFound.Row, ColNumEligStartDate).Address).Value)
                    sTheDay = GetDay(Range(Cells(rgFound.Row, ColNumEligStartDate).Address).Value)
                    If sConDate = "NULL" Or sTheDay <> "01" Then
                        '========================================================
                        ' Not A Vaild Date (Highlight cell that has invalid date)
                        '========================================================
                        rgFound.EntireRow.Interior.ColorIndex = 2
                        Cells(rgFound.Row, ColNumEligStartDate).Interior.ColorIndex = 8
                        Debug.Print ("Not A Valid Date")
                        
                    Else
                        
                        '==============================================================
                        ' Date is Valid (Check if date is in the future or the past)
                        '==============================================================
                        If IsFutureDate(sConDate) = True Then
                            '===========================================
                            ' Highlight entire row
                            '===========================================
                            rgFound.EntireRow.Interior.ColorIndex = 4
                        End If
                        '=============================================
                        'Check if the date is the first of the month
                        '=============================================
                        
                        If IsFirstofTheMonth(sTheDay) = False Then
                            Cells(rgFound.Row, ColNumEligStartDate).Interior.ColorIndex = 8
                        End If
                    End If
                End If
                
                Set rgFound = .FindNext(rgFound)
            Loop While Not rgFound Is Nothing And rgFound.Address <> FirstRangeFound
            
        Else
            Debug.Print sTheText & " was not found."
        End If
        
        Debug.Print ("----------------------------------------------------")
        Debug.Print ("")
        
        
        
    End With
    
End Function

Public Function IsFirstofTheMonth(sTheDay As String) As Boolean
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print (" IsFirstofTheMonth(sTheDay As String)")
    Debug.Print ("----------------------------------------------------")
    
    '======================================================================================
    'Takes a 2 digit number from a string then checks if it is the first day of the month.
    '======================================================================================
    If Not sTheDay = "01" Then
        'MsgBox ("Not the first of the month.")
        Debug.Print ("Not The First of The Month: " & sTheDay)
        '====================================
        ' Function Returns Bool Value of False
        '====================================
        IsFirstofTheMonth = False
    Else
        '====================================
        ' Function Returns Bool Value of True
        '====================================
        Debug.Print ("The First of The Month: " & sTheDay)
        IsFirstofTheMonth = True
    End If
    
    Debug.Print ("----------------------------------------------------")
    Debug.Print ("")
    
End Function

Public Sub CampaginSegmentFinder()
    
    IsActiveSheetName ("Sheet4")
    FindInActiveSheet ("The value of column " & """campaignSegmentGuid""" & " cannot be changed once it is set")
    
End Sub





Function StripNonAlpha(TextToReplace As String) As String
    Dim ObjRegex As Object
    Set ObjRegex = CreateObject("vbscript.regexp")
    With ObjRegex
        .Global = True
        .Pattern = "[^a-zA-Z\s\.]+"
        StripNonAlpha = .Replace(Replace(TextToReplace, "-", Chr(32)), vbNullString)
    End With
End Function

Private Function CannotCreateRecord()
    Dim SeacrhStr As String: SeacrhStr = "Cannot create a record for the Primary, since a record already exists for this organization with the same first initial, last name, and date of birth."
    'Look at the SSN of the Primary its most likly off by a few digits
    
End Function





Function RequiredFieldNotProvided()
'This Issue only applies to city of memp
'Check the System Assigned File Name columns (Column 2) for the string memp
    
'Required field FirstName was not provided on Dependent person. Required field LastName was not provided on Dependent person. Required field DateOfBirth was not provided on Dependent person. Required field GenderCode was not provided on Dependent person. GenderCode was not an accepted value (M,F,U) for Dependent person. Required field ProgramEligibilityIndicator was not provided on Dependent person. Required field HasEndStageRenalDisease was not provided on Dependent person. ProgramEligibilityIndicator was not an accepted value (Y,N) for Dependent person. HasEndStageRenalDisease was not an accepted value (Y,N) for Dependent person. Required field DependentCampaignSegmentGuid was not provided on Dependent person. DependentRelationshipTypeId is required if a dependent is provided.
'Delete column Z, BM, BT, AA, BF
    
    
    
    
End Function
Public Sub SSNAlreadyInUse()

    
  
   
    
End Sub




Private Function JustFindTheText(sTheText As String) As Range
    
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
                
                
                
                Set rgFound = .FindNext(rgFound)
            Loop While Not rgFound Is Nothing And rgFound.Address <> FirstRangeFound
        End If
    End With
End Function







