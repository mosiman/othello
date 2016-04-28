Attribute VB_Name = "ChanD_General"
Option Explicit


'Determine if the user wishes to exit.
Public Sub ConfirmExit()
    
    Dim Reply As Integer
    Dim Buttons As Integer
    
    Buttons = vbQuestion + vbYesNo + vbDefaultButton2
    
    Reply = MsgBox("Do you wish to exit?", Buttons, "Exit")
    
    If Reply = vbYes Then
        Beep
        End
    End If
    
End Sub

'Obtain name and path of file.
Public Function GetFile(Dialog As Control) As String
    
    Dialog.Filename = ""
    Dialog.InitDir = App.Path
    Dialog.Filter = "Text Files|*.txt|All Files |*.*"
    
    Dialog.ShowOpen
    
    GetFile = Dialog.Filename
    
End Function



'Determine if string S is too long and shortens it if it is.
Public Function StringTrim(ByVal S As String) As String

    Const TRIM_NUM = 15
    
    Dim St As String
    
    St = S
    
    If Len(S) > TRIM_NUM Then
    
        St = Left$(S, 12) & "..."
        
    End If

    StringTrim = St
    
End Function

'Prompts the user for a number within the range.
Public Function WithinRange() As Integer
    Const HIGH = 10             'Not inclusive
    Const LOW = 1               'Not inclusive
    
    Dim Num As Integer
    Dim Msg As String
    
    Msg = "Number? (" & Trim$(Str$(LOW)) & "," & Trim$(Str$(HIGH)) & ")"
    
    Do
        Num = Val(InputBox$(Msg, "Num"))
    Loop While Num > HIGH Or Num < LOW
    
    WithinRange = Num
    
End Function
'Checks file type.
Public Function CheckFileType(ByVal Str As String)
    
    Const LENFILETYPE = 3
    Const FILETYPE1 = "txt"
    Const FILETYPE2 = "rec"
    Dim TempStr As String
    Dim K As Integer
    Dim PeriodCheck As Boolean
    
    PeriodCheck = False
    
    'Retrieve file type extension.
    
    For K = 1 To Len(Str)
        If PeriodCheck = True Then
            TempStr = TempStr & Mid$(Str, K, 1)
        Else
            If Mid$(Str, K, 1) = "." Then
                PeriodCheck = True
            End If
        End If
    Next K
    
    'Assign value of extension to function.
    
    If TempStr = FILETYPE1 Then
        CheckFileType = FILETYPE1
    ElseIf TempStr = FILETYPE2 Then
        CheckFileType = FILETYPE2
    End If

End Function

'Reads a record file.
Public Sub ReadFileREC(Record() As HighScore, ByVal Filename As String, ByVal RecordLen As Integer, ByVal NumRecords As Integer)
    Dim K As Integer
    Dim clearStats(1 To 6) As HighScore
    

    
    For K = 1 To NumRecords
        clearStats(K).intTime = 32767
        clearStats(K).name = ""
        clearStats(K).points = 32767
    Next K
    
    K = 1
    
    Open Filename For Random As #1 Len = RecordLen
    Do While (Not EOF(1)) And K < 7
        Get #1, K, Record(K)
        K = K + 1
    Loop
    Close #1
    
    'If there is no existing record file, create a new file with improbable values.
    
    If Record(1).points = 0 Then
        SaveFile Filename, clearStats(), 6, RecordLen
        For K = 1 To NumRecords
            Record(K) = clearStats(K)
        Next K
    End If
    

    
End Sub

Public Sub SaveFile(ByVal Filename As String, Record() As HighScore, ByVal NumRecords As Integer, ByVal RecLength As Integer)
            
    Dim X As Integer
        
    On Error GoTo ErrorSwagger
        
    Kill Filename
        
    Open Filename For Random As #1 Len = RecLength
        
    For X = 1 To NumRecords
        Put #1, X, Record(X)
    Next X
        
    Close #1
        
    Exit Sub

ErrorSwagger:
    Resume Next
    
End Sub
