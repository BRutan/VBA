VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WordRootsQuizForm 
   Caption         =   "Word Roots Quiz"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5130
   OleObjectBlob   =   "WordRootsQuizForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WordRootsQuizForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Global Variables:
Dim quizSelection(), wrongAnswers() As String
Dim currQuizIndex, numQuizRoots, numCorrect As Integer
Dim contentRange As Range
Dim quizRunning As Boolean

Private Sub StartQuizButton_Click()
    If quizRunning = True Then
        ' If the quiz is already running then do nothing:
        Exit Sub
    End If
    ''''''''''''''  Check input validity: ''''''''''''''
    Dim errorMessage As String: errorMessage = "Error: "
    Dim hasError As Boolean: hasError = False
    ' Check entered number of inputs:
    If NumInput.value = "" Or NumInput.value = vbNullString Or Trim(NumInput.value & vbNullString) = vbNullString Then
        errorMessage = errorMessage + vbCr + "Please enter the number of words/roots for quiz."
        hasError = True
    ElseIf IsNumeric(NumInput.value) = False Then
        errorMessage = errorMessage + vbCr + "Number of words/roots must be an integer."
        hasError = True
    ElseIf CInt(NumInput.value) < 0 Then
        errorMessage = errorMessage + vbCr + "Number of words/roots must be greater than 0."
        hasError = True
    End If
    ' Check to see if quiz type option was selected:
    If RootQuizOptionButton.value = False And Top100WordsOptionButton.value = False Then
        errorMessage = errorMessage + vbCr + "Please select quiz type (Roots quiz or Top 100 Words quiz)."
        hasError = True
    End If
    ' Set word/root and defintion range depending on selected quiz type:
    If RootQuizOptionButton.value = True Then
        Set contentRange = ActiveSheet.Range("A2:B501")
    ElseIf Top100WordsOptionButton.value = True Then
        Set contentRange = ActiveSheet.Range("J2:K101")
    End If
    ' Check if # of inputs exceeds # of possible words:
    On Error GoTo DisplayErrors
    If CInt(NumInput.value) > contentRange.Rows.Count Then
        errorMessage = errorMessage + vbCr + "Please choose # of roots/words less than or equal to " + CStr(contentRange.Rows.Count)
        hasError = True
    End If
    ' Display message detailing error if necessary:
    If hasError = True Then
DisplayErrors:
        MsgBox errorMessage, vbOKOnly
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''  Set Globals: ''''''''''''''''''''''''''''''''''''''''''
    quizRunning = True
    ReDim wrongAnswers(2, 1)
    currQuizIndex = 1
    numQuizRoots = CInt(NumInput.value)
    ReDim quizSelection(numQuizRoots, 2)
    ''''''''''''''''''''''''''''''''''''''''''  Clean-up ''''''''''''''''''''''''''''''''''''''''''
    CurrentRootBox.Text = ""
    CurrentRootBox.MaxLength = 40
    CurrentRootBox.BackColor = &H80000005
    FirstAnswer.Caption = ""
    FirstAnswer.Locked = False
    SecondAnswer.Caption = ""
    SecondAnswer.Locked = False
    ThirdAnswer.Caption = ""
    ThirdAnswer.Locked = False
    FourthAnswer.Caption = ""
    FourthAnswer.Locked = False
    OutputIncorrectRange.Text = ""
    OutputIncorrectRange.BackColor = &H80000006
    PercentCorrectBox.Text = ""
    PercentCorrectBox.BackColor = &H80000006
    NumInput.Text = ""
    NumInput.BackColor = &H80000006
    '''''''''''''' Get the roots and definitions from active sheet: ''''''''''''''
    Dim RootsDefs() As String
    ReDim RootsDefs(contentRange.Rows.Count, 2)
    
    Dim i As Integer
    For i = 1 To contentRange.Rows.Count
        RootsDefs(i, 1) = contentRange.Cells(i, 1)
        RootsDefs(i, 2) = contentRange.Cells(i, 2)
    Next i
    
    ' Set the quiz selection based upon number of desired roots and random selection of lines in
    ' the contentRange:
    Dim selectedRootIndices As New IntVector
    Dim randNum, currQuizSelectIndex As Integer: currQuizSelectIndex = 1: i = 1
    
    If numQuizRoots = contentRange.Rows.Count Then
        For i = 1 To contentRange.Rows.Count
            quizSelection(i, 1) = RootsDefs(i, 1)
            quizSelection(i, 2) = RootsDefs(i, 2)
        Next i
    Else
        Do While currQuizSelectIndex <> numQuizRoots
            Randomize
            randNum = Int(numQuizRoots * Rnd + 1)
            If selectedRootIndices.Find(Int(randNum)) = False Then
                selectedRootIndices.Push Int(randNum)
                quizSelection(currQuizSelectIndex, 1) = contentRange.Cells(Int(randNum), 1)
                quizSelection(currQuizSelectIndex, 2) = contentRange.Cells(Int(randNum), 2)
                currQuizSelectIndex = currQuizSelectIndex + 1
            End If
        Loop
    End If
    
    ' Shuffle the quizSelection array:
    Dim tempRootDef(1, 1) As String
    Dim randNum2 As Integer
    Dim k As Integer
    For k = 1 To 3
        For i = 1 To numQuizRoots
            Randomize
            randNum = Int(numQuizRoots * Rnd + 1)
            Randomize
            randNum2 = Int(numQuizRoots * Rnd + 1)
            
            tempRootDef(1, 0) = quizSelection(Int(randNum), 1)
            tempRootDef(1, 1) = quizSelection(Int(randNum), 2)
            
            ' Swap the randomly selection roots with definition:
            quizSelection(Int(randNum), 1) = quizSelection(Int(randNum2), 1)
            quizSelection(Int(randNum), 2) = quizSelection(Int(randNum2), 2)
            quizSelection(Int(randNum2), 1) = tempRootDef(1, 0)
            quizSelection(Int(randNum2), 2) = tempRootDef(1, 1)
        Next i
    Next k
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''' Load the first root (then let the SubmitButton_Click handler take over): ''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim j As Integer: j = 1
    Dim definitions As New StringVector
    ' Update the currQuizIndex:
    currQuizIndex = 1
    ' Initially populate the box displaying the root:
    CurrentRootBox.Text = quizSelection(currQuizIndex, 1) + "(" + Trim(Str(currQuizIndex)) + "/" + Trim(Str(numQuizRoots)) + ")"
    ' Randomly select definitions from elsewhere in quizSelection:
    Dim currStr As String
    currStr = quizSelection(currQuizIndex, 2)
    definitions.Push currStr
    Do While j < 4
        Randomize
        randNum = Int(numQuizRoots * Rnd + 1)
        currStr = quizSelection(randNum, 2)
        If definitions.Find(currStr) = False And StrComp(currStr, quizSelection(j, 2)) <> 0 Then
            definitions.Push (currStr)
            j = j + 1
        End If
    Loop
    definitions.Shuffle (10)
    ' Fill in contents of answer radio buttons:
    FirstAnswer.Caption = definitions.Pop
    SecondAnswer.Caption = definitions.Pop
    ThirdAnswer.Caption = definitions.Pop
    FourthAnswer.Caption = definitions.Pop
    ' Disable the Start Quiz button until the quiz is over:
    ' StartQuizButton.Locked = True
End Sub
Private Sub SubmitButton_Click()
    ' If quiz isn't running then do nothing:
    If quizRunning = False Then
        Exit Sub
    End If
    ' Handle if no answer is selected:
    If FirstAnswer.value = False And SecondAnswer.value = False And ThirdAnswer.value = False And FourthAnswer.value = False Then
        MsgBox "Error: Select an answer."
        Exit Sub
    End If
    If currQuizIndex <> numQuizRoots + 1 Then
        ' Check answer validitiy:
        If FirstAnswer.value = True And StrComp(FirstAnswer.Caption, quizSelection(currQuizIndex, 2)) = 0 Then
            ' Increase correct answer count:
            numCorrect = numCorrect + 1
        ElseIf SecondAnswer.value = True And StrComp(SecondAnswer.Caption, quizSelection(currQuizIndex, 2)) = 0 Then
            ' Increase correct answer count:
            numCorrect = numCorrect + 1
        ElseIf ThirdAnswer.value = True And StrComp(ThirdAnswer.Caption, quizSelection(currQuizIndex, 2)) = 0 Then
            ' Increase correct answer count:
            numCorrect = numCorrect + 1
        ElseIf FourthAnswer.value = True And StrComp(FourthAnswer.Caption, quizSelection(currQuizIndex, 2)) = 0 Then
            ' Increase correct answer count:
            numCorrect = numCorrect + 1
        Else
            ' Display that choice was incorrect, using timer:
            CurrentRootBox.Text = quizSelection(currQuizIndex, 1) + " INCORRECT (" + Trim(quizSelection(currQuizIndex, 2)) + ")"
            ' Add incorrect answer to the wrongAnswer array:
            ReDim Preserve wrongAnswers(2, UBound(wrongAnswers) + 1)
            wrongAnswers(1, UBound(wrongAnswers, 2)) = quizSelection(currQuizIndex, 1)
            wrongAnswers(2, UBound(wrongAnswers, 2)) = quizSelection(currQuizIndex, 2)
            Application.Wait (Now + #12:00:04 AM#)
        End If
    Else
        ' Display number of correct answers:
        PercentCorrectBox.Text = "Correct:" + Trim(Str(numCorrect)) + "/" + Trim(Str(numQuizRoots))
        ' End quiz:
        quizRunning = False
        Exit Sub
    End If
    ' Unselect answer:
    FirstAnswer.value = False
    SecondAnswer.value = False
    ThirdAnswer.value = False
    FourthAnswer.value = False
    ' Load next root:
    currQuizIndex = currQuizIndex + 1
    CurrentRootBox.Text = quizSelection(currQuizIndex, 1) + "(" + Trim(Str(currQuizIndex)) + "/" + Trim(Str(numQuizRoots)) + ")"
    Dim definitions As New StringVector
    ' Randomly select other definitions:
    Dim j, randNum As Integer
    Dim currStr As String
    
    currStr = quizSelection(currQuizIndex, 2)
    definitions.Push (currStr)
    j = 1
    Do While j < 4
        Randomize
        randNum = Int(numQuizRoots * Rnd + 1)
        currStr = quizSelection(randNum, 2)
        If definitions.Find(currStr) = False And StrComp(currStr, quizSelection(j, 2)) <> 0 Then
            definitions.Push (currStr)
            j = j + 1
        End If
    Loop
    definitions.Shuffle (10)
    
    ' Load possible answers:
    FirstAnswer.Caption = definitions.Pop
    SecondAnswer.Caption = definitions.Pop
    ThirdAnswer.Caption = definitions.Pop
    FourthAnswer.Caption = definitions.Pop
    
End Sub
Private Sub FirstAnswer_Click()
    ' Do nothing if quiz is not running:
    If quizRunning = False Then
        FirstAnswer.value = False
        Exit Sub
    End If
    ' Unclick all other answers:
    SecondAnswer.value = False
    ThirdAnswer.value = False
    FourthAnswer.value = False
End Sub
Private Sub SecondAnswer_Click()
    ' Do nothing if quiz is not running:
    If quizRunning = False Then
        SecondAnswer.value = False
        Exit Sub
    End If
    ' Unclick all other answers:
    FirstAnswer.value = False
    ThirdAnswer.value = False
    FourthAnswer.value = False
End Sub
Private Sub ThirdAnswer_Click()
    ' Do nothing if quiz is not running:
    If quizRunning = False Then
        ThirdAnswer.value = False
        Exit Sub
    End If
    ' Unclick all other answers:
    FirstAnswer.value = False
    SecondAnswer.value = False
    FourthAnswer.value = False
End Sub
Private Sub FourthAnswer_Click()
    ' Do nothing if quiz is not running:
    If quizRunning = False Then
        FourthAnswer.value = False
        Exit Sub
    End If
    ' Unclick all other answers:
    FirstAnswer.value = False
    SecondAnswer.value = False
    ThirdAnswer.value = False
End Sub
Private Sub ExitButton_Click()
    Me.Hide
    Unload Me
End Sub
Private Sub OutputWrongButton_Click()
    ' Do nothing if quiz is running:
    If quizRunning = True Then
        Exit Sub
    End If
    ' Handle no output range:
    If OutputIncorrectRange.value = "" Or OutputIncorrectRange.value = vbNullString Then
        MsgBox "Error: Must provide output range to output incorrect words/roots."
        Exit Sub
    End If
    ' Check if desired output range intersects with the rootlist or the wordlist:
    If Application.Intersect(OutputIncorrectRange, ActiveSheet.Range("A1:B501")) Is Not Nothing Then
        MsgBox "Error: Output range intersects with Root List range. Choose another output range. "
        Exit Sub
    End If
    If Application.Intersect(OutputIncorrectRange, ActiveSheet.Range("J1:K109")) Is Not Nothing Then
        MsgBox "Error: Output range intersects with Top 100 Words range. Choose another output range. "
        Exit Sub
    End If
    
    ' Output the incorrect roots:
    Application.ScreenUpdating = False
    Dim outputRange As Range: Set outputRange = Range(OutputIncorrectRange.value)
    ' Resize the output range:
    Set outputRange = outputRange.Resize(UBound(wrongAnswers, 2), 2)
    ' Output the wrong answers:
    Dim i As Integer
    For i = 1 To UBound(wrongAnswers, 2)
        outputRange.Cells(i, 1) = wrongAnswers(1, i)
        outputRange.Cells(i, 2) = wrongAnswers(2, i)
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub Userform_Activate()
    CurrentRootBox.BackColor = &H80000006
    PercentCorrectBox.BackColor = &H80000006
    OutputIncorrectRange.BackColor = &H80000006
    
    FirstAnswer.value = False
    FirstAnswer.Locked = True
    SecondAnswer.value = False
    SecondAnswer.Locked = True
    ThirdAnswer.value = False
    ThirdAnswer.Locked = True
    FourthAnswer.value = False
    FourthAnswer.Locked = True
    
End Sub
Private Sub RootQuizOptionButton_Click()
    Top100WordsOptionButton.value = False
End Sub
Private Sub Top100WordsOptionButton_Click()
    RootQuizOptionButton.value = False
End Sub
