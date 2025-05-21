VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4176
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **�����s - �ϥ� Needleman-Wunsch ��k**
Private Sub CommandButton1_Click()
    Dim seq1 As String, seq2 As String
    Dim alignmentA As String, alignmentB As String
    Dim matchScore As Integer, mismatchPenalty As Integer, gapPenalty As Integer
    
    Dim minLength As Long
    Dim i As Long
    Dim matchCount As Long
    Dim differences As String
    Dim similarity As Double
    'Dim doc As Document
    'Dim rng As Range
    
    '�M���¹����
    
    ' ���o�ϥΪ̿�J���ǦC�A�òM�z�Ʀr/�Ů�
    seq1 = UCase(RemoveNonLetters(Me.TextBox1.Text))
    seq2 = UCase(RemoveNonLetters(Me.TextBox2.Text))

    ' �T�O��J������
    If seq1 = "" Or seq2 = "" Then
        MsgBox "�п�J���Ī���]�γJ�ս�ǦC�I", vbExclamation
        Exit Sub
    End If

    ' �]�w���Ѽ�
    matchScore = 1
    mismatchPenalty = -1
    gapPenalty = -1

    ' ������
    Call AlignSequences(seq1, seq2, matchScore, mismatchPenalty, gapPenalty, alignmentA, alignmentB)
    
    minLength = Len(alignmentA)
    matchCount = 0
    differences = "�t����m�G" & vbCrLf

    ' ��l�� Word ���ާ@
    'Set doc = ActiveDocument
    'Set rng = doc.Range(Start:=doc.Content.End - 1, End:=doc.Content.End)

    ' ����ǦC�ð��G�t��
    
    For i = 1 To minLength
        If Mid(alignmentA, i, 1) = Mid(alignmentB, i, 1) Then
            matchCount = matchCount + 1
        Else
            differences = differences & "��m " & i & ": " & _
                          Mid(alignmentA, i, 1) & " �� " & Mid(alignmentB, i, 1) & vbCrLf
            ' �b Word ��󤤼Хܤ��P������
            'rng.InsertAfter Mid(alignmentA, i, 1)
            'rng.Font.ColorIndex = wdRed
        End If

        ' ��s�i�ױ�
        Me.LabelProgress.Width = (i / minLength) * 200
        DoEvents
    Next i

    ' �p��ۦ���
    similarity = (matchCount / minLength) * 100

    ' ��ܵ��G
    Me.TextBox_Result.Text = "�ۦ��סG" & Format(similarity, "0.00") & "%" & vbCrLf & _
                        "�����סG" & minLength & " �Ӧr��" & vbCrLf & _
                        IIf(matchCount = minLength, "�����@�P�I", differences)

    ' ��ܵ��G��i���ʪ� TextBox�A������G
    Me.TextBox3.Text = "�̨Τ�ﵲ�G�G" & vbCrLf & _
                             AlignTextForDisplay(alignmentA) & vbCrLf & _
                             AlignTextForDisplay(alignmentB)
End Sub

' **Needleman-Wunsch ���t��k**
Sub AlignSequences(seq1 As String, seq2 As String, matchScore As Integer, mismatchPenalty As Integer, gapPenalty As Integer, alignmentA As String, alignmentB As String)
    Dim m As Integer, n As Integer
    Dim score() As Integer, traceback() As String
    Dim i As Integer, j As Integer
    Dim scoreDiag As Integer, scoreUp As Integer, scoreLeft As Integer

    m = Len(seq1)
    n = Len(seq2)

    ' ��l�Ưx�}
    ReDim score(0 To m, 0 To n)
    ReDim traceback(0 To m, 0 To n)

    For i = 0 To m
        score(i, 0) = i * gapPenalty
        traceback(i, 0) = "U" ' �V�W
    Next i

    For j = 0 To n
        score(0, j) = j * gapPenalty
        traceback(0, j) = "L" ' �V��
    Next j

    traceback(0, 0) = "E" ' �����аO

    ' ��R�x�}�]���ϥ� WorksheetFunction�^
    For i = 1 To m
        For j = 1 To n
            If Mid(seq1, i, 1) = Mid(seq2, j, 1) Then
                scoreDiag = score(i - 1, j - 1) + matchScore
            Else
                scoreDiag = score(i - 1, j - 1) + mismatchPenalty
            End If

            scoreUp = score(i - 1, j) + gapPenalty
            scoreLeft = score(i, j - 1) + gapPenalty

            ' ���̤j�ȡ]��ʭp��^
            If scoreDiag >= scoreUp And scoreDiag >= scoreLeft Then
                score(i, j) = scoreDiag
                traceback(i, j) = "D" ' �﨤�u
            ElseIf scoreUp >= scoreLeft Then
                score(i, j) = scoreUp
                traceback(i, j) = "U" ' �V�W
            Else
                score(i, j) = scoreLeft
                traceback(i, j) = "L" ' �V��
            End If
        Next j
    Next i

    ' �l���x�}�A�c�ؤ�ﵲ�G
    alignmentA = ""
    alignmentB = ""
    i = m
    j = n

    Do While i > 0 Or j > 0
        Select Case traceback(i, j)
            Case "D"
                alignmentA = Mid(seq1, i, 1) & alignmentA
                alignmentB = Mid(seq2, j, 1) & alignmentB
                i = i - 1
                j = j - 1
            Case "U"
                alignmentA = Mid(seq1, i, 1) & alignmentA
                alignmentB = "_" & alignmentB
                i = i - 1
            Case "L"
                alignmentA = "_" & alignmentA
                alignmentB = Mid(seq2, j, 1) & alignmentB
                j = j - 1
        End Select
    Loop
End Sub

' **��ơG�M�z��J�A�u�O�d A-Z**
Function RemoveNonLetters(ByVal inputStr As String) As String
    Dim i As Integer, outputStr As String
    outputStr = ""

    For i = 1 To Len(inputStr)
        If Mid(inputStr, i, 1) Like "[A-Za-z]" Then
            outputStr = outputStr & UCase(Mid(inputStr, i, 1))
        End If
    Next i

    RemoveNonLetters = outputStr
End Function

' **��ơG�����ܪ����G�A�K�[�Ů�H�ϨC�C���**
Function AlignTextForDisplay(ByVal inputStr As String) As String
    Dim alignedStr As String
    Dim i As Integer
    
    ' �ϥξA�����Ů�ӹ��
    alignedStr = ""
    For i = 1 To Len(inputStr)
        alignedStr = alignedStr & Mid(inputStr, i, 1) & " " ' �C�Ӧr���᭱�[�@�ӪŮ�
    Next i

    AlignTextForDisplay = alignedStr
End Function

Private Sub CommandButton2_Click()

    Me.TextBox1.Text = StrReverse(Me.TextBox1.Text)

End Sub

Private Sub CommandButton3_Click()

    Me.TextBox2.Text = StrReverse(Me.TextBox2.Text)

End Sub
