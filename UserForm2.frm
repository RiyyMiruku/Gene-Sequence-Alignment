VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4176
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **比對按鈕 - 使用 Needleman-Wunsch 算法**
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
    
    '清除舊對比資料
    
    ' 取得使用者輸入的序列，並清理數字/空格
    seq1 = UCase(RemoveNonLetters(Me.TextBox1.Text))
    seq2 = UCase(RemoveNonLetters(Me.TextBox2.Text))

    ' 確保輸入不為空
    If seq1 = "" Or seq2 = "" Then
        MsgBox "請輸入有效的基因或蛋白質序列！", vbExclamation
        Exit Sub
    End If

    ' 設定比對參數
    matchScore = 1
    mismatchPenalty = -1
    gapPenalty = -1

    ' 執行比對
    Call AlignSequences(seq1, seq2, matchScore, mismatchPenalty, gapPenalty, alignmentA, alignmentB)
    
    minLength = Len(alignmentA)
    matchCount = 0
    differences = "差異位置：" & vbCrLf

    ' 初始化 Word 文件操作
    'Set doc = ActiveDocument
    'Set rng = doc.Range(Start:=doc.Content.End - 1, End:=doc.Content.End)

    ' 比較序列並高亮差異
    
    For i = 1 To minLength
        If Mid(alignmentA, i, 1) = Mid(alignmentB, i, 1) Then
            matchCount = matchCount + 1
        Else
            differences = differences & "位置 " & i & ": " & _
                          Mid(alignmentA, i, 1) & " ≠ " & Mid(alignmentB, i, 1) & vbCrLf
            ' 在 Word 文件中標示不同的部分
            'rng.InsertAfter Mid(alignmentA, i, 1)
            'rng.Font.ColorIndex = wdRed
        End If

        ' 更新進度條
        Me.LabelProgress.Width = (i / minLength) * 200
        DoEvents
    Next i

    ' 計算相似度
    similarity = (matchCount / minLength) * 100

    ' 顯示結果
    Me.TextBox_Result.Text = "相似度：" & Format(similarity, "0.00") & "%" & vbCrLf & _
                        "比對長度：" & minLength & " 個字元" & vbCrLf & _
                        IIf(matchCount = minLength, "完全一致！", differences)

    ' 顯示結果於可捲動的 TextBox，對齊結果
    Me.TextBox3.Text = "最佳比對結果：" & vbCrLf & _
                             AlignTextForDisplay(alignmentA) & vbCrLf & _
                             AlignTextForDisplay(alignmentB)
End Sub

' **Needleman-Wunsch 比對演算法**
Sub AlignSequences(seq1 As String, seq2 As String, matchScore As Integer, mismatchPenalty As Integer, gapPenalty As Integer, alignmentA As String, alignmentB As String)
    Dim m As Integer, n As Integer
    Dim score() As Integer, traceback() As String
    Dim i As Integer, j As Integer
    Dim scoreDiag As Integer, scoreUp As Integer, scoreLeft As Integer

    m = Len(seq1)
    n = Len(seq2)

    ' 初始化矩陣
    ReDim score(0 To m, 0 To n)
    ReDim traceback(0 To m, 0 To n)

    For i = 0 To m
        score(i, 0) = i * gapPenalty
        traceback(i, 0) = "U" ' 向上
    Next i

    For j = 0 To n
        score(0, j) = j * gapPenalty
        traceback(0, j) = "L" ' 向左
    Next j

    traceback(0, 0) = "E" ' 結束標記

    ' 填充矩陣（不使用 WorksheetFunction）
    For i = 1 To m
        For j = 1 To n
            If Mid(seq1, i, 1) = Mid(seq2, j, 1) Then
                scoreDiag = score(i - 1, j - 1) + matchScore
            Else
                scoreDiag = score(i - 1, j - 1) + mismatchPenalty
            End If

            scoreUp = score(i - 1, j) + gapPenalty
            scoreLeft = score(i, j - 1) + gapPenalty

            ' 取最大值（手動計算）
            If scoreDiag >= scoreUp And scoreDiag >= scoreLeft Then
                score(i, j) = scoreDiag
                traceback(i, j) = "D" ' 對角線
            ElseIf scoreUp >= scoreLeft Then
                score(i, j) = scoreUp
                traceback(i, j) = "U" ' 向上
            Else
                score(i, j) = scoreLeft
                traceback(i, j) = "L" ' 向左
            End If
        Next j
    Next i

    ' 追溯矩陣，構建比對結果
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

' **函數：清理輸入，只保留 A-Z**
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

' **函數：對齊顯示的結果，添加空格以使每列對齊**
Function AlignTextForDisplay(ByVal inputStr As String) As String
    Dim alignedStr As String
    Dim i As Integer
    
    ' 使用適當的空格來對齊
    alignedStr = ""
    For i = 1 To Len(inputStr)
        alignedStr = alignedStr & Mid(inputStr, i, 1) & " " ' 每個字元後面加一個空格
    Next i

    AlignTextForDisplay = alignedStr
End Function

Private Sub CommandButton2_Click()

    Me.TextBox1.Text = StrReverse(Me.TextBox1.Text)

End Sub

Private Sub CommandButton3_Click()

    Me.TextBox2.Text = StrReverse(Me.TextBox2.Text)

End Sub
