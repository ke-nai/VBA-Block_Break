# VBA-Block_Break
![Block_Break](https://user-images.githubusercontent.com/66747535/100057761-6bce6980-2e6b-11eb-848b-7374994721ec.gif)
엑셀에서 VBA 매크로를 통해 실행할 수 있는 블럭 깨기 게임이다.

## 적용법
1. VBA 편집창에 들어간다.
2. 모듈이 아니라 적용할 시트의 코드 창에 아래의 코드를 모두 넣는다.
3. 매크로 직접 실행으로 Format 실행

## 코드
<details>
    <summary>코드보기</summary>

```
Function Find(a, b) '셀 확인하고 선택
    If Cells(a, b) = Cells(1, 2) And Cells(a, b) < 10 Then '선택된 색번호와 같은지
        Cells(a, b) = Cells(a, b) + 10 '선택됨
        Cells(a, b).Interior.Pattern = xlPatternGray50 '패턴 지정
        Cells(a, b).Interior.PatternColor = RGB(0, 0, 0)
        
        tmp = Find(a - 1, b) + Find(a, b - 1) + Find(a, b + 1) + Find(a + 1, b) '사방으로 확인
    End If
End Function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells(1, 1) = "=rows(" + Selection.Address + ")"
    m = Cells(1, 1) '행 크기
    Cells(1, 1) = "=columns(" + Selection.Address + ")"
    n = Cells(1, 1) '열 크기
    
    With Selection
        If m > 1 Or n > 1 Then '다중선택방지
        ElseIf Selection.Address = Cells(2, 2).Address Then
            Start '게임시작
        ElseIf 1 < .Row And .Row < 12 And 3 < .Column And .Column < 19 Then
            Click
        End If
    End With
    
    Cells(1, 21).Select '점수조작방지
End Sub

Function Click()
    Application.ScreenUpdating = False
    
    s = 0 '점수 계산용
    a = Selection.Row
    b = Selection.Column
    Cells(1, 2) = Selection.Value
    
    If Cells(1, 2) < 10 Then '새 선택
        For Each ce In Range(Cells(2, 4), Cells(11, 18)) '이전 선택 제거
            If ce.Value > 10 Then
                ce.Value = ce.Value - 10
                ce.Interior.Pattern = xlSolid
            End If
        Next
        
        If Selection <> 0 Then '빈셀 아님
            tmp = Find(a - 1, b) + Find(a, b - 1) + Find(a, b + 1) + Find(a + 1, b)
        End If
        
    Else '다시 선택
        Cells(1, 21).Select
        For Each ce In Range(Cells(2, 4), Cells(11, 18)) '선택된 타일 파괴
            If ce.Value > 10 Then '선택됨
                ce.Value = 0

                '아래쪽으로 당기기
                If ce.Row > 2 Then
                    Range(Cells(2, ce.Column), Cells(ce.Row - 1, ce.Column)).Copy Cells(3, ce.Column)
                End If
                With Cells(2, ce.Column)
                    .Value = 0
                    .Interior.Color = RGB(0, 0, 0)
                End With
        
                s = s + 1
            End If
        Next
        
        '오른쪽으로 당기기. 새 타일
        For j = 4 To 18
            If Cells(11, j) = 0 Then
                Range(Cells(2, 4), Cells(11, j - 1)).Copy Cells(2, 5)
                Range(Cells(2, 4), Cells(11, 4)).Interior.Pattern = xIPatternSolid
                Range(Cells(2, 4), Cells(11, 4)) = 0
                
                Cells(1, 3) = "=RandBetween(0,99)"
                K = 2 + (Cells(1, 3) + Cells(11, j) + j) Mod 6
                
                For i = 11 To K Step -1
                    Cells(1, 3) = "=RandBetween(0,99)"
                    Cells(i, 4) = 1 + (Cells(1, 3) + Cells(11, j) + i - K) Mod 4
                    Select Case Cells(i, 4)
                        Case 1
                            Cells(i, 4).Interior.Color = RGB(255, 100, 100)
                        Case 2
                            Cells(i, 4).Interior.Color = RGB(100, 255, 100)
                        Case 3
                            Cells(i, 4).Interior.Color = RGB(100, 100, 255)
                        Case 4
                            Cells(i, 4).Interior.Color = RGB(255, 255, 100)
                    End Select
                Next
                
                For i = K - 1 To 2 Step -1
                    Cells(i, 4).Interior.Color = RGB(0, 0, 0)
                Next
            End If
        Next
    End If
    
    Cells(5, 2) = Cells(5, 2) + 5 * s * (s - 1) ' 점수계산
    
    If Cells(5, 2) > Cells(8, 2) Then '맥스 스코어 갱신
        Cells(8.2) = Cells(5, 2)
    End If
    
    Application.ScreenUpdating = True
End Function
Sub Format()
    Application.ScreenUpdating = False
    
    Range("A1:XFD1048576").EntireRow.Clear
    Range("A1:XFD1048576").EntireColumn.Clear
    Range("T13:XFD1048576").EntireRow.Hidden = True
    Range("T13:XFD1048576").EntireColumn.Hidden = True
    
    With Range(Cells(1, 1), Cells(12, 19))
        .ColumnWidth = 4
        .RowHeight = 30
        .Interior.Color = RGB(0, 0, 0)
    End With
    With Cells(2, 2)
        .ColumnWidth = 15
        .Value = "Start"
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 100, 100)
    End With
    With Cells(4, 2)
        .Value = "Score"
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(100, 100, 255)
    End With
    With Cells(5, 2)
        .Value = 0
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 255)
    End With
    With Cells(7, 2)
        .Value = "Max Score"
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(100, 100, 255)
    End With
    With Cells(8, 2)
        .Value = 0
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 255)
    End With
    With Range(Cells(13, 1), Cells(13, 19))
        .MergeCells = True
        .RowHeight = 25
        .Interior.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Value = "두 칸 이상 연결된 타일을 터트려 점수를 얻는다. 한 번에 여러 타일을 터트릴수록 높은 점수를 얻는다."
    End With
    
    Start
End Sub

Function Start()
    Application.ScreenUpdating = False
    
    Cells(5, 2) = 0 '점수 초기화
    
    For Each ce In Range(Cells(2, 4), Cells(11, 18)) '랜덤 타일 생성
        Cells(1, 3) = "=RandBetween(1,4)"
        ce.Value = Cells(1, 3)
        Select Case ce.Value
            Case 1
                ce.Interior.Color = RGB(255, 100, 100)
            Case 2
                ce.Interior.Color = RGB(100, 255, 100)
            Case 3
                ce.Interior.Color = RGB(100, 100, 255)
            Case 4
                ce.Interior.Color = RGB(255, 255, 100)
        End Select
    Next
    
    With Range(Cells(2, 4), Cells(11, 18)) '타일 레이아웃 수정
        .Interior.Pattern = xlPatternSolid
        .NumberFormatLocal = """"""
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    
    Application.ScreenUpdating = True
End Function

Private Sub Worksheet_Activate()
    Cells(1, 21).Select
End Sub
```

</details>
