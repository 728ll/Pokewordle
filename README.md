Sub pokeword()
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim i As Integer
    
    If Range("I2") = 0 Then
        Call pokemax
    Else
        Call errorchk
        Call pokesort_beta
    End If
    
    Range("I2") = Range("I2") + 1
    If Range("I2") = 10 Then
    
    End If
       
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Function pokemax() '最大値を持つポケモンの表示
    Dim i As Integer
    Dim buf As Integer
    Dim buf_cell As Integer
    Dim num As Integer
    
    num = Range("I2")
    buf = 0
    
    For i = 1 To Cells(Rows.Count, 2).End(xlUp).Row
        If buf < Cells(i, 2) Then
            buf = Cells(i, 2)
            buf_cell = i
        End If
    Next
    
    Cells(num + 2, 7) = Cells(buf_cell, 1)
        
End Function

Function pokesort_beta() '２回目以降のソート

    Dim i As Integer
    Dim moji As Integer '何文字か
    Dim buf_moji As String 'moji目のカタカナ
    Dim buf_num As String 'moji目の結果の数字

    
    For i = 2 To Cells(Rows.Count, 7).End(xlUp).Row
        For moji = 1 To 5
            buf_num = Mid(Cells(i, 8), moji, 1)
            If buf_num = "0" Then
                buf_moji = Mid(Cells(i, 7), moji, 1)
                Call pts_change(moji, buf_moji, buf_num, i)
            End If
        Next
        
        For moji = 1 To 5
            buf_num = Mid(Cells(i, 8), moji, 1)
            If buf_num = "1" Then
                buf_moji = Mid(Cells(i, 7), moji, 1)
                Call pts_change(moji, buf_moji, buf_num, i)
            End If
        Next
        
        For moji = 1 To 5
            buf_num = Mid(Cells(i, 8), moji, 1)
            If buf_num = "2" Then
                buf_moji = Mid(Cells(i, 7), moji, 1)
                Call max_sisu(buf_moji, moji)
            End If
        Next
    Next
    Call pokemax
End Function

Function pts_change(moji As Integer, buf_moji As String, buf_num As String, y As Integer) '結果に対する優先度の変更
    
    Dim i As Integer
    Dim kaburi As String
    Dim x As Integer
    Dim chk As Integer
    
    
    Select Case buf_num
        Case "0" '一致なし
        
            For x = moji To 5
                If Mid(Cells(y, 7), x, 1) = buf_moji And Mid(Cells(y, 8), x, 1) <> 0 Then
                    chk = 1
                Else
                End If
            Next
            For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
                If InStr(Cells(i, 1), buf_moji) <> 0 And chk <> 1 Then
                    Cells(i, 2) = 0 '使われない文字を持っているポケモンは指数0
                End If
            Next
               
            
        Call get_sisu
        
        Case "1"
            For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
                For x = 1 To 5
                    If Mid(Cells(i, 1), x, 1) = buf_moji Then
                        If Cells(i, 2) = 0 Then
                        Else
                            Cells(i, 2) = Cells(i, 2) + 250
                        End If
                    End If
                Next
                
                kaburi = Mid(Cells(i, 1), moji, 1)
                If kaburi = buf_moji Then
                    Cells(i, 2) = 0 '同じ場所に同じ文字があるポケモンの指数は0
                End If
                
            Next
        End Select
        
End Function
Function get_sisu() '評価指数
    
    Dim i As Integer
    Dim x As Integer
    Dim str_chk As Integer
    Dim cnt As Integer
    Dim chk As String
    Dim moji As Integer
    
    str_chk = 0
    cnt = 0
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        cnt = 0
        If Cells(i, 2) <> 0 Then '0にしたポケモンの指数は変えない
        
            Cells(i, 2) = 1 '指数リセット
            For moji = 1 To 5
                chk = Mid(Cells(i, 1), moji, 1)
                For x = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                    If chk = Cells(x, 3) Then
                        Cells(i, 2) = Cells(i, 2) + Cells(x, 5)
                    End If
                    
                Next
                
                Do '被り文字の優先度を1文字分にする
                    str_chk = InStr(moji + 1, Cells(i, 1), Mid(Cells(i, 1), moji, 1))
                    If str_chk = 0 Then
                        Exit Do
                    Else
                        cnt = cnt + 1
                        For x = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                            If Mid(Cells(i, 1), moji, 1) = Cells(x, 3) Then
                                Cells(i, 2) = Cells(i, 2) - Cells(x, 5)
                            End If
                        Next
                    End If
                    Exit Do
                Loop
            Next
        End If
    Next

        
End Function

Function max_sisu(buf_moji As String, moji As Integer) '一致文字を持つポケモンの指数上げ

    Dim i As Integer
    Dim buf As Integer
  
    buf = 0
    
    For i = 1 To Cells(Rows.Count, 2).End(xlUp).Row
                   
        If buf < Cells(i, 2) And Mid(Cells(i, 1), moji, 1) = buf_moji Then
            If Cells(i, 2) = 0 Then
            Else
                Cells(i, 2) = Cells(i, 2) + 500
            End If
        End If
    Next
End Function

Sub reset()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim i As Integer
    
    For i = 2 To Cells(Rows.Count, 7).End(xlUp).Row
        Cells(i, 7) = ""
        Cells(i, 8) = ""
    Next
    Range("I2") = 0
    
    Call mojiPts
    Call sisu_reset
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Function mojiPts() '文字数に対する優先度付
    Dim i As Integer
    Dim cnt As Integer
    Dim buf As Integer
    
    buf = 0
    cnt = 0
    
    For i = Cells(Rows.Count, 4).End(xlUp).Row To 1 Step -1
    
        If Cells(i, 4) > buf Then '１つ下の文字より使用数が多い場合
            cnt = cnt + 1
            Cells(i, 5) = cnt
            buf = Cells(i, 4)
            
        ElseIf Cells(i, 4) = buf Then '１つ下の文字と使用数が同じ場合
            Cells(i, 5) = cnt
            buf = Cells(i, 4)
            
        Else
        
        End If
        
    Next
        
        
    
End Function

Function sisu_reset() '評価指数
    
    Dim i As Integer
    Dim x As Integer
    
    Dim moji As Integer
    
        
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 2) = 0 '指数リセット
        For moji = 1 To 5
            For x = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                If Mid(Cells(i, 1), moji, 1) = Cells(x, 3) Then
                    Cells(i, 2) = Cells(i, 2) + Cells(x, 5)
                End If
            Next
        Next
    Next
                
        
End Function

Function errorchk()
    Dim i As Integer
    Dim x As Integer
    
    For i = 2 To Cells(Rows.Count, 7).End(xlUp).Row
        If Len(Cells(i, 8)) <> 5 Then
            MsgBox ("5文字の数値で入力してください")
            End
        End If
        For x = 1 To 5
            If Mid(Cells(i, 8), x, 1) >= 0 And Mid(Cells(i, 8), x, 1) <= 2 Then
            
            Else
                MsgBox ("0から2の数値で入力してください")
                End
            End If
        Next
    Next
    
End Function

