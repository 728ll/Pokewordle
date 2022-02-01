# pokewordleSub pokeword()
    
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

Function pokemax() '・ｽﾅ托ｿｽl・ｽ・ｽ・ｽ・ｽﾂポ・ｽP・ｽ・ｽ・ｽ・ｽ・ｽﾌ表・ｽ・ｽ
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

Function pokesort_beta() '・ｽQ・ｽ・ｽﾚ以降・ｽﾌソ・ｽ[・ｽg

    Dim i As Integer
    Dim moji As Integer '・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ
    Dim buf_moji As String 'moji・ｽﾚのカ・ｽ^・ｽJ・ｽi
    Dim buf_num As String 'moji・ｽﾚの鯉ｿｽ・ｽﾊの撰ｿｽ・ｽ・ｽ

    
    For i = 2 To Cells(Rows.Count, 7).End(xlUp).Row
        For moji = 1 To 5
            buf_num = Mid(Cells(i, 8), moji, 1)
            If buf_num = "0" Then
                buf_moji = Mid(Cells(i, 7), moji, 1)
                Call pts_change(moji, buf_moji, buf_num)
            End If
        Next
        
        For moji = 1 To 5
            buf_num = Mid(Cells(i, 8), moji, 1)
            If buf_num = "1" Then
                buf_moji = Mid(Cells(i, 7), moji, 1)
                Call pts_change(moji, buf_moji, buf_num)
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

Function pts_change(moji As Integer, buf_moji As String, buf_num As String) '・ｽ・ｽ・ｽﾊに対ゑｿｽ・ｽ・ｽD・ｽ・ｽx・ｽﾌ変更
    
    Dim i As Integer
    Dim kaburi As String
    
    
    Select Case buf_num
        Case "0" '・ｽ・ｽv・ｽﾈゑｿｽ
            For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
                If InStr(Cells(i, 1), buf_moji) <> 0 Then
                    Cells(i, 2) = 0 '・ｽg・ｽ・ｽ・ｽﾈゑｿｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽﾄゑｿｽ・ｽ・ｽ|・ｽP・ｽ・ｽ・ｽ・ｽ・ｽﾍ指・ｽ・ｽ0
                End If
            Next
        Call get_sisu
        
        Case "1"
            For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            
                kaburi = Mid(Cells(i, 1), moji, 1)
                If kaburi = buf_moji Then
                    Cells(i, 2) = 0 '・ｽ・ｽ・ｽ・ｽ・ｽ齒奇ｿｽﾉ難ｿｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ|・ｽP・ｽ・ｽ・ｽ・ｽ・ｽﾌ指・ｽ・ｽ・ｽ・ｽ0
                End If
            Next
        End Select
        
End Function
Function get_sisu() '・ｽ]・ｽ・ｽ・ｽw・ｽ・ｽ
    
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
        If Cells(i, 2) <> 0 Then '0・ｽﾉゑｿｽ・ｽ・ｽ・ｽ|・ｽP・ｽ・ｽ・ｽ・ｽ・ｽﾌ指・ｽ・ｽ・ｽﾍ変ゑｿｽ・ｽﾈゑｿｽ
        
            Cells(i, 2) = 1 '・ｽw・ｽ・ｽ・ｽ・ｽ・ｽZ・ｽb・ｽg
            For moji = 1 To 5
                chk = Mid(Cells(i, 1), moji, 1)
                For x = 1 To Cells(Rows.Count, 3).End(xlUp).Row
                    If chk = Cells(x, 3) Then
                        Cells(i, 2) = Cells(i, 2) + Cells(x, 5)
                    End If
                    
                Next
                
                Do
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

Function max_sisu(buf_moji As String, moji As Integer) '・ｽ・ｽv・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽﾂポ・ｽP・ｽ・ｽ・ｽ・ｽ・ｽﾌ指・ｽ・ｽ・ｽ繧ｰ

    Dim i As Integer
    Dim buf As Integer
  
    buf = 0
    
    For i = 1 To Cells(Rows.Count, 2).End(xlUp).Row
                   
        If buf < Cells(i, 2) And Mid(Cells(i, 1), moji, 1) = buf_moji Then
            Cells(i, 2) = Cells(i, 2) + 500
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

Function mojiPts() '・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽﾉ対ゑｿｽ・ｽ・ｽD・ｽ・ｽx・ｽt
    Dim i As Integer
    Dim cnt As Integer
    Dim buf As Integer
    
    buf = 0
    cnt = 0
    
    For i = Cells(Rows.Count, 4).End(xlUp).Row To 1 Step -1
    
        If Cells(i, 4) > buf Then '・ｽP・ｽﾂ会ｿｽ・ｽﾌ包ｿｽ・ｽ・ｽ・ｽ・ｽ・ｽg・ｽp・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ鼾・
            cnt = cnt + 1
            Cells(i, 5) = cnt
            buf = Cells(i, 4)
            
        ElseIf Cells(i, 4) = buf Then '・ｽP・ｽﾂ会ｿｽ・ｽﾌ包ｿｽ・ｽ・ｽ・ｽﾆ使・ｽp・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ鼾・
            Cells(i, 5) = cnt
            buf = Cells(i, 4)
            
        Else
        
        End If
        
    Next
        
        
    
End Function

Function sisu_reset() '・ｽ]・ｽ・ｽ・ｽw・ｽ・ｽ
    
    Dim i As Integer
    Dim x As Integer
    
    Dim moji As Integer
    
        
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 2) = 0 '・ｽw・ｽ・ｽ・ｽ・ｽ・ｽZ・ｽb・ｽg
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
            MsgBox ("5・ｽ・ｽ・ｽ・ｽ・ｽﾌ撰ｿｽ・ｽl・ｽﾅ難ｿｽ・ｽﾍゑｿｽ・ｽﾄゑｿｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ")
            End
        End If
        For x = 1 To 5
            If Mid(Cells(i, 8), x, 1) >= 0 And Mid(Cells(i, 8), x, 1) <= 2 Then
            
            Else
                MsgBox ("0・ｽ・ｽ・ｽ・ｽ2・ｽﾌ撰ｿｽ・ｽl・ｽﾅ難ｿｽ・ｽﾍゑｿｽ・ｽﾄゑｿｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ")
                End
            End If
        Next
    Next
    
End Function

