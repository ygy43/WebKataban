'************************************************************************************
'*  ProgramID  �FKHPriceR1
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2010/06/24   �쐬�ҁFT.Fujiwara
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F���`�������R���كV���[�Y  
'*       �@�@�@�@�@�@�@  �R�p�q�a�P�@�@�@�@�@                                  
'*       �@�@�@�@�@�@�@  �l�R�p�q�`�P�@�@�@�@�@                                
'*       �@�@�@�@�@�@�@  �l�R�p�q�a�P�@
'*
'************************************************************************************
Module KHPriceR1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intStroke As Integer = 0
        Dim intLoopCnt As Integer
        Dim strOpArray() As String

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "3QRA1", "3QRB1"
                    '�d�l�Ȃ�(�P��)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '���ʃT�C�Y���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "MV3QRA1", "MV3QRB1"
                    '�d�l�L(�}�j�t�H�[���h)
                    '�T�u�v���[�g
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0



                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then


                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) = "+�ݻ" OrElse Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7) = "+Senser" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7)
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                            End If
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)



                            '2010/12/10 MOD RM1012055(1��VerUP:3QR�V���[�Y) START--->
                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7, 1) <> "M" Then
                                '2010/12/10 MOD RM1012055(1��VerUP:3QR�V���[�Y) <---END
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                            Dim strKt As String = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5)
                            Dim strKt2 As String = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1)
                            '�d���ڑ����Z���i�L�[
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)


                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                 strKt2.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(4)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                End If
                            End If
                            '���ʃT�C�Y���Z���i�L�[
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(5)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If

                            '���̓Z���T���Z���i�L�[
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(8)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        End If

                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 5).Trim

                    ''�}�j�z�[���h�ȊO��{���Z���i
                    'If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "8" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & objKtbnStrc.strcSelection.strOpSymbol(1) & "9"
                    '    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    'End If

                    '�ڑ����a���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim

                    End If

                    ''�d���ڑ����Z���i�L�[
                    'If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                    '    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then

                    '    Else
                    '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                    '                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                    '                                                   objKtbnStrc.strcSelection.strOpSymbol(4)
                    '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    '    End If
                    'End If

                Case "3QB1"
                    '�d�l�Ȃ�(�P��)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '���ʃT�C�Y���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '���͎d�l���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "3QE1"
                    '�d�l�Ȃ�(�P��)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�蓮���u���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If


                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '�I�v�V�������Z���i�L�[
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                          Select strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next


                Case "M3QB1"
                    '�d�l�L(�}�j�t�H�[���h)
                    '�T�u�v���[�g
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If intLoopCnt = 2 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    '���ʃT�C�Y���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = intQuantity

                    End If

                    '���͎d�l���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    End If

                Case "M3QE1", "M3QZ1"
                    '�d�l�L(�}�j�t�H�[���h)
                    '�T�u�v���[�g
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    If intLoopCnt = 2 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '�蓮���u���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    End If

                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    '�I�v�V�������Z���i�L�[
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                            strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    Next

                Case Else
                    '�d�l�L(�}�j�t�H�[���h)
                    '�T�u�v���[�g
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "H" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "ST" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    If intLoopCnt = 2 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            '2010/12/10 MOD RM1012055(1��VerUP:3QR�V���[�Y) START--->
                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7, 1) <> "M" Then
                                '2010/12/10 MOD RM1012055(1��VerUP:3QR�V���[�Y) <---END
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '�d���ڑ����Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    '���ʃT�C�Y���Z���i�L�[
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = intQuantity

                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

