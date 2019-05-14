'************************************************************************************
'*  ProgramID  �FKHPriceP0
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/05/22   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F���s�t���[   �e�r�l�Q�E�e�r�l�R�V���[�Y
'*
'************************************************************************************
Module KHPriceP0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'RM1802023_FSM3�V���[�Y�ǉ�
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim

                Case "FSM3"
                    '��{���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim

                    decOpAmount(UBound(decOpAmount)) = 1

                    '�o���u�I�v�V�������Z���i�L�[
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4", "5", "6"  '�X�e�����X�{�f�B
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim & "-SUS"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else           '����
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '�P�[�u�����Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '��t�A�^�b�`�����g���Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�Y�t���މ��Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�N���[���d�l���Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(13).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "FSM2"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            '��{���i�L�[
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1

                            '�P�[�u�����Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                                    Case "N", "P"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                  "N/P" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "A"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "A" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�u���P�b�g���Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�g���[�T�r���e�B���Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�j�[�h���ٕt�����Z���i�L�[
                            '��)"FSM2-N-U2L-H","FSM2-N-O5L-H","FSM2-N-S"
                            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                '�ڑ����a(�{�f�B�ގ�)����
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    '�{�f�B�������ގ��̏ꍇ
                                    Case "H04", "H06", "H08", "H10", "H08"
                                        '���ʃ����W����
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            '�Q���b�g���^�b�ȉ��̏ꍇ
                                            Case "005", "010", "020"
                                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-U2L-H"
                                                '�T���b�g���^�b�ȏ�̏ꍇ
                                            Case Else
                                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-O5L-H"
                                        End Select
                                        '�{�f�B���X�e�����X�ގ��̏ꍇ
                                    Case "S06", "S08"
                                        strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-S"
                                    Case Else
                                End Select

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�N���[���d�l���Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "D"
                            '��{���i�L�[
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            '�P�[�u�����Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(2).Trim)
                                    Case "N", "P"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                  "N/P" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "A"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "A" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�u���P�b�g���Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�N���[���d�l���Z���i�L�[
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-D-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
