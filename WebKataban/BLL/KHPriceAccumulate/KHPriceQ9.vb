'************************************************************************************
'*  ProgramID  �FKHPriceQ9
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/12/18   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FEXA�V���[�Y  (���k��C�p�@�p�C���b�g���Q�|�[�g�d���� ���`�G�A�u���[�o���u)
'*             �FGEXA�V���[�Y (���k��C�p�@�p�C���b�g���Q�|�[�g�d���� �}�j�z�[���h)
'*
'************************************************************************************
Module KHPriceQ9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "GEXA" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '�V�[���ގ����Z
                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "0" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2)
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

                '�R�C���I�v�V�������Z
                '�P�[�u������
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    'RM1612033 ���i�ǉ��Ή��̂��߁Acase�����Ɂu3�v��ǉ�  2016/12/19 �ǉ� ����
                    Case "1", "F", "3"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case ""
                            Case "2C"
                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Equals("1") Then
                                    'AC100V�̂݉��Z
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V �͉��Z����
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                        '���̑��I�v�V�������Z
                        If Not objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("") Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "2"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case ""
                            Case "2C"
                                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("1") Then
                                    'AC100V�̂݉��Z
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V �͉��Z����
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select

                ''���̑��I�v�V�������Z
                'If Not objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("") Then
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                '                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                '    decOpAmount(UBound(decOpAmount)) = 1
                'End If

                '�H�i�����H������
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

            End If

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "GEXA" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4)

                decOpAmount(UBound(decOpAmount)) = 1

                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "2C" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "1" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(3)
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(3)
                End If

            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
