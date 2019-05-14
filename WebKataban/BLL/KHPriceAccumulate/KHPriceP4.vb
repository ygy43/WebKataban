'************************************************************************************
'*  ProgramID  �FKHPriceP4
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/12/18   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F���W���[���N�[�����g�o���u   �f�b�u�d�Q�E�f�b�u�r�d�Q�V���[�Y
'*�@�X�V�����@�@�F
'*�@�@�@�@�@�@�@�@�I�v�V����B�i��t�j�̒ǉ�      RM0912039 2009/12/17 Y.Miura
'************************************************************************************
Module KHPriceP4

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intStation As Integer
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '�A��(���W���[�����Z)���i�L�[
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "A", "B"
                    intStation = 1
                Case Else
                    intStation = CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim)

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
            decOpAmount(UBound(decOpAmount)) = intStation

            '�R�C���I�v�V�������Z���i�L�[
            If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = intStation
            End If

            '���̑��I�v�V�������Z���i�L�[
            'RM0912039 2009/12/17 Y.Miura �I�v�V����B�i��t�j�ǉ�
            'If Len(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <> 0 Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            '                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
            '    decOpAmount(UBound(decOpAmount)) = intStation
            'End If
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim

                        'If intStation >= 3 And strOpArray(intLoopCnt).Trim.Equals("B") Then
                        '    decOpAmount(UBound(decOpAmount)) = 2
                        'Else
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        'End If

                        '��RM1303003 2013/03/04 Y.Tachi
                        decOpAmount(UBound(decOpAmount)) = 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "B"
                                If intStation >= 3 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "S"
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "A" Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "B" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                End If
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
