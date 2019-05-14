'************************************************************************************
'*  ProgramID  �FKHPriceP7
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/06/10   �쐬�ҁFM.Kojima
'*
'*  �T�v       �F���J�j�J���p���[�V�����_�@MCP�V���[�Y
'*
'************************************************************************************
Module KHPriceP7
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strSuiryoku As String '����
        Dim strStroke As String '�X�g���[�N
        Dim strLead As String '���[�h������

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strSuiryoku = objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            strStroke = objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            strLead = objKtbnStrc.strcSelection.strOpSymbol(5).Trim

            '��{���i�L�[�̐ݒ�
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MCP-W") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "00" & CdCst.Sign.Hypen & _
                    strSuiryoku & CdCst.Sign.Hypen & _
                    strStroke
            ElseIf (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MCP-S") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "00" & CdCst.Sign.Hypen & _
                    strSuiryoku
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'FA���Z
            If (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "FA") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "FA" & CdCst.Sign.Hypen & _
                    strSuiryoku
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '���b�h��[���˂�(N)���Z
            If (objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0) Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & _
                    strSuiryoku & _
                    CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�X�C�b�`���Z���i�L�[
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                    strSuiryoku & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '���[�h���������Z���i�L�[
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                        CdCst.Sign.Hypen & _
                        strLead
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
