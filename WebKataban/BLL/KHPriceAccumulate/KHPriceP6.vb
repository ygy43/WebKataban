'************************************************************************************
'*  ProgramID  �FKHPriceP6
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/06/09   �쐬�ҁFM.Kojima
'*
'*  �T�v       �F�����h�~�t�G���V�����_�@UFCD�V���[�Y
'*
'************************************************************************************
Module KHPriceP6
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strBoreSize As String           '���a
        Dim strStroke As String             '�X�g���[�N

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            strStroke = objKtbnStrc.strcSelection.strOpSymbol(3).Trim

            '�σX�g���[�N�ݒ�@
            intStroke = _
                KHKataban.fncGetStrokeSize(objKtbnStrc, _
                    CInt(strBoreSize), CInt(strStroke))

            '��{���i�L�[�̐ݒ�
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '�}�O�l�b�g����(L)���Z
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                "L"
            decOpAmount(UBound(decOpAmount)) = 1

            '�X�C�b�`���Z���i�L�[

            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '���[�h���������Z���i�L�[
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
