'************************************************************************************
'*  ProgramID  �FKHPriceO2
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/04/18   �쐬�ҁFNII A.Tatakashi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�G�A�n�C�h���u�[�X�^ �`�g�a�V���[�Y
'*
'************************************************************************************
Module KHPriceO2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)
        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
            decOpAmount(UBound(decOpAmount)) = 1

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
