'************************************************************************************
'*  ProgramID  �FKHPriceO3
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/04/17   �쐬�ҁFNII A.Tatakashi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�ሳ�d�󃌃M�����[�^�@�d�u�k�V���[�Y
'*
'************************************************************************************
Module KHPriceO3

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
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '�P�[�u���I�v�V�������Z���i�L�[
            If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�r�C�I�v�V�������Z���i�L�[
            If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�u���P�b�g�I�v�V�������Z���i�L�[
            If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
