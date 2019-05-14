'************************************************************************************
'*  ProgramID  �FKHPriceO7
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/04/26   �쐬�ҁFNII A.Takahashi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�K�C�h�t�V�����_�@�r�s�l�V���[�Y 
'*
'************************************************************************************
Module KHPriceO7

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'RM1404033 2014/04/11 �V�����_�X�C�b�`C5���Z
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            '�X�C�b�`���i���Z�L�[
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
            End If

            '���[�h�����i���Z�L�[
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "3"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
            End Select


            '�I�v�V�������Z���i�L�[
            'RM1307003 2013/07/04 �ǉ�
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Next
            End If

            '�N���[���d�l
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
