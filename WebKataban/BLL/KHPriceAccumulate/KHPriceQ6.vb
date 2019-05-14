'************************************************************************************
'*  ProgramID  �FKHPriceQ6
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/03/05   �쐬�ҁFT.Yagyu
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FSFR�ASFRT�V���[�Y
'*
'************************************************************************************
Module KHPriceQ6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strSw As String = ""

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            Dim bolC5Flag As Boolean

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
            objKtbnStrc.strcSelection.strOpSymbol(2)
            decOpAmount(UBound(decOpAmount)) = 1

            '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) START--->
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case ""
                    '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) <---END
                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5) = "R") Or (objKtbnStrc.strcSelection.strOpSymbol(5) = "L") Then
                            strSw = "S"
                        Else
                            strSw = "D"
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                        strSw & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "2"
                    Dim intSu As Integer
                    ReDim strPriceDiv(0)
                    '��{���i�L�[�p
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    'C5�`�F�b�N
                    bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(4) = "R") Or (objKtbnStrc.strcSelection.strOpSymbol(4) = "L") Then
                            strSw = "S"
                            intSu = 1
                        Else
                            strSw = "D"
                            intSu = 2
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                            objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                            strSw & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                "SW" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = intSu
                    End If

                    '�񎟓d�r���Z
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                            objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1
                    'strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5

            End Select
            '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) <---END

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

