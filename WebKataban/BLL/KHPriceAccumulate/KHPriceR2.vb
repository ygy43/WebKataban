'************************************************************************************
'*  ProgramID  �FKHPriceR2
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2010/03/26   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FSCPS�V���[�Y  (�y���V���V�����_)
'*             �FZSF�V���[�Y   (PP�^�C�v �j���[�W���C���g)
'*             �FSC3F�V���[�Y  (PP�^�C�v �X�s�[�h�R���g���[��)
'*
'************************************************************************************
Module KHPriceR2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer = 0

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '�X�g���[�N�ݒ�
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "ZSF" Or _
               objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "SC3F" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                           intStroke
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '��RM1312XXX 2013/11/28 �C��
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "ZSF" Then

                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SC3F" Then

                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

