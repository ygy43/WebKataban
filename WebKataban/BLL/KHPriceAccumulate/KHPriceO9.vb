'************************************************************************************
'*  ProgramID  FKHPriceO9
'*  Programผ  FPฟvZTuW[
'*
'*                                      ์ฌ๚F2007/04/26   ์ฌาFNII A.Takahashi
'*                                      XV๚F             XVาF
'*
'*  Tv       FJMtcณroู   uUOPOV[Y
'*
'************************************************************************************
Module KHPriceO9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            'z๑่`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '๎{ฟiL[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'IvVมZฟiL[
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strOpArray(intLoopCnt).Trim()
                decOpAmount(UBound(decOpAmount)) = 1

            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
