'************************************************************************************
'*  ProgramID  FKHPriceR8
'*  Programผ  FPฟvZTuW[
'*
'*                                      ์ฌ๚F2012/04/25   ์ฌาFY.Tachi
'*                                      XV๚F             XVาF
'*
'*  Tv       FธงM[^             qodPOOO
'*
'************************************************************************************
Module KHPriceR8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer = 0
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            'z๑่`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '๎{ฟiL[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2)
            decOpAmount(UBound(decOpAmount)) = 1

            'IvVมZฟiL[
            '2016/2/18 ฤกCณ
            'If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> " " Then
            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & "-" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

            '๑dr
            If objKtbnStrc.strcSelection.strKeyKataban.ToString = "4" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
