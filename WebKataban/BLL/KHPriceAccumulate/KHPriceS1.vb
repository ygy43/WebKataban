'************************************************************************************
'*  ProgramID  ÅFKHPriceS1
'*  Programñº  ÅFíPâøåvéZÉTÉuÉÇÉWÉÖÅ[Éã
'*
'*                                      çÏê¨ì˙ÅF2012/09/27   çÏê¨é“ÅFY.Tachi
'*                                      çXêVì˙ÅF             çXêVé“ÅF
'*
'*  äTóv       ÅFÉyÉìÉVÉãÉVÉäÉìÉ_ï°ìÆå`Å@ÇrÇbÇoÇcÇRÅ^ÇrÇbÇoÇcÇRÅ|Çk
'*
'************************************************************************************
Module KHPriceS1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intIndex As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionP4 As Boolean = False



        Try

            'îzóÒíËã`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)


            'ÉXÉgÉçÅ[ÉNéÊìæ
            'intStroke = objPrice.fncGetStrokeSize(objKtbnStrc.strcSelection.strSeriesKataban, _
            '                                      objKtbnStrc.strcSelection.strKeyKataban, _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                Case "SCPS3" 'SCPS3
                    Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1)
                        Case "M" 'SCPS3-M
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 101
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-105"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "SCPD3" 'SCPD3-D
                    Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1)
                        Case "D"
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 101
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "T" 'SCPD3-T
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 101 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-135"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 136 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-150"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-165"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 166 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-180"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-195"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 196 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 210
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-210"
                                    End Select
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-225"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 226 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-240"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-255"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "O" 'SCPD3-O
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-105"
                                    End Select
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-135"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 136 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-150"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-165"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 166 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-180"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-195"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 196 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 210
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-210"
                                    End Select
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-225"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 226 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-240"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-255"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "M", "Z", "K" 'SCPD3-M
                            Select Case True
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-135"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 136 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-150"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-165"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 166 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-180"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-195"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 196 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 210
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-210"
                                    End Select
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-225"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 226 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-240"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-255"
                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "F", "L"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "C" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "6"
                                        Select Case True
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 15
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-15"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-30"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-45"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-60"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 70
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-70"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 71 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 80
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-80"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 81 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 90
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-90"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 91
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-100"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        '2013/11/06 âøäiÉLÅ[èCê≥
                                        Select Case True
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 15
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-15"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-30"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-45"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-60"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 75
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-75"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 90
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-90"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-100"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 110
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-110"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 111 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 120
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-120"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 130
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-130"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 131 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 140
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-140"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 141 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 150
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-150"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 160
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-160"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 161 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 170
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-170"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 171 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 180
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-180"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 190
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-190"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 191 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 200
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-200"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 201 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 210
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-210"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 220
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-220"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 221 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 230
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-230"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 231 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 240
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-240"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 250
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-250"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 251
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-260"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "6"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-105"
                                        End Select
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 135
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-135"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 136 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 150
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-150"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 165
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-165"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 166 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 180
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-180"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 195
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-195"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 196 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 210
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "10"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-200"
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-210"
                                        End Select
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 225
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-225"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 226 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 240
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-240"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 255
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-255"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 256 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 260
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-260"
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 261
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-270"
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1

                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim 'É`ÉÖÅ[Éuì‡åaîªíË
                                    Case "6"
                                        Select Case True
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                                                'ÉXÉgÉçÅ[ÉNÇPÇOÅ`ÇRÇO
                                                'ó·"SCPD3-F-6-STR10-30"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR10-30"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                                                'ÉXÉgÉçÅ[ÉNÇRÇPÅ`ÇUÇO
                                                'ó·"SCPD3-F-6-STR31-60"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR31-60"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61
                                                'ÉXÉgÉçÅ[ÉNÇUÇPÅ`ÇPÇOÇO
                                                'ó·"SCPD3-F-6-STR61-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR61-100"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "10"
                                        Select Case True
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                                'ÉXÉgÉçÅ[ÉNÇPÇOÅ`ÇSÇT
                                                'ó·"SCPD3-F-10-STR10-45"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR10-45"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 100
                                                'ÉXÉgÉçÅ[ÉNÇSÇUÅ`ÇPÇOÇO
                                                'ó·"SCPD3-F-10-STR46-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR46-100"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 101
                                                'ÉXÉgÉçÅ[ÉNÇPÇOÇPÅ`ÇQÇOÇO
                                                'ó·"SCPD3-F-10-STR101-200"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR101-200"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "16"
                                        Select Case True
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                                                'ÉXÉgÉçÅ[ÉNÇPÇOÅ`ÇSÇT
                                                'ó·"SCPD3-F-16-STR10-45"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR10-45"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 100
                                                'ÉXÉgÉçÅ[ÉNÇSÇUÅ`ÇPÇOÇO
                                                'ó·"SCPD3-F-16-STR46-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR46-100"
                                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 101
                                                'ÉXÉgÉçÅ[ÉNÇPÇOÇPÅ`ÇQÇUÇO
                                                'ó·"SCPD3-F-16-STR101-260"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-F-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-STR101-260"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                        Case Else
                            'Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 10, 1)
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "C"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "6"
                                            Select Case True
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 15
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-15"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-30"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-45"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-60"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 70
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-70"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 71 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 80
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-80"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 81 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 90
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-90"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 91
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-100"
                                            End Select
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case Else
                                            Select Case True
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 15
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-15"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-30"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-45"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-60"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 75
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-75"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 90
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-90"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-100"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 110
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-110"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 111 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 120
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-120"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 130
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-130"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 131 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 140
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-140"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 141 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 150
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-150"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 160
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-160"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 161 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 170
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-170"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 171 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 180
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-180"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 190
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-190"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 191 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 200
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-200"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 201 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 210
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-210"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 220
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-220"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 221 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 230
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-230"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 231 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 240
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-240"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 250
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-250"
                                                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 251
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "C-260"
                                            End Select
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select
                Case "SCPH3"    'SCPH3
                    Select Case True
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-105"
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select


            ''ÉoÉäÉGÅ[ÉVÉáÉì(î˜ë¨)â¡éZâøäiÉLÅ[
            'Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            '    Case "F"
            '        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            '            Case "6"
            '                Select Case True
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
            '                        'ÉXÉgÉçÅ[ÉN10Å`30
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "30"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And _
            '                         CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
            '                        'ÉXÉgÉçÅ[ÉN31Å`60
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR31" & CdCst.Sign.Hypen & "60"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61
            '                        'ÉXÉgÉçÅ[ÉN61Å`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR61" & CdCst.Sign.Hypen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '            Case "10"
            '                Select Case True
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
            '                        'ÉXÉgÉçÅ[ÉN10Å`45
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "45"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And _
            '                         CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
            '                        'ÉXÉgÉçÅ[ÉN46Å`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR46" & CdCst.Sign.Hypen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101
            '                        'ÉXÉgÉçÅ[ÉN101Å`200
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR101" & CdCst.Sign.Hypen & "200"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '            Case "16"
            '                Select Case True
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
            '                        'ÉXÉgÉçÅ[ÉN10Å`45
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "45"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And _
            '                         CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
            '                        'ÉXÉgÉçÅ[ÉN46Å`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR46" & CdCst.Sign.Hypen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101
            '                        'ÉXÉgÉçÅ[ÉN101Å`260
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
            '                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR101" & CdCst.Sign.Hypen & "260"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '        End Select
            'End Select

            'œ∏ﬁ»Øƒâ¡éZâøäiÉLÅ[
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Or _
               Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3-L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'éxéùå`éÆâ¡éZâøäiÉLÅ[
            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "CB" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "FA" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "LB" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "LS" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If


            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Or Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                Select Case True
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "D"
                        intIndex = 4
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "F"
                        intIndex = 4
                        'Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 10, 1) = "C"
                    Case objKtbnStrc.strcSelection.strKeyKataban.Trim = "C"
                        intIndex = 6
                    Case Else
                        intIndex = 5
                End Select
                If objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim <> "" Then
                    'ÉXÉCÉbÉ`â¡éZâøäiÉLÅ[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(intIndex + 2).Trim)

                    If objKtbnStrc.strcSelection.strOpSymbol(intIndex + 1).Trim <> "" Then
                        'ÉäÅ[Éhê¸í∑Ç≥â¡éZâøäiÉLÅ[
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(intIndex + 1).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(intIndex + 2).Trim)
                    End If
                End If
                'intIndex = intIndex + 2
            End If

            'ÉIÉvÉVÉáÉìÅEïtëÆïiâ¡éZâøäiÉLÅ[
            intIndex = 0
            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                Case "SCPD3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            'êHïiêªë¢çHíˆå¸ÇØè§ïi
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1)
                                Case "D"
                                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(8)) <> 0 Then
                                            intIndex = 8
                                        End If
                                    Else
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(5)) <> 0 Then
                                            intIndex = 5
                                        End If
                                    End If
                                Case "Z", "K", "M"
                                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(9)) <> 0 Then
                                            intIndex = 9
                                        End If
                                    Else
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(6)) <> 0 Then
                                            intIndex = 6
                                        End If
                                    End If
                                Case Else
                            End Select

                        Case Else
                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1)
                                Case "D"
                                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(7)) <> 0 Then
                                            intIndex = 7
                                        End If
                                    Else
                                        If Len(objKtbnStrc.strcSelection.strOpSymbol(4)) <> 0 Then
                                            intIndex = 4
                                        End If
                                    End If
                                Case "F", "L"
                                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "C" Then
                                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                                            If Len(objKtbnStrc.strcSelection.strOpSymbol(9)) <> 0 Then
                                                intIndex = 9
                                            End If
                                        Else
                                            If Len(objKtbnStrc.strcSelection.strOpSymbol(6)) <> 0 Then
                                                intIndex = 6
                                            End If
                                        End If
                                    Else
                                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                                            If Len(objKtbnStrc.strcSelection.strOpSymbol(7)) <> 0 Then
                                                intIndex = 7
                                            End If
                                        Else
                                            If Len(objKtbnStrc.strcSelection.strOpSymbol(4)) <> 0 Then
                                                intIndex = 4
                                            End If
                                        End If
                                    End If
                                Case Else
                                    'Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 10, 1)
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "C"
                                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                                                If Len(objKtbnStrc.strcSelection.strOpSymbol(9)) <> 0 Then
                                                    intIndex = 9
                                                End If
                                            Else
                                                If Len(objKtbnStrc.strcSelection.strOpSymbol(6)) <> 0 Then
                                                    intIndex = 6
                                                End If
                                            End If
                                        Case Else
                                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                                                If Len(objKtbnStrc.strcSelection.strOpSymbol(8)) <> 0 Then
                                                    intIndex = 8
                                                End If
                                            Else
                                                If Len(objKtbnStrc.strcSelection.strOpSymbol(5)) <> 0 Then
                                                    intIndex = 5
                                                End If
                                            End If
                                    End Select
                            End Select
                    End Select

                Case Else
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Or _
                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                        If Len(objKtbnStrc.strcSelection.strOpSymbol(8)) <> 0 Then
                            intIndex = 8
                        End If
                    Else
                        If Len(objKtbnStrc.strcSelection.strOpSymbol(5)) <> 0 Then
                            intIndex = 5
                        End If
                    End If
            End Select

            If intIndex <> 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intIndex), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) = "SCPD3-D" Then
                        decOpAmount(UBound(decOpAmount)) = 2
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Next
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
