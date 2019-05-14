'************************************************************************************
'*  ProgramID  ：KHPrice86
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空パッド
'*
'************************************************************************************
Module KHPrice86

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim objKataban As New KHKataban

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            ' 基本価格キー
            Select Case True
                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "C"
                    ' ロングストロークホルダ付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("VSP" & CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(6).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "M"
                    ' 小形ホルダタイプ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("VSP" & CdCst.Sign.Hypen & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "Q"
                    ' 吸着痕防止タイプ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("VSP" & CdCst.Sign.Hypen & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                     objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("VSP" & CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            ' フリーホルダ加算価格キー
            Select Case True
                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "M"
                    '小形ホルダタイプ
                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "C"
                    ' ロングストロークホルダ付
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        Case "F1", "F2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "VSP" & CdCst.Sign.Hypen & "P" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "F1", "F2"
                            '機種毎に価格キーを設定
                            Select Case True
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "R"
                                    'スタンダードタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-R/A-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "S"
                                    'スポンジタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-S-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "B"
                                    'ベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-B-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "E"
                                    '長円タイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-E-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "L"
                                    ' ソフトタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-L-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "1"
                                    ' ソフトベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "K"
                                    ' 滑り止めタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-K-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "F"
                                    ' フラットタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-F-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    'RM1610027 Start
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "A"
                                    ' ソフトベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    'RM1610027 End
                            End Select
                    End Select
            End Select

            ' フリーホルダ加算価格キー
            If Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 2) = "-V" Then
                ' 機種毎に価格キーを設定
                Select Case True
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "R"
                        'スタンダードタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-R/A-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "S"
                        'スポンジタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-S-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "B"
                        'ベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-B-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "W"
                        '多段ベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-W-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "E"
                        '長円タイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-E-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "L"
                        'ソフトタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-L-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "1"
                        'ソフトベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "K"
                        '滑り止めタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-K-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "C"
                        'ロングストロークホルダ付
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-P-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "M"
                        '小型真空パッド　スタンダードタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-M-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "Q"
                        ' 吸着痕防止タイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-Q-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "F"
                        'フラットタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-F-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1610027 Start
                    Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "VSP" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "A"
                        'ソフトベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1610027 End
                End Select
            End If

        Catch ex As Exception

            Throw ex

        Finally

            objKataban = Nothing

        End Try

    End Sub

End Module
