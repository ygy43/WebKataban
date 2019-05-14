'************************************************************************************
'*  ProgramID  ：KHPriceM9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：アブソデックス　ＡＸ２０００Ｇ/ＡＸ２０００Ｔ
'*
'*  ・受付No：RM0907072  新型アブソデックス追加（AX1000T/AX2000T/AX4000T）
'*                                      更新日：2009/08/17   更新者：Y.Miura
'*  ・受付No：RM0908025  インターフェース仕様にCC-Linkを追加（AX1000T/AX2000T/AX4000T）
'*                                      更新日：2009/09/02   更新者：Y.Miura
'************************************************************************************
Module KHPriceM9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '取付ベース加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'ケーブル変更加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                'RM0907072 2009/08/17 Y.Miura
                'If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" Then
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                '    decOpAmount(UBound(decOpAmount)) = 1
                'Else
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                '    decOpAmount(UBound(decOpAmount)) = 1
                'End If                
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                Dim strOpSign As String
                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Case "TS", "TH"
                        strOpSign = "***T"
                    Case Else
                        strOpSign = "***"
                End Select
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & strOpSign & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & strOpSign & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                End If
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'ドライバ電源電圧加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'ノックピン加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '本体表面処理加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'RM0908025 2009/09/02 Y.Miura
            'インターフェース仕様加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol.Length > 13 Then       'RM0912039 オプション追加による表示不具合修正
                If objKtbnStrc.strcSelection.strOpSymbol(13).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                               "***" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
