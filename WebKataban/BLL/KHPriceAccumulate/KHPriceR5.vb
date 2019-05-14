'************************************************************************************
'*  ProgramID  ：KHPriceR5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2011/10/21   作成者：Y.Tachi
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線ブロックマニホールド             ＭＮ３Ｑ０／ＭＴ３Ｑ０
'*
'************************************************************************************
Module KHPriceR5
    'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
    Private Structure ItemNum
        Private dummy As Integer
        Public Const Elect1 As Integer = 1
        Public Const Elect2 As Integer = 2
        Public Const Wiring As Integer = 3
        Public Const Valve1 As Integer = 4
        Public Const Valve2 As Integer = 5
        Public Const Valve3 As Integer = 6
        Public Const Valve4 As Integer = 7
        Public Const Valve5 As Integer = 8
        Public Const Valve6 As Integer = 9
        Public Const Valve7 As Integer = 10
        Public Const Dummy1 As Integer = 11
        Public Const Dummy2 As Integer = 12
        Public Const Exhaust1 As Integer = 13
        Public Const Exhaust2 As Integer = 14
        Public Const Exhaust3 As Integer = 15
        Public Const Exhaust4 As Integer = 16
        Public Const Regulat1 As Integer = 17
        Public Const Regulat2 As Integer = 18
        Public Const EndL As Integer = 19
        Public Const EndR As Integer = 20
        Public Const Plug1 As Integer = 21
        Public Const Plug2 As Integer = 22
        Public Const Plug3 As Integer = 23
        Public Const Plug4 As Integer = 24
        Public Const Inspect1 As Integer = 25
        Public Const Inspect2 As Integer = 26
        Public Const Inspect3 As Integer = 27
        Public Const Inspect4 As Integer = 28

    End Structure

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intStationQty As Integer = 0
        Dim intValveQty As Integer = 0

        Dim intValve3PQty As Integer = 0
        Dim intValve3PDualQty As Integer = 0
        Dim intValve4PQty As Integer = 0

        'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
        Dim intValve307Qty As Integer = 0           'バルブブロック 7mmの数
        Dim intValve310Qty As Integer = 0           'バルブブロック10mmの数
        Dim intValve407Qty As Integer = 0           'バルブブロック 7mmの数
        Dim intValve410Qty As Integer = 0           'バルブブロック10mmの数

        Dim strOptionA As String = String.Empty     'オプションA

        Dim bolOptionS As Boolean = False
        Dim bolOptionSA As Boolean = False
        Dim bolOptionC As Boolean = False

        ' 2008/12/03 追加
        Dim ItemKiriIchikbn As String = String.Empty        '切換位置区分
        Dim ItemSosakbn As String = String.Empty            '操作区分
        Dim ItemKokei As String = String.Empty              '接続口径
        Dim ItemSyudoSochi As String = String.Empty         '手動装置
        Dim ItemHaisen As String = String.Empty             '配線接続
        Dim ItemOption As String = String.Empty             'オプション
        Dim ItemRensu As String = String.Empty              '連数
        Dim ItemDenatsu As String = String.Empty            '電圧

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN3Q0", "MT3Q0"
                    ItemKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim '切換位置区分
                    ItemSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim     '操作区分
                    ItemKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim       '接続口径
                    ItemSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(4).Trim  '手動装置
                    ItemHaisen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim      '配線接続
                    ItemOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      'オプション
                    ItemRensu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim       '連数
                    ItemDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim     '電圧
            End Select

            '連数設定
            intStationQty = CInt(ItemRensu)

            'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
            'オプション加算価格キー
            strOpArray = Split(ItemOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "A"
                        strOptionA = strOpArray(intLoopCnt).Trim
                    Case "E"
                    Case Else
                End Select
            Next

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                '形番・使用数が存在する場合
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                    bolOptionS = False
                    bolOptionSA = False
                    bolOptionC = False

                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, _
                             CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, _
                             CdCst.Manifold.InspReportEn.English
                        Case Else
                            Select Case intLoopCnt
                                'Case 1 To 2
                                Case ItemNum.Elect1 To ItemNum.Elect2
                                    '電装ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 3
                                Case ItemNum.Wiring
                                    '個別配線

                                    'Case 4 To 11
                                Case ItemNum.Valve1 To ItemNum.Valve7
                                    'バルブブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'バルブ数をカウント
                                    intValveQty = intValveQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'バルブ数をカウント
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 1)
                                        Case "3"
                                            Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5)
                                                Case "MN3Q0", "MT3Q0"
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                                        Case "66"
                                                            intValve3PDualQty = intValve3PDualQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                            End Select
                                    End Select

                                    'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(0, 5)
                                        Case "MN3Q0", "MT3Q0"
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(1, 1)
                                                Case "3"
                                                    intValve310Qty = intValve310Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select
                                    End Select
                                    'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                                Case ItemNum.Dummy1 To ItemNum.Dummy2
                                    'ダミーブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 13 To 16
                                Case ItemNum.Exhaust1 To ItemNum.Exhaust4
                                    '給排気ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 17 To 18
                                Case ItemNum.Regulat1 To ItemNum.Regulat2
                                    'レギュレータブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 19 To 20
                                Case ItemNum.EndL To ItemNum.EndR
                                    'エンドブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 21 To 24
                                Case ItemNum.Plug1 To ItemNum.Plug4
                                    'ブランクプラグ＆サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 26 To 28
                                Case ItemNum.Inspect2 To ItemNum.Inspect4
                                    '検査成績書＆ケーブル＆継手＆コネクタ＆ソケット
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select
                    End Select
                End If
            Next

            'MN3Q0シリーズのみ
            If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "N") <> 0 Then
                '取付レール長さ加算価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-BAA"
                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
            End If

            'ダイレクトマウント方式加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MT3Q0"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "MT3Q0-DM"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'ノンロック式手動装置加算価格キー
            If ItemSyudoSochi = "M" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0" & CdCst.Sign.Hypen & ItemSyudoSochi
                decOpAmount(UBound(decOpAmount)) = intValveQty
            End If

            'オプション加算価格キー
            strOpArray = Split(ItemOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "F"
                        '例 "N3Q0-F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0" & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt)
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                    Case "P", "N"
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
