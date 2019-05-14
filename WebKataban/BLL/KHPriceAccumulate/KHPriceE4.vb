'************************************************************************************
'*  ProgramID  ：KHPriceE4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド(個別配線・省配線)　ＭＮ３Ｅ０／ＭＮ４Ｅ０
'*
'*　変更　　　　：ダミーブロック追加  RM0911XXX 2009/11/16 Y.Miura
'************************************************************************************
Module KHPriceE4
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
        'Public Const Rail As Integer
        Public Const Inspect1 As Integer = 25
        Public Const Inspect2 As Integer = 26
        Public Const Inspect3 As Integer = 27
        Public Const Inspect4 As Integer = 28
        'Public Const Tube As Integer 
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

        Dim intValve307DQty As Integer = 0           'バルブブロック 7mm（２個内蔵形）の数
        Dim intValve310DQty As Integer = 0           'バルブブロック10mm（２個内蔵形）の数
        Dim intValve307EQty As Integer = 0           'バルブブロック 7mm(Eオプション)の数
        Dim intValve310EQty As Integer = 0           'バルブブロック10mm(Eオプション)の数

        Dim intRegulatRAQty As Integer = 0          'レギュレータブロックRAの数
        Dim intRegulatRBQty As Integer = 0          'レギュレータブロックRBの数

        Dim strOptionA As String = String.Empty     'オプションA

        Dim bolOptionS As Boolean = False
        Dim bolOptionSA As Boolean = False
        Dim bolOptionC As Boolean = False

        ' 2008/12/03 追加
        Dim ItemKiriIchikbn As String = String.Empty        '切換位置区分
        Dim ItemSosakbn As String = String.Empty            '操作区分
        Dim ItemKokei As String = String.Empty              '接続口径
        Dim ItemChoatsu As String = String.Empty            '調圧
        Dim ItemSyudoSochi As String = String.Empty         '手動装置
        Dim ItemHaisen As String = String.Empty             '配線接続
        Dim ItemTanshi As String = String.Empty             '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim ItemOption As String = String.Empty             'オプション
        Dim ItemRensu As String = String.Empty              '連数
        Dim ItemDenatsu As String = String.Empty            '電圧
        Dim ItemCleanShiyo As String = String.Empty         'クリーン仕様
        Dim ItemHosyo As String = String.Empty              '保証

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN3E0", "MN4E0", "MN3E00", "MN4E00"
                    ItemKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim '切換位置区分
                    ItemSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim     '操作区分
                    ItemKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim       '接続口径
                    ItemChoatsu = objKtbnStrc.strcSelection.strOpSymbol(4).Trim     '調圧
                    ItemSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim  '手動装置
                    ItemHaisen = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      '配線接続
                    ItemTanshi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim      '端子･ｺﾈｸﾀﾋﾟﾝ配列
                    ItemOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim      'オプション
                    ItemRensu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim       '連数
                    ItemDenatsu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim    '電圧
                    ItemCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim 'クリーン仕様
                    ItemHosyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim      '保証
                Case "MN3EX0", "MN4EX0"
                    ItemKiriIchikbn = ""                                            '切換位置区分
                    ItemSosakbn = ""                                                '操作区分
                    ItemKokei = objKtbnStrc.strcSelection.strOpSymbol(1).Trim       '接続口径
                    ItemChoatsu = objKtbnStrc.strcSelection.strOpSymbol(2).Trim     '調圧
                    ItemSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(3).Trim  '手動装置
                    ItemHaisen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim      '配線接続
                    ItemTanshi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim      '端子･ｺﾈｸﾀﾋﾟﾝ配列
                    ItemOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      'オプション
                    ItemRensu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim       '連数
                    ItemDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim     '電圧
                    ItemCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'クリーン仕様
                    ItemHosyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim      '保証
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
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        Case "MN3E00", "MN4E00"
                                            strOpRefKataban(UBound(strOpRefKataban)) = "N4E00-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
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
                                                Case "N3E00", "N4E00"
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2)
                                                        Case "66", "67", "76", "77"
                                                            intValve3PDualQty = intValve3PDualQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            '↓RM1301005 2013/01/08 Y.Tachi
                                                            intValve307DQty = intValve307DQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case Else
                                                            intValve3PQty = intValve3PQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            '↓RM1301005 2013/01/08 Y.Tachi
                                                            intValve307EQty = intValve307EQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                                Case Else
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                                        Case "66", "67", "76", "77"
                                                            intValve3PDualQty = intValve3PDualQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            '↓RM1301005 2013/01/08 Y.Tachi
                                                            intValve310DQty = intValve310DQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case Else
                                                            intValve3PQty = intValve3PQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            '↓RM1301005 2013/01/08 Y.Tachi
                                                            intValve310EQty = intValve310EQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                            End Select
                                        Case "4"
                                            intValve4PQty = intValve4PQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                    'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(0, 5)
                                        Case "N3E00"
                                            intValve307Qty = intValve307Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "N4E00"
                                            intValve407Qty = intValve407Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Else
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(1, 1)
                                                Case "3"
                                                    intValve310Qty = intValve310Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case "4"
                                                    intValve410Qty = intValve410Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
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

                                    'Case 12 To 15
                                Case ItemNum.Exhaust1 To ItemNum.Exhaust4
                                    '給排気ブロック
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-S") >= 0 Or _
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-SA") >= 0 Or _
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-C") >= 0 Then
                                        Select Case True
                                            Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-SA") >= 0
                                                bolOptionSA = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") - 1)
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-S") >= 0
                                                bolOptionS = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") - 1)
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("-C") >= 0
                                                bolOptionC = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C") - 1)
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    'Case 16 To 17
                                Case ItemNum.Regulat1 To ItemNum.Regulat2
                                    'レギュレータブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 18 To 19
                                Case ItemNum.EndL To ItemNum.EndR
                                    'エンドブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 20 To 23
                                Case ItemNum.Plug1 To ItemNum.Plug4
                                    'ブランクプラグ＆サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    '    'RM0911XXX 2009/11/16 Y.Miura 追加
                                    'Case ItemNum.Rail
                                    '    '取付けレール
                                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    '    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 24 To 27
                                Case ItemNum.Inspect1 To ItemNum.Inspect4
                                    '検査成績書＆ケーブル＆継手＆コネクタ＆ソケット
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            '仕切りタイプ給排気ブロック("-S","-SA","-C")選択加算加算価格キー
                            If bolOptionS = True Or _
                               bolOptionSA = True Or _
                               bolOptionC = True Then
                                Select Case True
                                    Case bolOptionS = True
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-S"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case bolOptionSA = True
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-SA"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case bolOptionC = True
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-C"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End Select
                            End If

                            'クリーン仕様加算価格キー
                            If ItemCleanShiyo = "P70" Then
                                Select Case intLoopCnt
                                    'Case 1 To 2
                                    Case ItemNum.Elect1 To ItemNum.Elect2
                                        '電装ブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-DENSO-BLOCK-" & _
                                                                                   ItemCleanShiyo
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        'Case 4 To 11
                                    Case ItemNum.Valve1 To ItemNum.Valve7
                                        'バルブブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4, 2) = "00" Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                            Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & _
                                            Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Hypen) - 6) & _
                                            CdCst.Sign.Hypen & ItemCleanShiyo
                                        Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                            Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Hypen) - 1) & _
                                            CdCst.Sign.Hypen & ItemCleanShiyo
                                        End If

                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                                        'ダミーブロック
                                    Case ItemNum.Dummy1 To ItemNum.Dummy2
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   ItemCleanShiyo
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        'Case 12 To 15
                                    Case ItemNum.Exhaust1 To ItemNum.Exhaust4
                                        '給排気ブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) & "*-" & _
                                                                                   ItemCleanShiyo
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        'Case 18 To 19
                                    Case ItemNum.EndL To ItemNum.EndR
                                        'エンドブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   ItemCleanShiyo
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                End Select
                            End If

                            'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                            'オゾン仕様加算
                            If Not strOptionA.Equals("") Then
                                Select Case intLoopCnt
                                    Case ItemNum.Regulat1, ItemNum.Regulat2
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(0, 7) & CdCst.Sign.Hypen & _
                                                                                   strOptionA
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End Select
                            End If

                    End Select
                End If
            Next

            '取付レール長さ加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-BAA"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            'オプション加算価格キー
            strOpArray = Split(ItemOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                'RM0911XXX 2009/11/16 Y.Miura ダミーブロック追加
                'Select Case strOpArray(intLoopCnt).Trim
                '    Case ""
                '    Case Else
                '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '        strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                '        decOpAmount(UBound(decOpAmount)) = intValve3PQty

                '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '        strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & "-DUAL"
                '        decOpAmount(UBound(decOpAmount)) = intValve3PDualQty

                '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '        strOpRefKataban(UBound(strOpRefKataban)) = "N4E0" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                '        decOpAmount(UBound(decOpAmount)) = intValve4PQty
                'End Select
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "A"            'オゾン
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N3E0-A"
                        decOpAmount(UBound(decOpAmount)) = intValve310Qty

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-A"
                        decOpAmount(UBound(decOpAmount)) = intValve410Qty
                    Case Else
                        '↓RM1301005 2013/01/08 Y.Tachi
                        If strOpArray(intLoopCnt).Trim = "E" Then
                            If intValve310EQty <> 0 Then
                                '例 "N3E0-E"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve310EQty
                            End If
                            If intValve307EQty <> 0 Then
                                '例 "N3E00-E"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N3E00" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve307EQty
                            End If
                            If intValve310DQty <> 0 Then
                                '例 "N3E0-E-DUAL"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve310DQty
                            End If
                            If intValve307DQty <> 0 Then
                                '例 "N3E00-E-DUAL"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N3E00" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve307DQty
                            End If
                            If intValve410Qty <> 0 Then
                                '例 "N4E0-E"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N4E0" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve410Qty
                            End If
                            If intValve407Qty <> 0 Then
                                '例 "N3E00-E-DUAL"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N4E00" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = intValve407Qty
                            End If
                        Else
                            '例 "N3E0-E"
                            '例 "N3E0-A"
                            '例 "N3E0-F"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt)
                            decOpAmount(UBound(decOpAmount)) = intValve3PQty

                            '例 "N3E0-E-DUAL"
                            '例 "N3E0-A-DUAL"
                            '例 "N3E0-F-DUAL"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "N3E0" & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt) & "-DUAL"
                            decOpAmount(UBound(decOpAmount)) = intValve3PDualQty

                            '例 "N4E0-E"
                            '例 "N4E0-A"
                            '例 "N4E0-F"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "N4E0" & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt)
                            decOpAmount(UBound(decOpAmount)) = intValve4PQty
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
