'************************************************************************************
'*  ProgramID  ：KHPrice71
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/13   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド(個別配線・省配線)　ＭＮ３Ｇ／ＭＮ４Ｇ
'*
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　MN3GA1/MN3GA2/MN3GAX12/MN4GA1/MN4GA2/MN4GAX12/
'*                           MN3GB1/MN3GB2/MN3GBX12/MN4GB1/MN4GB2/MN4GBX12にオプションボックス追加
'*  ・受付No：RM0904031  4GD2/4GE2機種追加
'*                                      更新日：2009/06/23   更新者：Y.Miura
'*  ・RM1005030 2010/05/12 Y.Miura マニホールド追加
'*                      MN3GD1/MN3GDE1/MN3GD2/MN3GE2/MN4GD1/MN4GE1/MN4GD2/MN4GE2/MN4GD3/MN4GE3
'************************************************************************************
Module KHPrice71

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intStationQty As Integer
        Dim intQuantity As Integer

        Dim intValveQty As Integer
        Dim intValveQty2SWD As Integer
        Dim intValveQty3SWD As Integer
        Dim intValveQty4SWD As Integer
        Dim intValveQty5SWD As Integer
        Dim intValveQty3P As Integer
        Dim intValveQty4P As Integer
        Dim intValveQtyDual1 As Integer
        Dim intValveQtyDual2 As Integer

        Dim intN3GA1_EV_OPT As Integer
        Dim intN3GA2_EV_OPT As Integer
        Dim intN3GB1_EV_OPT As Integer      '2013/03/05
        Dim intN3GB2_EV_OPT As Integer      '2013/03/05
        Dim intN4GA1_EV_OPT As Integer
        Dim intN4GA2_EV_OPT As Integer
        Dim intN4GB1_EV_OPT As Integer
        Dim intN4GB2_EV_OPT As Integer

        Dim intN4GA1_MP_OPT As Integer
        Dim intN4GA2_MP_OPT As Integer
        Dim intN4GB1_MP_OPT As Integer
        Dim intN4GB2_MP_OPT As Integer

        Dim intN3GA1_EV_OTH As Integer
        Dim intN3GA2_EV_OTH As Integer
        Dim intN3GB1_EV_OTH As Integer      '2013/03/05
        Dim intN3GB2_EV_OTH As Integer      '2013/03/05
        Dim intN4GA1_EV_OTH As Integer
        Dim intN4GA2_EV_OTH As Integer
        Dim intN4GB1_EV_OTH As Integer
        Dim intN4GB2_EV_OTH As Integer

        Dim intN3GA1_EV_3SWD As Integer
        Dim intN3GA1_EV_5SWD As Integer
        Dim intN3GA2_EV_3SWD As Integer
        Dim intN3GA2_EV_5SWD As Integer
        Dim intN3GB1_EV_3SWD As Integer      '2013/03/05
        Dim intN3GB1_EV_5SWD As Integer      '2013/03/05
        Dim intN3GB2_EV_3SWD As Integer      '2013/03/05
        Dim intN3GB2_EV_5SWD As Integer      '2013/03/05
        Dim intN4GA1_EV_3SWD As Integer
        Dim intN4GA1_EV_5SWD As Integer
        Dim intN4GA2_EV_3SWD As Integer
        Dim intN4GA2_EV_5SWD As Integer
        Dim intN4GB1_EV_3SWD As Integer
        Dim intN4GB1_EV_5SWD As Integer
        Dim intN4GB2_EV_3SWD As Integer
        Dim intN4GB2_EV_5SWD As Integer

        Dim intN3GA1_EV_3P As Integer
        Dim intN3GA2_EV_3P As Integer
        Dim intN3GB1_EV_3P As Integer      '2013/03/05
        Dim intN3GB2_EV_3P As Integer      '2013/03/05
        Dim intN4GA1_EV_4P As Integer
        Dim intN4GA2_EV_4P As Integer
        Dim intN4GB1_EV_4P As Integer
        Dim intN4GB2_EV_4P As Integer

        Dim strPortSize As String

        Dim strABDE As String       'RM1005030
        Dim str1234 As String       'RM1005030 追加

        Dim strKiriIchikbn As String = ""       '切換位置区分
        Dim strSosakbn As String = ""           '操作区分
        Dim strKokei As String = ""             '接続口径
        Dim strDensen As String = ""            '電線接続
        Dim strTanshi As String = ""            '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strSyudoSochi As String = ""        '手動装置
        Dim strOption As String = ""            'オプション
        Dim strRensu As String = ""             '連数
        Dim strDenatsu As String = ""           '電圧
        Dim strCleanShiyo As String = ""        'クリーン仕様
        Dim strHosyo As String = ""             '保証
        Dim strLion As String = ""              '二次電池
        Dim strOptionFP1 As String = ""         '食品製造工程向け商品 RM1610013

        Try

            '機種の文字列を取り出す       'RM1005030 追加
            strABDE = objKtbnStrc.strcSelection.strSeriesKataban.Trim.PadRight(6, " ").Substring(4, 1)
            str1234 = objKtbnStrc.strcSelection.strSeriesKataban.Trim.PadRight(6, " ").Substring(5, 1)

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                'Case "R", "U"
                Case "R", "U", "S", "V" 'RM1610013
                    strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim          '切換位置区分
                    strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim              '操作区分
                    strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim                '接続口径
                    strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim               '電線接続
                    strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                    strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim           '手動装置
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim               'オプション
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim                '連数
                    strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim              '電圧
                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim           'クリーン仕様
                    strHosyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim                '保証
                    If Not objKtbnStrc.strcSelection.strSeriesKataban.Contains("X12") Then
                        strLion = objKtbnStrc.strcSelection.strOpSymbol(13).Trim                 '二次電池
                    Else                                                                     'RM1610013 Start
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 13 Then              'RM1610013 Start 16/10/31
                            strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(13).Trim        '食品製造工程向け 
                        End If
                    End If
                    If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 14 Then              'RM1610013 Start
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(14).Trim        '食品製造工程向け 
                    End If                                                                   'RM1610013 End

                Case Else
                    strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim          '切換位置区分
                    strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim              '操作区分
                    strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim                '接続口径
                    strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '電線接続
                    strTanshi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                    strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim           '手動装置
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim               'オプション
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim                '連数
                    strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim              '電圧
                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim           'クリーン仕様
                    strHosyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim                '保証
                    If Not objKtbnStrc.strcSelection.strSeriesKataban.Contains("X12") Then
                        strLion = objKtbnStrc.strcSelection.strOpSymbol(12).Trim                 '二次電池
                    End If
            End Select


            'バルブブロック連数
            intStationQty = CInt(strRensu.Trim)

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                            '加算なし
                        Case Else
                            Select Case intLoopCnt
                                Case 1
                                    '電装ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 2 To 9
                                    '電磁弁付バルブブロック＆MP付バルブブロック
                                    '電磁弁付バルブブロックの時
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If (Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 4) = "N4GE" Or _
                                           Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 4) = "N3GE") And _
                                           Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 7) = "1" Then
                                            If (Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 4) = "N4GE" And _
                                               Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 6) = "CL") Or _
                                               (Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 4) = "N3GE" And _
                                               Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 10, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 6) = "CL") Then

                                                'G1・G2が入っている場合は別の価格キーを付与するよう処理を追加 RM1702017 追加
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G1") >= 1 Then
                                                    'G1が入っている場合はG1を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-CL-G1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-CL-G1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G2") >= 1 Then
                                                    'G2が入っていたら末尾に「-G2」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-CL-G2"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-CL-G2"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X1") >= 1 Then
                                                    'X1が入っていたら末尾に「-X1」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-CL-X1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-CL-X1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X") >= 1 Then
                                                    'Xが入っていたら末尾に「-X」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-CL-X"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-CL-X"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                Else
                                                    '従来通り
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-CL"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-CL"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                End If
                                            Else
                                                'G1・G2が入っている場合は別の価格キーを付与するよう処理を追加 RM1702017 追加
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G1") >= 1 Then
                                                    'G1が入っている場合はG1を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-C-G1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-C-G1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G2") >= 1 Then
                                                    'G2が入っていたら末尾に「-G2」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-C-G2"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-C-G2"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X1") >= 1 Then
                                                    'Xが入っていたら末尾に「-X」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-C-X1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-C-X1"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X") >= 1 Then
                                                    'Xが入っていたら末尾に「-X」を付与
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-C-X"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-C-X"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                Else
                                                    'G2も入っていなかったら従来通り
                                                    If Left(strDensen.Trim, 1) = "T" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-C"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-C"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End If
                                                End If
                                            End If
                                        Else
                                            'G1・G2が入っている場合は別の価格キーを付与するよう処理を追加 RM1702017 追加
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G1") >= 1 Then
                                                'G1が入っている場合はG1を付与
                                                If Left(strDensen.Trim, 1) = "T" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-G1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G2") >= 1 Then
                                                'G2が入っていたら末尾に「-G2」を付与
                                                If Left(strDensen.Trim, 1) = "T" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-G2"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G2"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X1") >= 1 Then
                                                'X1が入っていたら末尾に「-G2」を付与
                                                If Left(strDensen.Trim, 1) = "T" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-X1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-X1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X") >= 1 Then
                                                'Xが入っていたら末尾に「-G2」を付与
                                                If Left(strDensen.Trim, 1) = "T" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N-X"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-X"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            Else
                                                'G2も入っていなかったら従来通り
                                                If Left(strDensen.Trim, 1) = "T" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-A2N"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1)
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        End If

                                        '切換位置区分が"1","11"の時,数量はバルブブロック部の使用数を集計する
                                        If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                            intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            '切換位置区分が"2","3","4","5"の時、数量はバルブブロック部の使用数の2倍を集計する
                                            intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                        End If

                                        '電磁弁数(バルブ数)をカウントする①
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                            Case "2"
                                                intValveQty2SWD = intValveQty2SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "3"
                                                intValveQty3SWD = intValveQty3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "4"
                                                intValveQty4SWD = intValveQty4SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "5"
                                                intValveQty5SWD = intValveQty5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select

                                        '電磁弁数(バルブ数)をカウントする②
                                        intValveQty = intValveQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 1)
                                            Case "3"
                                                intValveQty3P = intValveQty3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "4"
                                                intValveQty4P = intValveQty4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select

                                        '電磁弁数(デュアル3ポート弁のバルブ数)をカウントする③
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2)
                                            Case "66", "67", "76", "77"
                                                Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1)
                                                    Case "1"
                                                        intValveQtyDual1 = intValveQtyDual1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Case "2"
                                                        intValveQtyDual2 = intValveQtyDual2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End Select
                                        End Select

                                        'ミックスマニホールドの時、1タイプ、2タイプそれぞれの使用数をカウントする
                                        If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") <> 0 Then
                                            Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5)
                                                'Case "N3GA1"       'RM10060XX
                                                Case "N3GA1", "N3GD1"
                                                    intN3GA1_EV_OPT = intN3GA1_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    '切換位置区分が"1","11"の時,数量はバルブブロック部の使用数を集計する
                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN3GA1_EV_OTH = intN3GA1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        '切換位置区分が"2","3","4","5"の時、数量はﾊﾞﾙﾌﾞﾌﾞﾛｯｸ部の使用数の2倍を集計する
                                                        intN3GA1_EV_OTH = intN3GA1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN3GA1_EV_3SWD = intN3GA1_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN3GA1_EV_5SWD = intN3GA1_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN3GA1_EV_3P = intN3GA1_EV_3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    'Case "N3GA2"       'RM10060XX
                                                Case "N3GA2", "N3GD2"
                                                    intN3GA2_EV_OPT = intN3GA2_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN3GA2_EV_OTH = intN3GA2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN3GA2_EV_OTH = intN3GA2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN3GA2_EV_3SWD = intN3GA2_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN3GA2_EV_5SWD = intN3GA2_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN3GA2_EV_3P = intN3GA2_EV_3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    '↓RM1303003 2013/03/05 Y.Tachi
                                                Case "N3GB1", "N3GE1"
                                                    intN3GB1_EV_OPT = intN3GB1_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    '切換位置区分が"1","11"の時,数量はバルブブロック部の使用数を集計する
                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN3GB1_EV_OTH = intN3GB1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        '切換位置区分が"2","3","4","5"の時、数量はﾊﾞﾙﾌﾞﾌﾞﾛｯｸ部の使用数の2倍を集計する
                                                        intN3GB1_EV_OTH = intN3GB1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN3GB1_EV_3SWD = intN3GB1_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN3GB1_EV_5SWD = intN3GB1_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN3GB1_EV_3P = intN3GB1_EV_3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case "N3GB2", "N3GE2"
                                                    intN3GB2_EV_OPT = intN3GB2_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN3GB2_EV_OTH = intN3GB2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN3GB2_EV_OTH = intN3GB2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN3GB2_EV_3SWD = intN3GB2_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN3GB2_EV_5SWD = intN3GB2_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN3GB2_EV_3P = intN3GB2_EV_3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    '↑RM1303003 2013/03/05 Y.Tachi
                                                    'Case "N4GA1"       'RM10060XX
                                                Case "N4GA1", "N4GD1"
                                                    intN4GA1_EV_OPT = intN4GA1_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN4GA1_EV_OTH = intN4GA1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN4GA1_EV_OTH = intN4GA1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN4GA1_EV_3SWD = intN4GA1_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN4GA1_EV_5SWD = intN4GA1_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN4GA1_EV_4P = intN4GA1_EV_4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    'Case "N4GA2"       'RM10060XX
                                                Case "N4GA2", "N4GD2"
                                                    intN4GA2_EV_OPT = intN4GA2_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN4GA2_EV_OTH = intN4GA2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN4GA2_EV_OTH = intN4GA2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN4GA2_EV_3SWD = intN4GA2_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN4GA2_EV_5SWD = intN4GA2_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN4GA2_EV_4P = intN4GA2_EV_4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    'Case "N4GB1"       'RM10060XX
                                                Case "N4GB1", "N4GE1"
                                                    intN4GB1_EV_OPT = intN4GB1_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN4GB1_EV_OTH = intN4GB1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN4GB1_EV_OTH = intN4GB1_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN4GB1_EV_3SWD = intN4GB1_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN4GB1_EV_5SWD = intN4GB1_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN4GB1_EV_4P = intN4GB1_EV_4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    'Case "N4GB2"       'RM10060XX
                                                Case "N4GB2", "N4GE2"
                                                    intN4GB2_EV_OPT = intN4GB2_EV_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                                        intN4GB2_EV_OTH = intN4GB2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        intN4GB2_EV_OTH = intN4GB2_EV_OTH + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                                    End If

                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1)
                                                        Case "3"
                                                            intN4GB2_EV_3SWD = intN4GB2_EV_3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case "5"
                                                            intN4GB2_EV_5SWD = intN4GB2_EV_5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select

                                                    intN4GB2_EV_4P = intN4GB2_EV_4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select
                                        End If
                                    Else
                                        'MP付バルブブロックの時
                                        Select Case Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim)
                                            Case 8
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case 9
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case Else
                                                '接続口径除去
                                                Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9, 1)
                                                    Case "-"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8)
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Case Else
                                                        'If objKtbnStrc.strcSelection.strKeyKataban = "R" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Then
                                                        If objKtbnStrc.strcSelection.strKeyKataban = "R" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                                                            objKtbnStrc.strcSelection.strKeyKataban = "S" Or objKtbnStrc.strcSelection.strKeyKataban = "V" Then 'RM1610013
                                                            Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 10, 1)
                                                                Case "-"
                                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9)
                                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                                Case Else
                                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 10)
                                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            End Select
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9)
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End If
                                                End Select
                                        End Select

                                        'ミックスマニホールドの時、1タイプ、2タイプそれぞれの仕様数をカウントする
                                        If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") <> 0 Then
                                            Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5)
                                                Case "N4GA1", "N4GD1"   'RM10060XX
                                                    intN4GA1_MP_OPT = intN4GA1_MP_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case "N4GA2", "N4GD2"
                                                    intN4GA2_MP_OPT = intN4GA2_MP_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case "N4GB1", "N4GE1"
                                                    intN4GB1_MP_OPT = intN4GB1_MP_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case "N4GB2", "N4GE2"
                                                    intN4GB2_MP_OPT = intN4GB2_MP_OPT + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select
                                        End If
                                    End If

                                    '  '     'FPシリーズ加算 RM1610034
                                    '       If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-C4") Or
                                    '          objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-C6") Or
                                    '          objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-C8") Then
                                    '           Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                    '               Case "S", "V"
                                    '                   ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '                   ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '                   ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '                   strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                    '                   decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    '           End Select
                                    '       End If

                                Case 10
                                    'ミックスブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FPシリーズ　加算 RM1610034
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-MIX") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If


                                Case 11 To 14
                                    '個別給排気スペーサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FPシリーズ給気スペーサ、排気スペーサ加算 RM1610034
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-GWS4") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-GWS6") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-GWS8") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-GWS10") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                    'FPシリーズインストップ弁スペーサ加算
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-IS") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                Case 15 To 17
                                    '給排気ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FPシリーズ加算 RM1610034
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-6") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-6X") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                Case 18 To 19
                                    '仕切りブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FPシリーズ加算 RM1610034
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-SA") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-S") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-SP") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-SE") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                Case 20 To 21
                                    'エンドブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FPシリーズ加算 RM1610034
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-ER") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-EXR") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "S", "V"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                Case 22 To 24
                                    'ブランクプラグ＆サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 22
                                    '    'DINレール
                                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    '    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 25 To 26
                                    'Case 23 To 24
                                    'ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 27
                                    'Case 25
                                    'タグ銘板
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-TAG"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            'Select Case Left(strOpRefKataban(UBound(strOpRefKataban)), 8)
                            '    Case "MN4G-BAA"
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                            'End Select

                            'クリーンルーム仕様加算価格キー
                            If strCleanShiyo.Trim = "P70" Then
                                Select Case intLoopCnt
                                    Case 1
                                        '電装ブロック(クリーンルーム仕様加算)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-DENSO-BLOCK-" & _
                                                                                   strCleanShiyo.Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 2 To 9
                                        '電磁弁付バルブブロック＆MP付バルブブロック(クリーンルーム仕様加算)
                                        '電磁弁付バルブブロックの時
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-" & strCleanShiyo.Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            'MP付バルブブロックの時
                                            Select Case Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim)
                                                Case 8
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                               strCleanShiyo.Trim
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case 9
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                               strCleanShiyo.Trim
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Case Else
                                                    '接続口径除去
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9, 1)
                                                        Case "-"
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8) & CdCst.Sign.Hypen & _
                                                                                                       strCleanShiyo.Trim
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9) & CdCst.Sign.Hypen & _
                                                                                                       strCleanShiyo.Trim
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                            End Select
                                        End If
                                    Case 10
                                        'ミックスブロック(クリーンルーム仕様加算)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   strCleanShiyo.Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 15 To 17
                                        '給排気ブロック(クリーンルーム仕様加算)
                                        'If objKtbnStrc.strcSelection.strKeyKataban = "R" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban = "R" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban = "S" Or objKtbnStrc.strcSelection.strKeyKataban = "V" Then 'RM1610013
                                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, 1) = "K" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8) & "-*-" & _
                                                                                           strCleanShiyo.Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) & "-*-" & _
                                                                                           strCleanShiyo.Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7, 1) = "K" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) & "-*-" & _
                                                                                           strCleanShiyo.Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) & "-*-" & _
                                                                                           strCleanShiyo.Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    Case 18 To 19
                                        '仕切りブロック(クリーンルーム仕様加算)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   strCleanShiyo.Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 20 To 21
                                        'エンドブロック(クリーンルーム仕様加算)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   strCleanShiyo.Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End Select
                            End If
                    End Select
                End If
            Next

            'DINレール
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "MN4G-BAA"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            '電線接続／省配線接続　加算価格キー
            If strDensen.Trim <> "" Then
                If Left(strDensen.Trim, 1) <> "T" Then
                    'ミックスマニホールド以外の時(通常)
                    If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") = 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strDensen.Trim
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    Else
                        'ミックスマニホールドの時
                        If intN3GA1_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim       'RM10060XX
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "1-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_OTH
                        End If
                        If intN3GA2_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "2-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_OTH
                        End If
                        If intN4GA1_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA1-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN4GA1_EV_OTH
                        End If
                        If intN4GA2_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA2-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN4GA2_EV_OTH
                        End If
                        If intN4GB1_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB1-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN4GB1_EV_OTH
                        End If
                        If intN4GB2_EV_OTH <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB2-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strDensen.Trim
                            decOpAmount(UBound(decOpAmount)) = intN4GB2_EV_OTH
                        End If
                    End If
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "S", "E", "Q"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        decOpAmount(UBound(decOpAmount)) = intValveQty + intValveQty2SWD + intValveQtyDual1 + intValveQtyDual2 + intValveQty3SWD + intValveQty5SWD + intValveQty4SWD

                    Case "K", "A", "L"
                        If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") = 0 Then
                            'ミックスマニホールド以外の時(通常)
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intValveQty
                        Else
                            'ミックスマニホールドの時
                            If intN3GA1_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim         'RM10060XX
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_OPT
                            End If
                            If intN3GA2_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_OPT
                            End If
                            If intN4GA1_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA1-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN4GA1_EV_OPT
                            End If
                            If intN4GA2_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA2-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN4GA2_EV_OPT
                            End If
                            If intN4GB1_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB1-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN4GB1_EV_OPT
                            End If
                            If intN4GB2_EV_OPT <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB2-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intN4GB2_EV_OPT
                            End If
                        End If
                    Case "H"
                        'If Not objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" And Not objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                        If Not objKtbnStrc.strcSelection.strKeyKataban = "R" And Not objKtbnStrc.strcSelection.strKeyKataban = "U" And _
                           Not objKtbnStrc.strcSelection.strKeyKataban = "S" And Not objKtbnStrc.strcSelection.strKeyKataban = "V" Then 'RM1610013
                            If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") = 0 Then
                                'ミックスマニホールド以外の時（通常）
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQty3SWD - intValveQty5SWD
                            Else
                                'ミックスマニホールドの時
                                If intN3GA1_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim     'RM10060XX
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_OPT - intN3GA1_EV_3SWD - intN3GA1_EV_5SWD
                                End If
                                If intN3GA2_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_OPT - intN3GA2_EV_3SWD - intN3GA2_EV_5SWD
                                End If

                                '↓RM1303003 2013/03/05 Y.Tachi
                                If intN3GB1_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN3GB1_EV_OPT - intN3GB1_EV_3SWD - intN3GB1_EV_5SWD
                                End If
                                If intN3GB2_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN3GB2_EV_OPT - intN3GB2_EV_3SWD - intN3GB2_EV_5SWD
                                End If
                                '↑RM1303003 2013/03/05 Y.Tachi

                                If intN4GA1_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA1-" & strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN4GA1_EV_OPT - intN4GA1_EV_3SWD - intN4GA1_EV_5SWD
                                End If
                                If intN4GA2_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA2-" & strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN4GA2_EV_OPT - intN4GA2_EV_3SWD - intN4GA2_EV_5SWD
                                End If
                                If intN4GB1_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB1-" & strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "1-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN4GB1_EV_OPT - intN4GB1_EV_3SWD - intN4GB1_EV_5SWD
                                End If
                                If intN4GB2_EV_OPT <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB2-" & strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & "2-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intN4GB2_EV_OPT - intN4GB2_EV_3SWD - intN4GB2_EV_5SWD
                                End If
                            End If
                        End If
                    Case "F"
                        If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "X12") = 0 Then
                            'ミックスマニホールド以外の時(通常)
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "MN3GA1"
                                Case "MN3GA1", "MN3GD1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQtyDual1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    'RM0904031 2009/06/23 Y.Miura
                                    'Case "MN3GA2"
                                Case "MN3GA2", "MN3GD2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQtyDual2

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                Case "MN4GA1"
                                    If intValveQty3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual1

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    End If
                                    If intValveQty4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                    End If
                                Case "MN4GA2"
                                    If intValveQty3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual2

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                    End If
                                    If intValveQty4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                    End If

                                    'RM10060XX
                                Case "MN4GD1"
                                    If intValveQty3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual1

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    End If
                                    If intValveQty4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                    End If

                                    'RM0904031 2009/06/23 Y.Miura
                                Case "MN4GD2"
                                    If intValveQty3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual2

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                    End If
                                    If intValveQty4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                    End If

                                Case "MN3GB1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intStationQty

                                    'RM0904031 2009/06/23 Y.Miura
                                Case "MN3GE1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intStationQty

                                Case "MN3GB2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intStationQty

                                    'RM0904031 2009/06/23 Y.Miura
                                Case "MN3GE2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intStationQty

                                Case "MN4GB1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB1-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty - intValveQtyDual1

                                    'RM0904031 2009/06/23 Y.Miura
                                Case "MN4GE1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty - intValveQtyDual1

                                Case "MN4GB2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB2-" & strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty - intValveQtyDual2

                                    'RM0904031 2009/06/23 Y.Miura
                                Case "MN4GE2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty - intValveQtyDual2

                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "MN3GAX12"
                                    If intN3GA1_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_OPT - intValveQtyDual1

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    End If
                                    If intN3GA2_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_OPT - intValveQtyDual2

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                    End If
                                Case "MN4GAX12"
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_3P - intValveQtyDual1
                                    End If
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    End If
                                    If intN3GA2_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_3P - intValveQtyDual2
                                    End If
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GA2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                    End If
                                    If intN4GA1_EV_4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GA1_EV_4P
                                    End If
                                    If intN4GA2_EV_4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GA2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GA2_EV_4P
                                    End If

                                    '↓RM1303003 2013/03/05 Y.Tachi
                                Case "MN4GDX12"
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GD1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA1_EV_3P - intValveQtyDual1
                                    End If
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GD1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual1
                                    End If
                                    If intN3GA2_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GD2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN3GA2_EV_3P - intValveQtyDual2
                                    End If
                                    If intN3GA1_EV_3P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN3GD2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                    End If
                                    If intN4GA1_EV_4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GD1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GA1_EV_4P
                                    End If
                                    If intN4GA2_EV_4P <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GD2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GA2_EV_4P
                                    End If
                                    '↑RM1303003 2013/03/05 Y.Tachi

                                Case "MN3GBX12"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2
                                Case "MN4GBX12"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1

                                    If intN4GB1_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GB1_EV_OPT
                                    End If

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GB2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2

                                    If intN4GB2_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GB2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GB2_EV_OPT
                                    End If

                                    '↓RM1303003 2013/03/05 Y.Tachi
                                Case "MN4GEX12"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GE1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual1

                                    If intN4GB1_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GE1-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GB1_EV_OPT
                                    End If

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "MN3GE2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                    decOpAmount(UBound(decOpAmount)) = intValveQtyDual2

                                    If intN4GB2_EV_OPT <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "MN4GE2-" & strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = intN4GB2_EV_OPT
                                    End If
                                    '↑RM1303003 2013/03/05 Y.Tachi
                            End Select
                        End If
                    Case "Z4", "Z5"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "4G-" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
            '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                If Not strOption.Contains("H") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                    decOpAmount(UBound(decOpAmount)) = intValveQty3SWD + intValveQty5SWD
                End If
            End If

            '端子・コネクタピン加算価格キー
            If strTanshi = "W1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strTanshi
                decOpAmount(UBound(decOpAmount)) = intValveQty - (intValveQty2SWD + intValveQty3SWD + intValveQty4SWD + intValveQty5SWD)
            End If

            If strHosyo.Trim = "UL" Then
                '接続口径(継手エルボ)加算価格キー
                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                        Select Case intLoopCnt
                            Case 2 To 9
                                '電磁弁付バルブブロック＆MP付バルブブロック
                                '"-C3N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C3N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-C4N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C4N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-C6N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C6N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-C8N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C8N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-C10N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C10N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-CD4"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-CD6"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-CD8"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD8N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-CL4"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-CL6"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-CL8"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                                '"-06N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-06N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                                '"-08N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08N") <> 0 Then
                                    'MP付バルブブロックは価格加算なし
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-08N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                        End Select
                    End If
                Next
            Else
                '接続口径(継手エルボ)加算価格キー
                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                        If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3G") <> 0 And _
                           Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = "1" Then
                            strPortSize = "-S"
                        Else
                            strPortSize = ""
                        End If
                        Select Case intLoopCnt
                            Case 2 To 9
                                '電磁弁付バルブブロック＆MP付バルブブロック
                                '2013/10/16 修正
                                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) <> "MN4GE" Or _
                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) = "MN4GEX" Then
                                    '"-CL4","-CL4NC","-CL4NO"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CL6","-CL6NC","-CL6NO"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CL8","-CL8NC","-CL8NO"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-C18","-C18NC","-C18NO"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C18") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C18" & _
                                                                                           strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C18" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CL18"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL18") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL18" & _
                                                                                           strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL18" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD18"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD18") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD18" & _
                                                                                           strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD18" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD4"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                  (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CD4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CD4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CD4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD4" & _
                                                                                               strPortSize
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD4" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD6"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                  (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CD6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CD6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CD6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD6" & _
                                                                                               strPortSize
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD6" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD8"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                  (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CD8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CD8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CD8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD8" & _
                                                                                               strPortSize
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD8" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD10"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD10") <> 0 Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                  (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CD10"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                      (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CD10"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                       (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CD10"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD10" & _
                                                                                               strPortSize
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CD10" & _
                                                                                       strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CF"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CF") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CF" & _
                                                                                   strPortSize
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C3N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                              (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-C3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-C3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-C3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C3N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-C4N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                               (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-C4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-C4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-C4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C4N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-M5N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-M5N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-C3G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C3G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C4G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C4G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-M5G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-M5G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C6N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-C6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-C6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-C6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C6N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-C8N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-C8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-C8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-C8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C8N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-06N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-06N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C6G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C6G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C8G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C8G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-06G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-06G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C10N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C10N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-08N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-08N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C8"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C8" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C10"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-C10" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL3N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                 (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL3N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL3N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-CL4N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                 (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL4N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-CL4G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL4G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL6G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL6N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                 (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL6N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL6N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-CL8N"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8N") <> 0 And _
                                             Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) = "N" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPS") <> 0 And _
                                                 (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPS-CL8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MPD") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MPD-CL8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") <> 0 And _
                                                   (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GB" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN4GE" Or _
                                                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) = "MN3GE") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "R-MP-CL8N"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If

                                    '"-CL8G"
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                    '    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-CL8G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                        End Select
                    End If
                Next
            End If

            '2011/06/16 ADD RM1106028(7月VerUP:MN4G-ULシリーズ　価格積上げ) START --->
            If strHosyo.Trim = "UL" Then
                For intLoopCnt = 2 To 9
                    '仕様書形番が選択されていること、かつ、仕様書使用数が入っていること
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length <> 0 And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                        Select Case intLoopCnt
                            Case 2 To 9
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), "-MP") = 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "UL"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                        End Select
                    End If
                Next
            End If
            '2011/06/16 ADD RM1106028(7月VerUP:MN4G-ULシリーズ　価格積上げ) <--- END

            'RM1308014 2013/08/07 Y.Tachi
            '2013/08/21 修正
            'RM1310004 2013/10/01 追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN3GA1", "MN3GA2", "MN3GB1", "MN3GB2", "MN4GA1", "MN4GA2", "MN4GB1", "MN4GB2", _
                     "MN3GD1", "MN3GD2", "MN3GE1", "MN3GE2", "MN4GD1", "MN4GD2", "MN4GE1", "MN4GE2"
                    If strLion.Trim = "P4" Then
                        '接続口径(継手エルボ)加算価格キー
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                '3ポート弁で2個内蔵形でない場合はシングルの価格になる
                                If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3G") <> 0 And _
                                   Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = "1" Then
                                    strPortSize = "-S"
                                Else
                                    strPortSize = ""
                                End If
                                '3ポート弁で2個内蔵形でない場合はシングルの価格になる
                                If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3GA11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3GA21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA12") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA13") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA14") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA15") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA22") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA23") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA24") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GA25") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB12") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB13") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB14") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB15") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB22") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB23") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB24") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GB25") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3GD11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N3GD21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD12") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD13") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD14") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD15") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD22") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD23") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD24") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GD25") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE11") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE12") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE13") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE14") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE15") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE21") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE22") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE23") <> 0 Or _
                                    InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE24") <> 0 Or InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N4GE25") <> 0 Then
                                    'シングル価格キー
                                    Select Case intLoopCnt
                                        Case 2 To 9
                                            '電磁弁付バルブブロック＆MP付バルブブロック
                                            '"-C4","-C4NC","-C4NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-OP-P4-C4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                            '"-C6","-C6NC","-C6NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-OP-P4-C6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                            '"-C8","-C8NC","-C8NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) & "-OP-P4-C8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                    End Select
                                Else
                                    '3ポート弁2個内蔵型価格キー
                                    Select Case intLoopCnt
                                        Case 2 To 9
                                            '電磁弁付バルブブロック＆MP付バルブブロック
                                            '"-C4","-C4NC","-C4NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) & "-OP-P4-C4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                            '"-C6","-C6NC","-C6NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) & "-OP-P4-C6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                            '"-C8","-C8NC","-C8NO"
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Then
                                                'MP付バルブブロックは価格加算なし
                                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) & "-OP-P4-C8"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                    End Select
                                End If
                                '給排気ﾌﾞﾛｯｸ加算(P4)
                                Select Case intLoopCnt
                                    'Case 13 To 14                                    
                                    Case 15 To 17
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-Q") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                End Select
                            End If
                        Next

                    End If

            End Select

            '電圧
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                'Case "R", "U"
                Case "R", "U", "S", "V" 'RM1610013
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select



            '食品製造工程向け商品 RM1610013
            '電磁弁付バルブブロック＆MP付バルブブロック
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim.Contains("X12") Then
                    Dim intN1 As Integer = 0 'N4GB1
                    Dim intN2 As Integer = 0 'N4GB2
                    intN1 = intN3GA1_EV_OPT + intN3GB1_EV_OPT + intN4GA1_EV_OPT + intN4GB1_EV_OPT + intN4GA1_MP_OPT + intN4GB1_MP_OPT
                    intN2 = intN3GA2_EV_OPT + intN3GB2_EV_OPT + intN4GA2_EV_OPT + intN4GB2_EV_OPT + intN4GA2_MP_OPT + intN4GB2_MP_OPT
                    If intN1 > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N4GB1-FP1"
                        decOpAmount(UBound(decOpAmount)) = intN1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N4GB1R-V-FP1"
                        decOpAmount(UBound(decOpAmount)) = intN1
                    End If
                    If intN2 > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N4GB2-FP1"
                        decOpAmount(UBound(decOpAmount)) = intN2
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N4GB2R-V-FP1"
                        decOpAmount(UBound(decOpAmount)) = intN2
                    End If
                Else
                    '2016/10/31 追加
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-V-FP1"
                    decOpAmount(UBound(decOpAmount)) = strRensu
                End If

            End If

            '食品製造工程向け商品 RM1610013
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim.Contains("X12") Then
                Else
                    If strOptionFP1.Contains("FP1") Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-FP1"
                        'decOpAmount(UBound(decOpAmount)) = 1
                        decOpAmount(UBound(decOpAmount)) = strRensu
                    End If
                End If
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
