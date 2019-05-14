'************************************************************************************
'*  ProgramID  ：KHPrice61
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド用電磁弁(マニホールド)
'*             ：Ｍ３ＧＡ１／Ｍ３ＧＡ２／Ｍ３ＧＡ３
'*             ：Ｍ４ＧＡ１／Ｍ４ＧＡ２／Ｍ４ＧＡ３／Ｍ４ＧＡ４
'*             ：Ｍ３ＧＢ１／Ｍ３ＧＢ２
'*             ：Ｍ４ＧＢ１／Ｍ４ＧＢ２／Ｍ４ＧＢ３／Ｍ４ＧＢ４
'*
'* 更新履歴 　 ：
'*                                      更新日：2007/09/26   更新者：NII A.Takahashi
'*  ・デュアル3ポート/継手追加に伴い、継手加算ロジック修正
'*  ・M4G*4において、DINレールが正しく加算されていない箇所修正
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　M3GA1/M3GA2/M4GA1/M4GA2/
'*                           M3GB1/M3GB2/M4GB1/M4GB2にオプションボックス追加
'*  ・受付No：RM0904031  4GD2/4GE2機種追加
'*                                      更新日：2009/06/23   更新者：Y.Miura
'*    二次電池対応                       RM1005030 2010/05/25 Y.Miura
'************************************************************************************
Module KHPrice61

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean

        Dim intStationQty As Integer = 0
        Dim intQuantity As Integer = 0
        Dim intValveQty As Integer = 0
        Dim intValveQtyDual As Integer = 0
        Dim intValveQty3P As Integer = 0
        Dim intValveQty4P As Integer = 0
        Dim intValveQty1SWD As Integer = 0
        Dim intValveQty2SWD As Integer = 0
        Dim intValveQty3SWD As Integer = 0
        Dim intValveQty4SWD As Integer = 0
        Dim intValveQty5SWD As Integer = 0

        Dim strPortSize As String

        Dim strKiriIchikbn As String = ""       '切換位置区分
        Dim strSosakbn As String = ""           '操作区分
        Dim strKokei As String = ""             '接続口径
        Dim strSyudoSochi As String = ""        '手動装置
        Dim strDensen As String = ""            '電線接続
        Dim strTanshi As String = ""            '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strOption As String = ""            'オプション
        Dim strMountType As String = ""         'マウントタイプ
        Dim strPilotType As String = ""         'パイロットタイプ
        Dim strRensu As String = ""             '連数
        Dim strDenatsu As String = ""           '電圧
        Dim strCleanShiyo As String = ""        'クリーン仕様
        Dim strHosyo As String = ""             '保証
        Dim strLion As String = ""              '二次電池

        Dim strABDE As String                   'RM1005030 2010/05/17 Y.Miura
        Dim str1234 As String                   'RM1005030 2010/05/17 Y.Miura  

        Try
            '機種の文字列を取り出す       'RM1005030 追加
            strABDE = objKtbnStrc.strcSelection.strSeriesKataban.Trim.PadRight(5, " ").Substring(3, 1)
            str1234 = objKtbnStrc.strcSelection.strSeriesKataban.Trim.PadRight(6, " ").Substring(4, 1)

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "R", "U", "S", "V"
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "M3GA1", "M3GB1", _
                             "M4GA1", "M4GB1", _
                             "M3GA2", "M3GB2", _
                             "M4GA2", "M4GB2"

                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          '手動装置
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(9).Trim           'マウントタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(11).Trim            '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim         'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(13).Trim              '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 14 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(14).Trim           '二次電池
                            End If
                        Case "M3GD1", "M3GE1", _
                             "M4GD1", "M4GE1", _
                             "M3GD2", "M3GE2", _
                             "M4GD2", "M4GE2"

                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          '手動装置
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(9).Trim           'マウントタイプ
                            strPilotType = objKtbnStrc.strcSelection.strOpSymbol(10).Trim          'パイロットタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(11).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(12).Trim            '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(13).Trim         'クリーン仕様
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 14 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(14).Trim           '二次電池
                            End If

                        Case "M3GA3", "M4GA3", "M4GB3"
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(8).Trim           'マウントタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim             '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim         'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim              '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 13 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(13).Trim           '二次電池
                            End If

                        Case "M3GD3", "M4GD3", "M4GE3"
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(8).Trim           'マウントタイプ
                            strPilotType = objKtbnStrc.strcSelection.strOpSymbol(9).Trim          'パイロットタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(11).Trim             '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim         'クリーン仕様
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 13 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(13).Trim           '二次電池
                            End If

                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        'RM0904031 2009/06/23 Y.Miura
                        'Case "M3GA1", "M3GA1    T", "M3GB1", "M3GB1    T", _
                        '     "M4GA1", "M4GA1    T", "M4GB1", "M4GB1    T", _
                        '     "M3GA2", "M3GA2    T", "M3GB2", "M3GB2    T", _
                        '     "M4GA2", "M4GA2    T", "M4GB2", "M4GB2    T"
                        Case "M3GA1", "M3GA1    T", "M3GB1", "M3GB1    T", _
                             "M4GA1", "M4GA1    T", "M4GB1", "M4GB1    T", _
                             "M3GA2", "M3GA2    T", "M3GB2", "M3GB2    T", _
                             "M4GA2", "M4GA2    T", "M4GB2", "M4GB2    T", _
                             "M3GD1", "M3GD1    T", "M3GE1", "M3GE1    T", _
                             "M4GD1", "M4GD1    T", "M4GE1", "M4GE1    T", _
                             "M3GD2", "M3GD2    T", "M3GE2", "M3GE2    T", _
                             "M4GD2", "M4GD2    T", "M4GE2", "M4GE2    T"

                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim          '手動装置
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(8).Trim           'マウントタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim            '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim         'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(12).Trim              '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 13 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(13).Trim           '二次電池
                            End If

                            'RM0904031 2009/06/23 Y.Miura
                            'Case "M3GA3", "M3GA3    T", _
                            '     "M4GA3", "M4GA3    T", "M4GA4", "M4GA4    T", _
                            '     "M4GB3", "M4GB3    T", "M4GB4", "M4GB4    T"
                        Case "M3GA3", "M3GA3    T", _
                             "M4GA3", "M4GA3    T", "M4GA4", "M4GA4    T", _
                             "M4GB3", "M4GB3    T", "M4GB4", "M4GB4    T", _
                             "M3GD3", "M3GD3    T", _
                             "M4GD3", "M4GD3    T", "M4GD4", "M4GD4    T", _
                             "M4GE3", "M4GE3    T", "M4GE4", "M4GE4    T"

                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                            strTanshi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '端子･ｺﾈｸﾀﾋﾟﾝ配列
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                            strMountType = objKtbnStrc.strcSelection.strOpSymbol(7).Trim           'マウントタイプ
                            strRensu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim               '連数
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim         'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim              '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 12 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(12).Trim           '二次電池
                            End If

                    End Select
            End Select


            'バルブブロック連数
            intStationQty = CInt(strRensu)

            'サブプレート価格キー
            If Left(strDensen, 1) = "T" Then
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    'RM0904031 2009/06/23 Y.Miura
                    'Case "M4GB4"
                    Case "M4GB4", "M4GE4"
                        If InStr(strKokei, "15") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SPT-" & _
                                                                       strRensu & _
                                                                       CdCst.Sign.Hypen & "15"

                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SPT-" & _
                                                                       strRensu
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SPT-" & _
                                                                   strRensu
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Else
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    'RM0904031 2009/06/23 Y.Miura
                    'Case "M4GB4"
                    Case "M4GB4", "M4GE4"
                        If InStr(strKokei, "15") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SP-" & _
                                                                       strRensu & "-15"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SP-" & _
                                                                       strRensu
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SP-" & _
                                                                   strRensu
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'サブプレート・クリーンルーム仕様加算価格キー
            If strCleanShiyo = "P70" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SP-" & _
                                                           strCleanShiyo & CdCst.Sign.Hypen & _
                                                           strRensu
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        '2010/09/17 RM1009006(10月VerUP:不具合対応 チューブ抜具オプション加算不正) START --->
                        Case CdCst.Manifold.InspReportJp.SelectValue, _
                             CdCst.Manifold.InspReportEn.SelectValue, _
                             CdCst.Manifold.TubeRemover.Necessity, _
                             CdCst.Manifold.TubeRemover.UnNecessity, _
                            CdCst.Manifold.InspReportJp.English, CdCst.Manifold.InspReportEn.English, _
                            CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportEn.Japanese
                            'Case CdCst.Manifold.InspReportJp.SelectValue, _
                            '     CdCst.Manifold.InspReportEn.SelectValue
                            '2010/09/17 RM1009006(10月VerUP:不具合対応 チューブ抜具オプション加算不正) <--- END
                            '加算なし
                        Case Else
                            'RM1803032_スペーサ行追加対応
                            Select Case intLoopCnt
                                Case 1 To 10
                                    '電磁弁

                                    'G1・G2/X・X1が入っている場合は別の価格キーを付与するよう処理を追加 RM1702017 追加
                                    'If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G1") = 0 Then
                                    '    'G1が入っていない場合は、GP2が入っているかどうかをチェック
                                    '    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G2") = 0 Then
                                    '        'G2も入っていなかったら従来通り
                                    '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1)
                                    '    Else
                                    '        'G2が入っていたら末尾に「-G2」を付与
                                    '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G2"
                                    '    End If
                                    'Else
                                    '    'G1が入っている場合はG1を付与
                                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G1"
                                    'End If

                                    'G1・G2/X・X1が入っている場合は別の価格キーを付与するよう処理を追加 RM1702017 追加
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G1") >= 1 Then
                                        'G1が入っている場合はG1を付与
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G1"
                                        'G1が入っていない場合は、GP2が入っているかどうかをチェック
                                    ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-G2") >= 1 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-G2"
                                    ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X1") >= 1 Then
                                        'X1が入っていたら末尾に「-X1」を付与
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-X1"
                                    ElseIf InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-X") >= 1 Then
                                        'Xが入っていたら末尾に「-X」を付与
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-X"
                                    Else
                                        'G2も入っていなかったら従来通り
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1)
                                    End If

                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1) = "1" Then
                                        intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End If

                                    '電磁弁数(バルブ数)をカウントする①
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1)
                                        Case "1"
                                            intValveQty1SWD = intValveQty1SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
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

                                    Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1)
                                        Case "3"
                                            intValveQty3P = intValveQty3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "4"
                                            intValveQty4P = intValveQty4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                    '電磁弁数(デュアル3ポート弁のバルブ数)をカウントする③
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                        Case "66", "67", "76", "77"
                                            intValveQtyDual = intValveQtyDual + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
                                Case 11 To 12
                                    'マスキングプレート
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        'RM0904031 2009/06/23 Y.Miura
                                        'Case "M4GB4", "M4GA4"
                                        Case "M4GA4", "M4GE4", "M4GD4"
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-MP"
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            End If
                                        Case "M4GB4"
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-MP"
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = "4GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 3)
                                            End If
                                        Case Else
                                            Select Case Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim)
                                                Case 6
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                                Case 7
                                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "R" And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "U" _
                                                        And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "S" And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "V" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 3)
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "R-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2)
                                                    End If
                                                Case 8
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "R-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 3)
                                                Case Else
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7, 1)
                                                        Case "-"
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                                        Case Else
                                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "R" And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "U" _
                                                                And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "S" And objKtbnStrc.strcSelection.strKeyKataban.Trim <> "V" Then
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 3)
                                                            ElseIf Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, 1) = "-" Then
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "R-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2)
                                                            Else
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "R-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 3)
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case 13 To 16
                                    '個別給排気スペーサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                Case 17 To 20
                                    'ブランクプラグ,サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                Case 21
                                    'ねじプラグ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") + 1, Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim))
                                    'Case 20
                                    '    'DINレール
                                    '    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                    '        Case "M4GB4"
                                    '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-QD-" & strKokei
                                    '        Case Else
                                    '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '            strOpRefKataban(UBound(strOpRefKataban)) = "4G-BAA"
                                    '    End Select
                                Case 23 To 24
                                    'Case 22 To 23
                                    'ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            End Select


                            'Select Case Left(strOpRefKataban(UBound(strOpRefKataban)), 6)
                            '    Case "4G-BAA"
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                            'End Select

                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4GB4" And _
                               InStr(strOpRefKataban(UBound(strOpRefKataban)), "M4GB4-QD") <> 0 Then
                                'M4GB4はDINﾚｰﾙの長さに関わらず、一律価格
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

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

                            Select Case intLoopCnt
                                Case 1 To 10
                                    'AC110Vの時、電圧加算
                                    If strDenatsu = "5" Then
                                        If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "419") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC(2)"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                            End Select

                            'クリーンルーム仕様加算価格キー
                            If strCleanShiyo = "P70" Then
                                Select Case intLoopCnt
                                    Case 1 To 10
                                        '電磁弁(クリーンルーム仕様加算)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1) & "-" & strCleanShiyo
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 11 To 12
                                        'マスキングプレート(クリーンルーム仕様加算)
                                        Select Case Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim)
                                            Case 6
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2) & "-" & strCleanShiyo
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case 7
                                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2) & "-" & strCleanShiyo
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 3) & "-" & strCleanShiyo
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            Case Else
                                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                    If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) = 8 Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 3) & "-" & strCleanShiyo
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    Else
                                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, 1)
                                                            Case "-"
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 2) & "-" & strCleanShiyo
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                            Case Else
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 3) & "-" & strCleanShiyo
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End Select
                                                    End If
                                                Else
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7, 1)
                                                        Case "-"
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2) & "-" & strCleanShiyo
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Case Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = "4G" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 3) & "-" & strCleanShiyo
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select
                End If
            Next

            'DINレール
            If objKtbnStrc.strcSelection.decDinRailLength <> 0 Then
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    'RM0904031 2009/06/23 Y.Miura
                    'Case "M4GB4"
                    Case "M4GB4", "M4GE4"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-QD-" & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                        'strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "4G-BAA"
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                End Select
            End If

            '電線接続／省配線接続加算価格キー
            If strDensen <> "" Then
                If Left(strDensen, 1) = "T" Then
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & _
                                                                   strDensen
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strDensen
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    '機種名-A2Nキーを作る
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-A2N"
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               strDensen
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "K"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                    Case "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                        '    'RM1709013 オプション追加（X,X1）
                        'Case "X"
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                               strOpArray(intLoopCnt).Trim
                        '    decOpAmount(UBound(decOpAmount)) = intValveQty
                        'Case "X1"
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                               strOpArray(intLoopCnt).Trim
                        '    decOpAmount(UBound(decOpAmount)) = intValveQty
                    Case "H"
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                            If intValveQty3SWD <> 0 Or intValveQty5SWD <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                                decOpAmount(UBound(decOpAmount)) = intValveQty3SWD + intValveQty5SWD
                            End If
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intValveQty - (intValveQty3SWD + intValveQty5SWD)
                        End If
                    Case "Q"
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "M4GB4"
                                If strMountType <> "D" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               strKokei
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case Else
                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" _
                                    Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    'ダブルソレノイドは２倍加算
                                    decOpAmount(UBound(decOpAmount)) = intValveQty + intValveQty2SWD + intValveQtyDual + intValveQty3SWD + intValveQty5SWD + intValveQty4SWD
                                Else

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                End If
                        End Select
                    Case "S", "E"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        decOpAmount(UBound(decOpAmount)) = intValveQty + intValveQty2SWD + intValveQtyDual + intValveQty3SWD + intValveQty5SWD + intValveQty4SWD

                    Case "F"
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            'RM0904031 2009/06/23 Y.Miura
                            'Case "M3GA1"
                            Case "M3GA1", "M3GD1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M3GA2"
                            Case "M3GA2", "M3GD2"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M3GA3"
                            Case "M3GA3", "M3GD3"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GA1"
                            Case "M4GA1", "M4GD1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA1-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim()
                                decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim() & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M4GA1-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim()
                                decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GA2"
                            Case "M4GA2", "M4GD2"
                                If intValveQty3P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'RM0904031 2009/06/23 Y.Miura
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & _
                                    '                                                    strOpArray(intLoopCnt).Trim
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual
                                End If

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA2" & "-" & _
                                '                                           strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual

                                If intValveQty4P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M4GA" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                End If
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GA3"
                            Case "M4GA3", "M4GD3"
                                If intValveQty3P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty3P
                                End If

                                If intValveQty4P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M4GA" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty4P
                                End If
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GA4"
                            Case "M4GA4", "M4GD4"
                                If intValveQty4P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M4GA" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & "-" & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty
                                End If
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M3GB1"
                            Case "M3GB1", "M3GE1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M3GB2"
                            Case "M3GB2", "M3GE2"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GB2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GB1"
                            Case "M4GB1", "M4GE1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GB1-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M4GB1-" & strOpArray(intLoopCnt).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intStationQty - intValveQtyDual
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GB2"
                            Case "M4GB2", "M4GE2"
                                If intValveQty3P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M3GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & CdCst.Sign.Hypen & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty3P - intValveQtyDual
                                End If

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M3GA2-" & strOpArray(intLoopCnt).Trim & "-DUAL"
                                strOpRefKataban(UBound(strOpRefKataban)) = "M3G" & strABDE & str1234 & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & "-DUAL"
                                decOpAmount(UBound(decOpAmount)) = intValveQtyDual

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M4GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & CdCst.Sign.Hypen & _
                                strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GB3"
                            Case "M4GB3", "M4GE3"
                                If intValveQty3P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M4GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & CdCst.Sign.Hypen & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intValveQty3P
                                End If

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = "M4GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & CdCst.Sign.Hypen & _
                                strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                'RM0904031 2009/06/23 Y.Miura
                                'Case "M4GB4"
                            Case "M4GB4", "M4GE4"
                                If intValveQty4P <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    'strOpRefKataban(UBound(strOpRefKataban)) = "M4GB" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) & CdCst.Sign.Hypen & _
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M4G" & strABDE & str1234 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = intStationQty
                                End If
                        End Select
                    Case "Z4", "Z5"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "4G-" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                If Not strOption.Contains("H") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                End If
            End If

            '大気開放加算価格キー
            If strPilotType <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strPilotType
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '端子・コネクタピン加算価格キー
            If strTanshi = "W1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strTanshi
                decOpAmount(UBound(decOpAmount)) = intValveQty1SWD
            End If

            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) START --->
            If strHosyo = "UL" Then
                For intLoopCnt = 1 To 10
                    '仕様書形番が選択されていること、かつ、仕様書使用数が入っていること
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length <> 0 And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                        Select Case intLoopCnt
                            Case 1 To 10
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & strHosyo
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        End Select
                    End If
                Next
            End If
            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) <--- END

            '2010/08/24 ADD RM1008009(9月VerUP:M4G-P4シリーズ　価格積上げ) START --->
            '二次電池仕様時
            If strLion <> "" Then
                '外部ﾊﾟｲﾛｯﾄ(K)加算価格キー 
                Select Case strOption
                    Case "K"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & strLion & "-" & strOption
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select

            End If
            '2010/08/24 ADD RM1008009(9月VerUP:M4G-P4シリーズ　価格積上げ) <--- END


            If strHosyo = "UL" Then
                '接続口径(継手エルボ)加算価格キー
                Dim strPartKataban As String = String.Empty
                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length <> 0 And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then

                        '3ポート弁で2個内蔵形でない場合はシングルの価格になる
                        If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GA11") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GA21") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GB11") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GB21") <> 0 Then
                            strPortSize = "-S"
                        Else
                            strPortSize = ""
                        End If
                        Select Case intLoopCnt
                            Case 1 To 10
                                '電磁弁付バルブブロック
                                '"-CD3"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD3") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD3") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD3") <> 0 Then
                                    '"-CD3N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD3N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD3N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD3N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD3N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD3"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If
                                '"-CD4"
                                If objKtbnStrc.strcSelection.strKeyKataban = "R" And _
                                    Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4") <> 0 Then
                                        '"-CD4N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CD4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                Else
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4") <> 0 Then
                                        '"-CD4N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                End If
                                '"-CD6"
                                If objKtbnStrc.strcSelection.strKeyKataban = "R" And _
                                    Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6") <> 0 Then
                                        '"-CD6N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CD6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                Else
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6") <> 0 Then
                                        '"-CD6N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                End If
                                '"-CD8"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD8") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD8") <> 0 Then
                                    '"-CD8N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD8N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD8N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD8N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD8"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If
                                '"-CD10"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD10") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD10") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD10") <> 0 Then
                                    '"-CD10N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD10N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD10N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD10N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD10N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD10"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If
                                '"-CL3"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL3") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3") <> 0 Then
                                    '"-CL3N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL3N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL3N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL3"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If
                                '"-CL4"
                                If objKtbnStrc.strcSelection.strKeyKataban = "R" And _
                                    Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4") <> 0 Then
                                        '"-CL4N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CL4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                Else
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4") <> 0 Then
                                        '"-CL4N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL4N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                End If
                                '"-CL6"
                                If objKtbnStrc.strcSelection.strKeyKataban = "R" And _
                                    Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6") <> 0 Then
                                        '"-CL6N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CL6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                Else
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6") <> 0 Then
                                        '"-CL6N"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL6N"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL6"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                    End If
                                End If
                                '"-CL8"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL8") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8") <> 0 Then
                                    '"-CL8N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL8N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL8N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL8"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If
                                '"-CL10"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL10") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL10") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL10") <> 0 Then
                                    '"-CL10N"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL10N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL10N") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL10N") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL10N"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL10"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If

                                End If

                                '"-C3N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C3N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-C3N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-C4N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C4N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-C4N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-C6N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C6N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-C6N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-C8N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C8N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-C8N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-C10N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C10N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-C10N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-06N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "06N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-06N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                                '"-08N"
                                If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "08N") <> 0 Or _
                                   InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "08N") <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3) & "-08N"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                        End Select
                    End If
                Next
            Else
                '接続口径(継手エルボ)加算価格キー
                'Dim bol4Chara As Boolean = False
                Dim strPartKataban As String = String.Empty
                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length <> 0 And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then

                        '3ポート弁で2個内蔵形でない場合はシングルの価格になる
                        If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GA11") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GA21") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GB11") <> 0 Or _
                           InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "3GB21") <> 0 Then
                            strPortSize = "-S"
                        Else
                            strPortSize = ""
                        End If

                        If strLion = "" Or strLion = "FP1" Then
                            '既存（二次電池加算無し）
                            Select Case intLoopCnt
                                Case 1 To 10
                                    '電磁弁付バルブブロック
                                    '"-CL4"
                                    If (objKtbnStrc.strcSelection.strKeyKataban = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U") And _
                                         Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Or _
                                            InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4") <> 0 Or _
                                            InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4") <> 0 Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CL4"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    Else

                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                                    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL4"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL4"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If
                                    '"-CL6"
                                    If (objKtbnStrc.strcSelection.strKeyKataban = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U") And _
                                        Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6") <> 0 Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CL6"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    Else
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                                    objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL6"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL6"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If
                                    '"-CL8"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL8"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CL10"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL10") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL10"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL10"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CL18"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL18") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL18"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CL18"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-C18"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C18") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-C18" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD18"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD18") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD18") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD18" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD4"
                                    If (objKtbnStrc.strcSelection.strKeyKataban = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U") And _
                                        Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CD4" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    Else
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD4" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD6"
                                    If (objKtbnStrc.strcSelection.strKeyKataban = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U") And _
                                        Right(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "E1" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6") <> 0 Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "R-CD6" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    Else

                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6") <> 0 Then
                                            If (objKtbnStrc.strcSelection.strKeyKataban = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U") Then
                                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD6" & strPortSize
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If

                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD6" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If
                                    '"-CD8"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD8") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD8" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD8" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CD10"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD10") <> 0 Then
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1) <> "N" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD10" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CD10" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-CF"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CF") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CF") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CF") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-CF" & strPortSize
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C3N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C3N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C4N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C4N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-M5N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-M5N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "M5N") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-M5N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                    '"-C3G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C3G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C4G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C4G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-M5G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-M5G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "M5G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-M5G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C6N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C6N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C8N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-06N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-06N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C6G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C6G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C8G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-06G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-06G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C10N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C10N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-08N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-08N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "08N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-08NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "08NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-08NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "08NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-08N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C8"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-C10"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C10" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL3N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL3N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL3NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL3NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL3N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL4N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL4N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL4G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL4G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL6G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL6G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL6N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6N") <> 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6NO") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6NO") = 0) And _
                                           (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6NC") = 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL6N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL8N"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8N") <> 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8N") <> 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8N") <> 0) And _
                                          (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8NO") = 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8NO") = 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8NO") = 0) And _
                                          (InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8NC") = 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8NC") = 0 Or _
                                          InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8NC") = 0) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL8N" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If

                                    '"-CL8G"
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8G") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8G") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL8G" & strPortSize
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                            End Select
                        Else
                            '二次電池加算
                            'RM1310067 2013/10/23 追加
                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).PadRight(5, " ").Substring(0, 5).Trim
                                Case "3GA11", "3GA21", "3GB11", "3GB21", "4GA11", "4GA12", _
                                     "4GA13", "4GA14", "4GA15", "4GA21", "4GA22", "4GA23", _
                                     "4GA24", "4GA25", "4GB11", "4GB12", "4GB13", "4GB14", _
                                     "4GB15", "4GB21", "4GB22", "4GB23", "4GB24", "4GB25", _
                                     "4GB31", "4GB32", "4GB33", "4GB34", "4GB35", "4GB41", _
                                     "4GB42", "4GB43", "4GB44", "4GB45", "4GA31", "4GA32", _
                                     "4GA33", "4GA34", "4GA35", "4GA41", "4GA42", "4GA43", _
                                     "4GA44", "4GA45", "3GA31", _
                                     "3GD11", "3GD21", "3GE11", "3GE21", "4GD11", "4GD12", _
                                     "4GD13", "4GD14", "4GD15", "4GD21", "4GD22", "4GD23", _
                                     "4GD24", "4GD25", "4GE11", "4GE12", "4GE13", "4GE14", _
                                     "4GE15", "4GE21", "4GE22", "4GE23", "4GE24", "4GE25", _
                                     "4GE31", "4GE32", "4GE33", "4GE34", "4GE35", "4GE41", _
                                     "4GE42", "4GE43", "4GE44", "4GE45", "3GD31", "4GD31", _
                                     "4GD32", "4GD33", "4GD34", "4GD35"
                                    '先頭4文字をセット
                                    'bol4Chara = True
                                    strPartKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Substring(0, 4).Trim
                                Case Else
                                    '先頭6文字をセット
                                    'bol4Chara = False
                                    strPartKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).PadRight(6, " ").Substring(0, 6).Trim
                            End Select

                            Select Case intLoopCnt
                                Case 1 To 10
                                    '電磁弁付バルブブロック
                                    '"-C4"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C4") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-C4"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C6"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C6") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-C6"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C8"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C8") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-C8"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C10"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C10") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-C10"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-C12"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C12") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C12") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C12") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-C12"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    '"-M5"
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "M5") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "M5") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strPartKataban & "-OP-P4-M5"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                            End Select

                        End If


                        '共通
                        Select Case intLoopCnt
                            Case 11 To 12
                                'MP付バルブブロック
                                'Bタイプのみ加算する
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                    Case "M3GB1", "M3GB2", "M4GB1", "M4GB2", "M4GB3", "M3GE1", "M3GE2", "M4GE1", "M4GE2", "M4GE3"
                                        '"-CL4"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL4"
                                                '↓RM1210067 2012/11/05 修正
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CL4"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL4"
                                                '↓RM1210067 2012/11/05 修正
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL4"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CL6"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL6"
                                                '↓RM1210067 2012/11/05 修正
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CL6"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL6"
                                                '↓RM1210067 2012/11/05 修正
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL6"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CL8"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL8"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4) & "R-CL8"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL8"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4) & "-CL8"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CL10"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL10") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL10"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CL10"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL10"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL10"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CL18"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CL18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL18") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL18"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CL18"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CL18"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL18"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-C18"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "C18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C18") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-C18" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-C18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-C18" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CD18"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD18") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD18") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD18" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CD18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD18" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CD18" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CD4"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD4") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD4") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD4" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CD4" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD4" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CD4" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CD6"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD6") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD6") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD6" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4) & "R-CD6" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD6" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4) & "-CD6" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CD8"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD8") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD8") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD8" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CD8" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD8" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CD8" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CD10"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CD10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CD10") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CD10") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD10" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CD10" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CD10" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CD10" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-CF"
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CF") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "CF") <> 0 Or _
                                           InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CF") <> 0 Then
                                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "R-CF" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CF" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C3N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C3N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C4N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C4N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-M5N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-M5N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "M5N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-M5N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                        '"-C3G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C3G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C3G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C3G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C3G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C4G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C4G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C4G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C4G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C4G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-M5G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-M5G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-M5G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "M5G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-M5G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C6N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C6N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C8N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-06N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-06N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C6G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C6G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C6G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C6G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C6G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C8G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-06G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-06G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-06G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "06G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-06G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C10N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C10N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-08N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-08N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-08N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "08N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-08N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C8"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C8") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C8") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C8") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C8" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-C10"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-C10") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-C10") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "C10") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-C10" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL3N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL3N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL3N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL3N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL3N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL4N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL4N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL4G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL4G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL4G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL4G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL4G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL6G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL6G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL6N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL6N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL6N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL6N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL8N"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8N") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8N") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL8N" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If

                                        '"-CL8G"
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                            objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-CL8G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim, "-CL8G") <> 0 Or _
                                               InStr(1, objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, "CL8G") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                'strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 4) & "-CF" & strPortSize
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-CL8G" & strPortSize
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                End Select
                        End Select
                    End If
                Next
            End If

            If strLion = "FP1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-FP1"
                decOpAmount(UBound(decOpAmount)) = strRensu
            End If

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "M4GA4", "M4GB4"
                    '接続口径(G/N)加算価格キー
                    Select Case Right(strKokei, 1)
                        Case "G", "N"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       Right(strKokei, 1)
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "M4GA4"
                                    decOpAmount(UBound(decOpAmount)) = intValveQty + 1
                                Case "M4GB4"
                                    decOpAmount(UBound(decOpAmount)) = intValveQty
                            End Select
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                    End Select
            End Select

            '電圧
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R", "U", "S", "V"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
