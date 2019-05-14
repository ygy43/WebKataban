'************************************************************************************
'*  ProgramID  ：KHPrice26
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*
'*  概要       ：ブロックマニホールド用電磁弁単体
'*             ：３ＧＡ１／３ＧＡ２／３ＧＡ３
'*             ：４ＧＡ１／４ＧＡ２／４ＧＡ３／４ＧＡ４
'*             ：３ＧＢ１／３ＧＢ２
'*             ：４ＧＢ１／４ＧＢ２／４ＧＢ３／４ＧＢ４
'*
'*【修正履歴】
'*                                      更新日：2007/05/09   更新者：NII A.Takahashi
'*  ・オプション「K」において、4GB4の場合のみの価格積上げロジック削除
'*                                      更新日：2007/09/26   更新者：NII A.Takahashi
'*  ・継手オプション追加により、継手加算ロジック追加
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　3GA1/3GA2/4GA1/4GA2/3GB1/3GB2/4GB1/4GB2にオプションボックス追加
'*  ・受付No：RM0904031  4GD2/4GE2機種追加
'*                                      更新日：2009/06/23   更新者：Y.Miura
'*  二次電池対応                         更新日：2010/05/25   更新者：Y.Miura
'************************************************************************************
Module KHPrice60

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean

        Dim strKiriIchikbn As String = ""   '切換位置区分
        Dim strSosakbn As String = ""       '操作区分
        Dim strKokei As String = ""         '接続口径
        Dim strSyudoSochi As String = ""    '手動装置
        Dim strDensen As String = ""        '電線接続
        Dim strTanshi As String = ""        '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strOption As String = ""        'オプション
        Dim strTaiki As String = ""         '大気開放タイプ
        Dim strDenatsu As String = ""       '電圧
        Dim strCleanShiyo As String = ""    'クリーン仕様
        Dim strHosyo As String = ""         '保証
        Dim strLion As String = ""          '二次電池

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'RM0904031 2009/06/23 Y.Miura
                'Case "3GA1", "3GA2", _
                '     "4GA1", "4GA2", _
                '     "3GB1", "3GB2", _
                '     "4GB1", "4GB2"
                Case "3GA1", "3GA2", _
                     "4GA1", "4GA2", _
                     "3GB1", "3GB2", _
                     "4GB1", "4GB2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子・ｺﾈｸﾀﾋﾟﾝ配列
                        strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          '手動装置
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 12 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(12).Trim           '二次電池
                        End If
                    Else
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                        strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim          '手動装置
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 10 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(10).Trim           '二次電池
                        End If
                    End If
                Case "3GD1", "3GD2", _
                     "4GD1", "4GD2", _
                     "3GE1", "3GE2", _
                     "4GE1"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim          '手動装置
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                        strTaiki = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 11 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(11).Trim           '二次電池
                        End If
                    Else
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                        strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim          '手動装置
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 10 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(10).Trim           '二次電池
                        End If
                    End If
                    '↓RM1310067 2013/10/23 追加
                Case "4GE2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim          '手動装置
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                            strTaiki = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim               '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 11 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(11).Trim           '二次電池
                            End If
                        Else
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                        End If
                    Else
                        'キー型番の変更、およびオプション数の変更に伴い、以下の内容を合わせて修正  2016/11/22 修正 松原
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                            'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                            strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(5).Trim          '手動装置
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim             '電圧
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          'クリーン仕様
                            strHosyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '保証
                            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 10 Then
                                strLion = objKtbnStrc.strcSelection.strOpSymbol(10).Trim           '二次電池
                            End If
                        Else
                            strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                            strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                            '以下一項目分ずらす  2016/11/22 修正 松原
                            strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                            strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                            strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim             '電圧
                        End If
                    End If
                Case "3GA3", "4GA3", "4GA4", _
                     "4GB3", "4GB4"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strTanshi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              '端子・ｺﾈｸﾀﾋﾟﾝ配列
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 11 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(11).Trim            '二次電池
                        End If
                    Else
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(6).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(8).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 9 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(9).Trim            '二次電池
                        End If
                    End If
                Case "3GD3", "4GD3","4GE3"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim             '電圧
                        strTaiki = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 10 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(10).Trim            '二次電池
                        End If
                    Else
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(6).Trim             '電圧
                        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          'クリーン仕様
                        strHosyo = objKtbnStrc.strcSelection.strOpSymbol(8).Trim               '保証
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 9 Then
                            strLion = objKtbnStrc.strcSelection.strOpSymbol(9).Trim            '二次電池
                        End If
                    End If
            End Select

            '基本価格キー
            '↓RM1310067 2013/10/23 追加
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                'キー型番の変更に伴い修正  2016/11/22 修正 松原
                Case "R", "S", "T"
                    'Case "R", "S"
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "4GE2"
                            'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                            If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & strSosakbn & "R"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & strKiriIchikbn & strSosakbn & "R-" & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "4GE2"
                            'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                            If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & strSosakbn
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & strKiriIchikbn & strSosakbn & "-" & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & strSosakbn
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            '配管ねじ加算価格キー
            Select Case Right(strKokei, 1)
                Case "G", "N"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        'If strKiriIchikbn = "66" Or strKiriIchikbn = "76" Or _
                        '    strKiriIchikbn = "77" Or strKiriIchikbn = "67" Then

                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "66R" & _
                        '                                               CdCst.Sign.Hypen & _
                        '                                               Right(strKokei, 1)
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                        'Else
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & _
                        '                                               CdCst.Sign.Hypen & _
                        '                                               Right(strKokei, 1)
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                        'End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   Right(strKokei, 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                    End If
            End Select

            '接続口径
            Select Case strKokei
                Case "C18", "CL18", "CD18", "CD4", "CD6", "CD8", "CD10", "CF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If InStr(objKtbnStrc.strcSelection.strSeriesKataban.Trim, "3G") <> 0 And _
                       (InStr(strKiriIchikbn, "1") <> 0 Or _
                       InStr(strKiriIchikbn, "11") <> 0) Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strKokei & CdCst.Sign.Hypen & "S"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strKokei
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '接続口径によってチェック区分を変える
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "4GA1", "3GA1"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C3N", "C4N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GA2", "3GA2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C8N", "C6N", "06N", "C4G", "C6G", "C8G", "06G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GA3", "3GA3"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C8N", "C10N", "08N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB1", "3GB1"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "06N", "06G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB2", "3GB2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "08N", "08G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB3", "3GB3"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        Select Case strKokei
                            Case "10N", "08N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

            '大気開放加算価格キー
            If strTaiki <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strTaiki
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'クリーン仕様加算価格キー
            If strCleanShiyo <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & _
                                                           strSosakbn & CdCst.Sign.Hypen & strCleanShiyo
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '電線接続・省配線接続加算価格キー
            If strDensen <> "" Then
                '↓RM1310067 2013/10/23 追加
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "4GE2"
                        'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                            'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strDensen
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & strDensen
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & strDensen
                End Select
                Select Case strKiriIchikbn
                    Case "1", "11"
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "66", "67", "76", "77", "2", "3", "4", "5"
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 2
                End Select
            End If

            '端子・ｺﾈｸﾀﾋﾟﾝ配列
            If strTanshi <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strTanshi
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "F"
                        Select Case strKiriIchikbn
                            Case "66", "67", "76", "77"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "DUAL"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "K"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "S", "E", "Q"     'オプション「Q」を同処理に追加 2017/01/17 追加

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        If strKiriIchikbn <> "1" And strKiriIchikbn <> "11" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "H"
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                            If strSosakbn = "9" Then
                            Else
                                '↓RM1310067 2013/10/23 追加
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                    Case "4GE2"
                                        'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                        If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                            'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                        Else
                            '↓RM1310067 2013/10/23 追加
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "4GE2"
                                    'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                    Case Else
                        '↓RM1310067 2013/10/23 追加
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "4GE2"
                                'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "T" Then
                                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim <> "1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            '電圧加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "4GA4", "4GB4"
                    If strDenatsu = "5" Then
                        If strKiriIchikbn = "1" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC(2)"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) START --->
            'RM1210067 2013/04/04 不具合対応
            'ＵＬ仕様加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "3GA1", "3GA2", "3GA3", "4GA1", "4GA2", "4GA3", _
                     "3GB1", "3GB2", "4GB1", "4GB2", "4GB3"
                    If strHosyo = "UL" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & strHosyo
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select
            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) <--- END

            '二次電池加算    'RM1005030 2010/05/25 Y.Miura 追加
            If strLion <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                Select Case strKiriIchikbn
                    Case "1", "11"
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-OP-" & strLion & CdCst.Sign.Hypen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "2", "3", "4", "5"
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-OP-" & strLion & CdCst.Sign.Hypen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "66", "67", "76", "77"
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strKiriIchikbn & _
                                                                   "-OP-" & strLion & CdCst.Sign.Hypen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            '電圧
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select

            'オプション(H)
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
               objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                If strSosakbn = "9" Then
                    If Not strOption.Contains("H") Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module