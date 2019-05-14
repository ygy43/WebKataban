'************************************************************************************
'*  ProgramID  ：KHPrice70
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド電磁弁付バルブブロック単品
'*             ：Ｎ３ＧＡ１・２／Ｎ３ＧＢ１・２／Ｎ４ＧＡ１・２／Ｎ４ＧＢ１・２
'*             ：Ｎ３ＧＤ１・２／Ｎ３ＧＥ１・２／Ｎ４ＧＤ１・２／Ｎ４ＧＥ１・２
'*
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　N3GA1/N3GA2/N4GA1/N4GA2/N3GB1/N3GB2/N4GB1/N4GB2にオプションボックス追加
'************************************************************************************
Module KHPrice70

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intQuantity As Integer

        Dim strKiriIchikbn As String = ""   '切換位置区分
        Dim strSosakbn As String = ""       '操作区分
        Dim strKokei As String = ""         '接続口径
        Dim strCable As String = ""         'ケーブル長さ
        Dim strTanshi As String = ""        '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strSyudoSochi As String = ""    '手動装置
        Dim strDensen As String = ""        '電線接続
        Dim strOption As String = ""        'オプション
        Dim strDenatsu As String = ""       '電圧
        Dim strCleanShiyo As String = ""    'クリーン仕様
        Dim strOptionFP1 As String = ""     '食品製造工程向け商品 RM1610013

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                'Case "R"
                Case "R", "S" 'RM1610013
                    If objKtbnStrc.strcSelection.strSeriesKataban.Contains("GD") Or _
                       objKtbnStrc.strcSelection.strSeriesKataban.Contains("GE") Then
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strCable = objKtbnStrc.strcSelection.strOpSymbol(6).Trim                'ケーブル長さ
                        strTanshi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(8).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(9).Trim             '電圧
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 10 Then
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10).Trim          'クリーン仕様
                        End If
                    Else
                        strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                        strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                        strKokei = objKtbnStrc.strcSelection.strOpSymbol(4).Trim               '接続口径
                        strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim              '電線接続
                        strCable = objKtbnStrc.strcSelection.strOpSymbol(6).Trim                'ケーブル長さ
                        strTanshi = objKtbnStrc.strcSelection.strOpSymbol(7).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                        strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(8).Trim          '手動装置
                        strOption = objKtbnStrc.strcSelection.strOpSymbol(9).Trim              'オプション
                        strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(10).Trim             '電圧
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 11 Then
                            strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11).Trim          'クリーン仕様
                        End If
                        If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 13 Then              'RM1610013 Start
                            strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(13).Trim        '食品製造工程向け 
                        End If                                                                   'RM1610013 End
                    End If
                Case Else
                    strKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        '切換位置区分
                    strSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '操作区分
                    strKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim               '接続口径
                    strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).Trim              '電線接続
                    strCable = objKtbnStrc.strcSelection.strOpSymbol(5).Trim                'ケーブル長さ
                    strSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(6).Trim          '手動装置
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7).Trim              'オプション
                    strDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim             '電圧
                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(9).Trim          'クリーン仕様
            End Select

            '数量設定
            Select Case strKiriIchikbn
                Case "1"
                    intQuantity = 1
                Case "11"
                    intQuantity = 1
                Case "66", "67", "76", "77"
                    intQuantity = 2
                Case "2"
                    intQuantity = 2
                Case "3"
                    intQuantity = 2
                Case "4"
                    intQuantity = 2
                Case "5"
                    intQuantity = 2
            End Select

            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                Case "N4GE", "N3GE"
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then 'RM1610013
                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R" & CdCst.Sign.Hypen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        If (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "N4GE" And strKiriIchikbn = "1") Or _
                           (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "N3GE" And strKiriIchikbn = "1") Then
                            If Left(strKokei, 2) = "CL" Then
                                If Left(strDensen, 1) = "A" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-A2N" & _
                                                                               strCable & "-CL"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-CL"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                If Left(strDensen, 1) = "A" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-A2N" & _
                                                                               strCable & "-C"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-C"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            End If
                        Else
                            If Left(strDensen, 1) = "A" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                           strKiriIchikbn & _
                                                                           strSosakbn & "-A2N" & _
                                                                           strCable
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                           strKiriIchikbn & _
                                                                           strSosakbn
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If
                Case Else
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then 'RM1610013
                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R" & CdCst.Sign.Hypen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else

                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & CdCst.Sign.Hypen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            'クリーン仕様加算価格キー
            If strCleanShiyo <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           strKiriIchikbn & _
                                                           strSosakbn & CdCst.Sign.Hypen & _
                                                           strCleanShiyo
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '接続口径(継手エルボ)加算価格キー
            'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then 'RM1610013
                Select Case True
                    Case Left(strKokei, 2) = "CL" Or _
                         Left(strKokei, 2) = "CD" Or _
                         Left(strKokei, 2) = "CF" Or _
                         Left(strKokei, 3) = "C18" Or _
                         Right(strKokei, 1) = "N" Or _
                         Right(strKokei, 1) = "G"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        If InStr(objKtbnStrc.strcSelection.strSeriesKataban.Trim, "3G") <> 0 And _
                           (InStr(strKiriIchikbn, "1") <> 0 Or _
                            InStr(strKiriIchikbn, "11") <> 0) Then

                            If strDensen.Trim = "A2N" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & _
                                           strKokei & CdCst.Sign.Hypen & "S-A2N"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R" & CdCst.Sign.Hypen & _
                                           strKokei & CdCst.Sign.Hypen & "S"
                            End If

                        Else
                            If Right(Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4), 1) = "B" Then
                                If strDensen.Trim = "A2N" Then
                                    If strCable = "" Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei & "-A2N"
                                    ElseIf strKiriIchikbn = "1" Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei & "-A2N21"
                                    Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei & "-A2N2"
                                    End If
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei
                                End If
                            Else
                                If strDensen.Trim = "A2N" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei & "-A2N"
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-" & strKokei
                                End If
                            End If
                        End If
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Else

                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) <> "N4GE" Then

                    Select Case True
                        Case Left(strKokei, 2) = "CL" Or _
                             Left(strKokei, 2) = "CD" Or _
                             Left(strKokei, 2) = "CF" Or _
                             Left(strKokei, 3) = "C18"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If InStr(objKtbnStrc.strcSelection.strSeriesKataban.Trim, "3G") <> 0 And _
                               (InStr(strKiriIchikbn, "1") <> 0 Or _
                                InStr(strKiriIchikbn, "11") <> 0) Then
                                If InStr(strKokei, "N") <> 0 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               Left(strKokei, InStr(strKokei, "N") - 1) & _
                                                                               CdCst.Sign.Hypen & "S"
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               strKokei & CdCst.Sign.Hypen & "S"
                                End If
                            Else
                                If InStr(strKokei, "N") <> 0 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & Left(strKokei, InStr(strKokei, "N") - 1)
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & strKokei
                                End If
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                End If
            End If
            '電線接続加算価格キー
            If strDensen <> "" Then
                If strDensen <> "A2N" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               strDensen
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                End If
            End If

            '端子・ｺﾈｸﾀﾋﾟﾝ配列
            If strTanshi <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strTanshi
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション　加算価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "F"
                        Select Case strKiriIchikbn
                            Case "66", "67", "76", "77"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "DUAL"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "S", "E"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        If strKiriIchikbn <> "1" And strKiriIchikbn <> "11" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "H"
                        'If Not objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
                        If Not objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" And Not objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then 'RM1610013
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Then
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then 'RM1610013
                If Not strOption.Contains("H") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "R-H"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '電圧
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                'Case "R"
                Case "R", "S" 'RM1610013
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select

            '食品製造工程向け商品 RM1610013
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                If strOptionFP1.Contains("FP1") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
