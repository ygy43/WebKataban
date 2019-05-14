'************************************************************************************
'*  ProgramID  ：KHPrice05
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2006/12/25   作成者：NII K.Sudoh
'*                                      更新日：2008/11/04   更新者：T.Sato
'*
'*  概要       ：ＧＸ３１００／ＧＸ５１００の単価計算を行う
'*               ＧＫ３１００の単価計算を行う
'*
'*  変更
'*    機種追加：ＧＫ３１００(Ｄ）／５１００    RM1004012 2010/04/22 Y.Miura
'************************************************************************************
Module KHPrice05

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intMaxOption As Integer

        Try
            '要素の最大数を取得   RM1004012 2010/04/22 Y.Miura 追加
            intMaxOption = UBound(objKtbnStrc.strcSelection.strOpSymbol)

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "GX82"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "T"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                              CdCst.Sign.Hypen & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                              CdCst.Sign.Hypen & "T"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "M"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                              CdCst.Sign.Hypen & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                              CdCst.Sign.Hypen & "M"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "S"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                              CdCst.Sign.Hypen & _
                                                                              objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                              CdCst.Sign.Hypen & "S"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "GTA"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "T"
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "M"
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "M"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "M"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "S"
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "U"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "N"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "M"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "M"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "R"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(1) & "W" & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                  CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case Else
                    '価格キー設定
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "GX"
                            'RM1312084 2013/12/25
                            If objKtbnStrc.strcSelection.strKeyKataban = "A" Or _
                               objKtbnStrc.strcSelection.strKeyKataban = "B" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "GX" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "GX" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "GK"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'RM1004012 2010/04/22 Y.Miura 
                            'strOpRefKataban(UBound(strOpRefKataban)) = "GK" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                            '                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                            '                                                  CdCst.Sign.Hypen & _
                            '                                                  objKtbnStrc.strcSelection.strOpSymbol(4)
                            If intMaxOption >= 5 Then
                                strOpRefKataban(UBound(strOpRefKataban)) = "GK" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  "D-" & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(5)
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = "GK" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                  CdCst.Sign.Hypen & _
                                                                                  objKtbnStrc.strcSelection.strOpSymbol(4)
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'オプション価格
                    'RM1004012 2010/04/22 Y.Miura 
                    'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    If intMaxOption >= 5 Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "G"
                                    '価格キー設定
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "GX" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                      CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                      CdCst.Sign.Hypen & _
                                                                                      strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    '価格キー設定
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                               "D-" & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                    Else
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "G"
                                    '価格キー設定
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "GX" & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                                      CdCst.Sign.Hypen & _
                                                                                      objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                                                      CdCst.Sign.Hypen & _
                                                                                      strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    '価格キー設定
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                               CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
