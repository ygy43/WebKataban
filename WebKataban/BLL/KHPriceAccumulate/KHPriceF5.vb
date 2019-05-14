'************************************************************************************
'*  ProgramID  ：KHPriceF5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ハイブリロボ　３アクション空圧ロボット　ＨＲＬ－３Ａ
'*
'************************************************************************************
'Module KHPriceF5

'    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
'                                   ByRef strOpRefKataban() As String, _
'                                   ByRef decOpAmount() As Decimal)

'        Dim intRZStroke As Integer

'        Try

'            '配列定義
'            ReDim strOpRefKataban(0)
'            ReDim decOpAmount(0)

'            '中間STまるめ処理
'            Select Case True
'                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 149
'                    intRZStroke = 75
'                Case 150 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And _
'                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 349
'                    intRZStroke = 150
'                Case 350 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
'                    intRZStroke = 350
'            End Select

'            '基本価格キー
'            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-3A-" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & intRZStroke.ToString
'            decOpAmount(UBound(decOpAmount)) = 1

'            'オプション加算価格キー(スイッチ)
'            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
'                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-3A-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
'                decOpAmount(UBound(decOpAmount)) = 1

'                'オプション加算価格キー(レール)
'                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-3A-RAIL"
'                decOpAmount(UBound(decOpAmount)) = 1

'                'オプション加算価格キー(リード線長さ)
'                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
'                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                    strOpRefKataban(UBound(strOpRefKataban)) = "HRL-3A-" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
'                    decOpAmount(UBound(decOpAmount)) = 1
'                End If
'            End If

'            'オプション加算価格キー(落下防止機構)
'            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
'                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-3A-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
'                decOpAmount(UBound(decOpAmount)) = 1
'            End If

'        Catch ex As Exception

'            Throw ex

'        End Try

'    End Sub

'End Module
