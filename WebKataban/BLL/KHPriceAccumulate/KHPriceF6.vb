'************************************************************************************
'*  ProgramID  ：KHPriceF6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ハイブリロボ　３アクション空圧ロボット　ＨＲ－３Ｂ
'*
'************************************************************************************
'Module KHPriceF6

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
'                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 199
'                    intRZStroke = 75
'                Case 200 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
'                            CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 499
'                    intRZStroke = 200
'                Case 500 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
'                    intRZStroke = 500
'            End Select

'            '基本価格キー
'            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'            strOpRefKataban(UBound(strOpRefKataban)) = "HR-3B-" & intRZStroke.ToString
'            decOpAmount(UBound(decOpAmount)) = 1

'            'オプション加算価格キー(バルブ有無)
'            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
'                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                strOpRefKataban(UBound(strOpRefKataban)) = "HR-3B-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
'                decOpAmount(UBound(decOpAmount)) = 1
'            End If

'            'オプション加算価格キー(スイッチ)
'            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
'                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                strOpRefKataban(UBound(strOpRefKataban)) = "HR-3B-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
'                decOpAmount(UBound(decOpAmount)) = 1

'                'オプション加算価格キー(リード線長さ)
'                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
'                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
'                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
'                    strOpRefKataban(UBound(strOpRefKataban)) = "HR-3B-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
'                    decOpAmount(UBound(decOpAmount)) = 1
'                End If
'            End If

'        Catch ex As Exception

'            Throw ex

'        End Try

'    End Sub

'End Module
