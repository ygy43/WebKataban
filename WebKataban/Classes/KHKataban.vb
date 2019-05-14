Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHKataban
    Private dalKataban As New KatabanDAL

    ''' <summary>
    ''' ＥＬ品判定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strElFlg">ELフラグ（1:EL品/0:中国圧力容器輸出不可商品）</param>
    ''' <returns></returns>
    ''' <remarks>引当てた形番がＥＬ品かどうかチェックする</remarks>
    Public Function fncELKatabanCheck(objCon As SqlConnection, ByVal strKataban As String, ByVal strElFlg As String) As Boolean
        Dim dt As New DataTable
        fncELKatabanCheck = False

        Try
            dt = dalKataban.fncSelectELKataban(objCon, strKataban, strElFlg)

            If dt.Rows.Count > 0 Then
                fncELKatabanCheck = True
            Else
                fncELKatabanCheck = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 販売数量単位情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strQtyUnitNm"> 販売数量単位</param>
    ''' <returns></returns>
    ''' <remarks>形番より販売数量単位を取得する</remarks>
    Public Function fncQtyUnitInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                   ByVal strLanguageCd As String, ByRef strQtyUnitNm As String, _
                                   ByRef objKtbnStrc As KHKtbnStrc) As Boolean
        Dim dt As New DataTable
        fncQtyUnitInfo = False

        Try
            dt = dalKataban.fncSelectQtyUnitInfo(objCon, strKataban, strLanguageCd, strQtyUnitNm)
            If dt.Rows.Count > 0 Then
                strQtyUnitNm = IIf(IsDBNull(dt.Rows(0)("qty_unit_nm")), dt.Rows(0)("default_unit_nm"), dt.Rows(0)("qty_unit_nm"))
                objKtbnStrc.strcSelection.strSalesUnit = dt.Rows(0)("sales_unit")
                objKtbnStrc.strcSelection.strSapBaseUnit = dt.Rows(0)("sap_base_unit")
                objKtbnStrc.strcSelection.strQuantityPerSalesUnit = IIf(IsDBNull(dt.Rows(0)("quantity_per_sales_unit")), 0, dt.Rows(0)("quantity_per_sales_unit"))
                objKtbnStrc.strcSelection.strOrderLot = IIf(IsDBNull(dt.Rows(0)("order_lot")), 0, dt.Rows(0)("order_lot"))
                fncQtyUnitInfo = True
            Else
                fncQtyUnitInfo = False
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 在庫情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strStockPlaceCd">在庫場所コード</param>
    ''' <param name="intStockQty">基準在庫数</param>
    ''' <param name="intShipmentQty">出荷可能数</param>
    ''' <param name="strStockContent">在庫内容</param>
    ''' <returns></returns>
    ''' <remarks>形番より在庫情報を取得する</remarks>
    Public Function fncStockInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                 ByVal strLanguageCd As String, ByVal strStockPlaceCd As String, _
                                 ByRef intStockQty As Integer, ByRef intShipmentQty As Integer, _
                                 ByRef strStockContent As String) As Boolean
        Dim dt As New DataTable
        fncStockInfo = False

        Try
            dt = dalKataban.fncSelectStockInfo(objCon, strKataban, strLanguageCd, strStockPlaceCd, intStockQty, intShipmentQty, strStockContent)

            If dt.Rows.Count > 0 Then
                intStockQty = dt.Rows(0)("stock_qty")
                intShipmentQty = dt.Rows(0)("shipment_qty")
                strStockContent = dt.Rows(0)("stock_content")

                fncStockInfo = True
            Else
                fncStockInfo = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 電圧情報取得
    ''' </summary>
    ''' <param name="objKtbnStrc">引当情報</param>
    ''' <param name="strVoltage">電圧</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd">営業所コード</param>
    ''' <returns></returns>
    ''' <remarks>電圧情報テーブルを読み込み電圧情報を取得する</remarks>
    Public Shared Function fncVoltageInfoGet(ByVal objKtbnStrc As KHKtbnStrc, ByVal strVoltage As String, _
                                      Optional ByRef strCountryCd As String = Nothing, _
                                      Optional ByRef strOfficeCd As String = Nothing) As String
        Dim objOption As New KHOptionCtl
        Dim objCon As New SqlClient.SqlConnection(My.Settings.connkhdb)

        Dim intVoltage As Integer
        Dim strVoltageDiv As String = Nothing
        Dim strSeriesKataban As String = Nothing
        Dim strKeyKataban As String = Nothing
        Dim strPortSize As String = Nothing
        Dim strCoil As String = Nothing
        Dim dt As New DataTable

        '標準電圧にデフォルト設定
        fncVoltageInfoGet = CdCst.VoltageDiv.Other

        Try
            '電圧検索情報取得
            Call objOption.subVoltageSearchInfoGet(objKtbnStrc, strVoltage, intVoltage, strVoltageDiv, _
                                                   strSeriesKataban, strKeyKataban, strPortSize, strCoil)

            If objKtbnStrc.strcSelection.dt_vol Is Nothing OrElse objKtbnStrc.strcSelection.dt_vol.Rows.Count <= 0 Then
                'DBオープン
                objCon.Open()
                '電圧の取得
                dt = KatabanDAL.fncSelectVoltageInfo(objCon, strPortSize, strCoil, strSeriesKataban, strKeyKataban, strVoltageDiv, strCountryCd, strOfficeCd)

                For Each dr In dt.Rows
                    'マッチする電圧が存在した場合は電圧区分を設定する
                    If intVoltage = dr("std_voltage") Then
                        If dr("std_voltage_flag").ToString.Trim <> "" Then
                            fncVoltageInfoGet = dr("std_voltage_flag").ToString
                            Exit For
                        End If
                    End If
                Next
            Else
                Dim strWhere As String = String.Empty
                strWhere = "series_kataban='" & strSeriesKataban & "' AND key_kataban='" & strKeyKataban & "' "
                If Not strPortSize Is Nothing Then strWhere &= " AND port_size='" & strPortSize & "' "
                If Not strCoil Is Nothing Then strWhere &= " AND coil='" & strCoil & "' "
                If strVoltageDiv = CdCst.PowerSupply.Div1 Then
                    strWhere &= " AND voltage_div='" & CdCst.PowerSupply.AC & "' "
                Else
                    strWhere &= " AND voltage_div='" & CdCst.PowerSupply.DC & "' "
                End If
                Dim dr() As DataRow = objKtbnStrc.strcSelection.dt_vol.Select(strWhere)
                For inti As Integer = 0 To dr.Length - 1
                    'マッチする電圧が存在した場合は電圧区分を設定する
                    If intVoltage = dr(inti)("std_voltage") Then
                        If dr(inti)("std_voltage_flag").ToString.Trim <> "" Then
                            fncVoltageInfoGet = dr(inti)("std_voltage_flag").ToString
                            Exit For
                        End If
                    End If
                Next
            End If

            If fncVoltageInfoGet <> CdCst.VoltageDiv.Standard AndAlso strCountryCd IsNot Nothing Then
                If Not KHKataban.fncVoltageIsStandard(strVoltage, strCountryCd, strOfficeCd) Then
                    '標準電圧
                    fncVoltageInfoGet = CdCst.VoltageDiv.Standard
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objCon Is Nothing Then If Not objCon.State = ConnectionState.Closed Then objCon.Close()
            objCon = Nothing
        End Try

    End Function

    ''' <summary>
    ''' ストローク調整
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intBoreSize">口径</param>
    ''' <param name="intStroke">ストローク</param>
    ''' <returns></returns>
    ''' <remarks>ストロークのサイズを調整する</remarks>
    Public Shared Function fncGetStrokeSize(ByVal objKtbnStrc As KHKtbnStrc, _
                                     ByVal intBoreSize As Integer, _
                                     ByVal intStroke As Integer) As Integer
        Dim objCon As New SqlClient.SqlConnection(My.Settings.connkhdb)
        Dim dt As New DataTable
        fncGetStrokeSize = intStroke

        Try
            If objKtbnStrc.strcSelection.dt_Stroke Is Nothing OrElse objKtbnStrc.strcSelection.dt_Stroke.Rows.Count <= 0 Then
                'DBオープン
                objCon.Open()

                dt = KatabanDAL.fncSelectStrokeSize(objCon, objKtbnStrc, intBoreSize, intStroke)

                For Each dr In dt.Rows
                    '入力ストローク以上のサイズに設定する
                    If intStroke <= dr("std_stroke") Then
                        fncGetStrokeSize = dr("std_stroke")
                    Else
                        Exit For
                    End If
                Next
            Else
                Dim strWhere As String = String.Empty
                strWhere = "series_kataban='" & objKtbnStrc.strcSelection.strSeriesKataban & "' "
                strWhere &= " AND key_kataban='" & objKtbnStrc.strcSelection.strKeyKataban & "' "
                strWhere &= " AND bore_size='" & intBoreSize & "' "
                strWhere &= " AND country_cd='" & objKtbnStrc.strcSelection.strMadeCountry & "' "

                Dim dr() As DataRow = objKtbnStrc.strcSelection.dt_Stroke.Select(strWhere)
                For inti As Integer = 0 To dr.Length - 1
                    '入力ストローク以上のサイズに設定する
                    If intStroke <= dr(inti)("std_stroke") Then
                        fncGetStrokeSize = dr(inti)("std_stroke")
                    Else
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objCon Is Nothing Then If Not objCon.State = ConnectionState.Closed Then objCon.Close()
            objCon = Nothing
        End Try

    End Function

    ''' <summary>
    '''  利用機能情報取得
    ''' </summary>
    ''' <param name="strUserId">ユーザＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strSelectLang">選択言語</param>
    ''' <param name="intUseFncInfoLvl">利用機能情報レベル</param>
    ''' <param name="strUseFncInfo">利用機能情報</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks>
    ''' ・基幹Ｉ／Ｆ(1)
    ''' ・販社管理システムＩ／Ｆ(2)
    ''' ・掛率マスタメンテナンス(3)
    ''' ・ユーザーマスタメンテナンス(4)
    ''' ・情報マスタメンテナンス(5)
    ''' </remarks>
    Public Shared Sub subUseFncInfoGet(ByVal strUserId As String, _
                                 ByVal strSessionId As String, _
                                 ByVal strSelectLang As String, _
                                 ByVal intUseFncInfoLvl As Integer, _
                                 ByRef strUseFncInfo() As String, objKtbnStrc As KHKtbnStrc)
        Dim intUseInfoLvl As Integer
        Dim strUseFncDiv() As String
        Dim intLoopCnt As Integer
        Dim strQtyUnitNm As String = Nothing

        Try
            '配列初期化
            ReDim strUseFncDiv(5)
            ReDim strUseFncInfo(5)

            For intLoopCnt = 1 To 5
                strUseFncDiv(intLoopCnt) = False
            Next

            '表示付加情報レベル設定
            intUseInfoLvl = intUseFncInfoLvl

            While intUseInfoLvl > 0
                '表示付加情報レベル計算
                If intUseInfoLvl >= 16 Then
                    '情報マスタメンテナンス
                    intUseInfoLvl = intUseInfoLvl - 16
                    strUseFncDiv(1) = "True"
                ElseIf intUseInfoLvl >= 8 Then
                    'ユーザーマスタメンテナンス
                    intUseInfoLvl = intUseInfoLvl - 8
                    strUseFncDiv(2) = "True"
                ElseIf intUseInfoLvl >= 4 Then
                    '掛率マスタメンテナンス
                    intUseInfoLvl = intUseInfoLvl - 4
                    strUseFncDiv(3) = "True"
                ElseIf intUseInfoLvl >= 2 Then
                    '販社管理システムＩ／Ｆ
                    intUseInfoLvl = intUseInfoLvl - 2
                    strUseFncDiv(4) = "True"
                ElseIf intUseInfoLvl >= 1 Then
                    '基幹Ｉ／Ｆ
                    intUseInfoLvl = intUseInfoLvl - 1
                    strUseFncDiv(5) = "True"
                End If
            End While

            '配列を逆にする(価格区分順にソートする)
            Array.Reverse(strUseFncDiv)

            '配列初期化
            ReDim strUseFncInfo(UBound(strUseFncDiv))
            For intLoopCnt = 1 To strUseFncDiv.Length - 1
                strUseFncInfo(intLoopCnt) = strUseFncDiv(intLoopCnt - 1)
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' MAXソレノイド値取得
    ''' </summary>
    ''' <param name="strElectKataban">電装形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetMaxSol_01(ByVal strElectKataban As String) As Integer
        fncGetMaxSol_01 = 0
        Select Case strElectKataban
            Case "N4E0-TM1C"
                fncGetMaxSol_01 = 5
            Case "N4E0-T52", "N4E0-TM1B", "N4E0-TM52", "N4E0-T6A0", "N4E0-T6C0", "N4E0-T6E0", "N4E0-T6J0", "N4E0-T52R"
                fncGetMaxSol_01 = 8
            Case "N4E0-TM1A"
                fncGetMaxSol_01 = 10
            Case "N4E0-T50", "N4E0-T631", "N4E0-T6A1", "N4E0-T6C1", "N4E0-T6E1", _
                 "N4E0-T6J1", "N4E0-T6K1", "N4E0-T6G1", "N4E0-T7G1", "N4E0-T50R", _
                 "N4E0-T7D1", "N4E0-T7N1", "N4E0-T7EC1", "N4E0-T7ECT1" '2016/08/19 RM1608024 T7EC Append
                fncGetMaxSol_01 = 16
            Case "N4E0-T51", "N4E0-T51R", "N3Q0-T51", "N3Q0-T51U", "N3Q0-T51R", "N3Q0-T51UR"
                fncGetMaxSol_01 = 18
            Case "N4E0-T30", "N4E0-T53", "N4E0-T5B", "N4E0-T30R", "N4E0-T53R", _
                 "N3Q0-T30", "N3Q0-T53", "N3Q0-T53U", "N3Q0-T30U", "N4E0-T30N", "N4E0-T30NR", _
                 "N3Q0-T30R", "N3Q0-T53R", "N3Q0-T53UR", "N3Q0-T30UR"
                fncGetMaxSol_01 = 24
            Case "N4E0-T5C"     '28点に変更
                fncGetMaxSol_01 = 28
            Case "N4E0-T7G2", "N4E0-T7D2", "N4E0-T7N2", "N4E0-T7EC2", "N4E0-T7ECT2" '2016/08/19 RM1608024 T7EC Append
                fncGetMaxSol_01 = 32
        End Select

    End Function

    ''' <summary>
    ''' ソレノイドMAX設定
    ''' </summary>
    ''' <param name="strValue">要素画面で選択したデータの集合</param>
    ''' <param name="intManifold">Manihold種類</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetMaxSol(strValue() As String, ByVal intManifold As Integer) As UInteger
        'ソレノイドMAX設定
        fncGetMaxSol = 0
        Select Case intManifold
            Case 3
                If Left(strValue(9), 1) = "T" Then
                    Select Case Left(strValue(9), 2)
                        Case "T3"
                            If Left(strValue(9), 3) = "T30" Then
                                fncGetMaxSol = 24
                            End If
                        Case "T5"
                            If Left(strValue(9), 3) = "T50" Then
                                fncGetMaxSol = 16
                            End If
                        Case Else
                            If Left(strValue(9), 2) = "T6" And _
                               strValue(9).Length = 4 Then
                                If Right(strValue(9), 1) = "1" Then
                                    fncGetMaxSol = 16
                                Else
                                    fncGetMaxSol = 8
                                End If
                            Else
                                fncGetMaxSol = 16
                            End If
                    End Select
                End If
            Case 4

                'New4G対応 2017/01/16
                Dim strDensen As String = ""        '電線／省配線接続

                Select Case strValue(3).Trim
                    Case "R"
                        strDensen = strValue(5).Trim               '電線接続
                    Case Else
                        strDensen = strValue(4).Trim               '電線接続
                End Select
                'New4G対応 End

                'New4G対応のため、変数strDensenで判断するように変更  2016/01/16 変更 松原
                If Left(strDensen.ToString.Trim, 1) = "T" Then
                    Select Case Left(strDensen.ToString.Trim, 2)
                        Case "T1"
                            Select Case Left(strDensen.ToString.Trim, 3)
                                Case "T10"
                                    'MAXソノレイド連数を14→16に修正 2017/01/17  
                                    fncGetMaxSol = 16
                                Case "T11"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T3"
                            Select Case Left(strDensen.ToString.Trim, 3)
                                Case "T30"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T5"
                            Select Case Left(strDensen.ToString.Trim, 3)
                                Case "T50"
                                    fncGetMaxSol = 16
                                Case "T51"
                                    fncGetMaxSol = 18
                                Case "T52"
                                    fncGetMaxSol = 8
                                Case "T53"
                                    fncGetMaxSol = 24
                            End Select

                            'New4G対応 T8※の場合のパターンを追加 2017/01/16 追加 松原
                        Case "T8"
                            Dim Len As Integer = strDensen.Length
                            Select Case Mid(strDensen.ToString.Trim, Len, 1)
                                Case "1"
                                    fncGetMaxSol = 16
                                Case "2"
                                    fncGetMaxSol = 32
                            End Select

                        Case Else
                            If Left(strDensen.ToString.Trim, 2) = "T6" Then
                                If Len(strDensen.ToString.Trim) = 4 Then
                                    If Mid(strDensen.ToString.Trim, 4, 1) = "1" Then
                                        fncGetMaxSol = 16
                                    Else
                                        fncGetMaxSol = 8
                                    End If
                                Else
                                    fncGetMaxSol = 16
                                End If
                            Else
                                fncGetMaxSol = 16
                            End If
                    End Select
                End If
            Case 7
                If strValue(3) = "R" Then
                    If Left(strValue(5), 1) = "T" Then
                        Select Case Left(strValue(5), 2)
                            Case "T1"
                                Select Case Left(strValue(5), 3)
                                    Case "T10"
                                        fncGetMaxSol = 16
                                    Case "T11"
                                        fncGetMaxSol = 24
                                End Select
                            Case "T3"
                                If Left(strValue(5), 3) = "T30" Then
                                    fncGetMaxSol = 24
                                End If
                            Case "T5"
                                Select Case Left(strValue(5), 3)
                                    Case "T50"
                                        fncGetMaxSol = 16
                                    Case "T51"
                                        fncGetMaxSol = 18
                                    Case "T52"
                                        fncGetMaxSol = 8
                                    Case "T53"
                                        fncGetMaxSol = 24
                                End Select
                            Case "T6"
                                If strValue(5).Length = 4 Then
                                    If strValue(5).Substring(3, 1) = "1" Then
                                        fncGetMaxSol = 16
                                    Else
                                        fncGetMaxSol = 8
                                    End If
                                Else
                                    fncGetMaxSol = 16
                                End If
                            Case "T7"
                                If strValue(5).Length >= 4 Then
                                    If strValue(5).Substring(3, 1) = "1" Then
                                        fncGetMaxSol = 16
                                    Else
                                        If strValue(5).ToString.Length >= 5 Then
                                            If strValue(5).Substring(4, 1) = "1" Then
                                                fncGetMaxSol = 16 '電線接続"T7SP1"を選択時
                                            Else
                                                fncGetMaxSol = 8
                                            End If
                                        Else
                                            fncGetMaxSol = 8
                                        End If
                                    End If
                                Else
                                    fncGetMaxSol = 8
                                End If
                            Case "T8"
                                Select Case Right(strValue(5), 1)
                                    Case "2"
                                        fncGetMaxSol = 32
                                    Case "1"
                                        fncGetMaxSol = 16
                                End Select
                            Case Else
                                fncGetMaxSol = 16
                        End Select
                    End If
                Else
                    If Left(strValue(4), 1) = "T" Then
                        Select Case Left(strValue(4), 2)
                            Case "T1"
                                Select Case Left(strValue(4), 3)
                                    Case "T10"
                                        fncGetMaxSol = 14
                                    Case "T11"
                                        fncGetMaxSol = 24
                                End Select
                            Case "T3"
                                If Left(strValue(4), 3) = "T30" Then
                                    fncGetMaxSol = 24
                                End If
                            Case "T5"
                                Select Case Left(strValue(4), 3)
                                    Case "T50"
                                        fncGetMaxSol = 16
                                    Case "T51"
                                        fncGetMaxSol = 18
                                    Case "T52"
                                        fncGetMaxSol = 8
                                    Case "T53"
                                        fncGetMaxSol = 24
                                End Select
                            Case "T6"
                                If strValue(4).Length = 4 Then
                                    If strValue(4).Substring(3, 1) = "1" Then
                                        fncGetMaxSol = 16
                                    Else
                                        fncGetMaxSol = 8
                                    End If
                                Else
                                    fncGetMaxSol = 16
                                End If
                            Case "T7"
                                If strValue(4).Length >= 4 Then
                                    If strValue(4).Substring(3, 1) = "1" Then
                                        fncGetMaxSol = 16
                                    Else
                                        If strValue(4).ToString.Length >= 5 Then
                                            If strValue(4).Substring(4, 1) = "1" Then
                                                fncGetMaxSol = 16 '電線接続"T7SP1"を選択時
                                            Else
                                                fncGetMaxSol = 8
                                            End If
                                        Else
                                            fncGetMaxSol = 8
                                        End If
                                    End If
                                Else
                                    fncGetMaxSol = 8
                                End If
                            Case "T8"
                                Select Case Right(strValue(4), 1)
                                    Case "2"
                                        fncGetMaxSol = 32
                                    Case "1"
                                        fncGetMaxSol = 16
                                End Select
                            Case Else
                                fncGetMaxSol = 16
                        End Select
                    End If
                End If
            Case 9
                Select Case strValue(6).ToString
                    Case "T30", "T31"
                        fncGetMaxSol = 20
                    Case Else
                        If strValue(6).Substring(0, 2) = "T6" Then
                            If strValue(6).Length > 3 Then
                                If strValue(6).Substring(3, 1) = "1" Then
                                    fncGetMaxSol = 16
                                Else
                                    fncGetMaxSol = 8
                                End If
                            Else
                                fncGetMaxSol = 16
                            End If
                        Else
                            If strValue(6) = "T10" Then
                                fncGetMaxSol = 19
                            Else
                                fncGetMaxSol = 16
                            End If
                        End If
                End Select
            Case 10
                'ｿﾚﾉｲﾄﾞMAX
                Select Case strValue(0).ToString
                    Case "MN3S0", "MN4S0", "MT3S0", "MT4S0"
                        Select Case strValue(6).ToString
                            Case "T10", "T10R"
                                fncGetMaxSol = 14
                            Case "T11", "T11R", "T30", "T30R"
                                fncGetMaxSol = 24
                            Case "T50", "T50R", "T621", "T6A1", "T6C1", "T6E1"
                                fncGetMaxSol = 16
                            Case Else
                                If Left(strValue(6).ToString, 2) = "T6" Then
                                    If strValue(6).ToString.Length = 4 Then
                                        If Mid(strValue(6).ToString, 4, 1) = "1" Then
                                            fncGetMaxSol = 16
                                        Else
                                            fncGetMaxSol = 8
                                        End If
                                    Else
                                        fncGetMaxSol = 16
                                    End If
                                Else
                                    fncGetMaxSol = 16
                                End If
                        End Select
                End Select
            Case 13
                Dim strSolenoidVal As String = String.Empty
                If strValue(0).ToString = "MN4TB1" Or strValue(0).ToString = "MN4TB2" Then
                    strSolenoidVal = strValue(6).ToString                            'ソレノイドMAX比較値
                Else
                    strSolenoidVal = strValue(4).ToString                            'ソレノイドMAX比較値
                End If
                'ソレノイドMAX
                If strSolenoidVal = "T30" Or strSolenoidVal = "T31" Then
                    fncGetMaxSol = 20
                Else
                    If Left(strSolenoidVal, 2) = "T6" Then
                        If strSolenoidVal.Length = 4 Then
                            If strSolenoidVal.Substring(3, 1) = "1" Then
                                fncGetMaxSol = 16
                            Else
                                fncGetMaxSol = 8
                            End If
                        Else
                            fncGetMaxSol = 16
                        End If
                    ElseIf strSolenoidVal = "T10" Then
                        fncGetMaxSol = 19
                    Else
                        fncGetMaxSol = 16
                    End If
                End If
            Case 14
                'ソレノイドMAX
                Select Case strValue(4)
                    Case "T11R", "T30R"
                        fncGetMaxSol = 8
                    Case "T9DAR", "T9GAR"
                        fncGetMaxSol = 12
                    Case "T9L0R", "T9L8R", "T9LXR"
                        fncGetMaxSol = 24
                    Case Else
                        fncGetMaxSol = 0
                End Select
            Case 15
                Dim strOptionY As String = Nothing
                Dim str() As String = strValue(6).ToString.Split(",")
                For inti As Integer = 0 To str.Length - 1
                    If str(inti).Contains("Y") Then
                        strOptionY = str(inti)
                    End If
                Next
                'ソレノイドＭＡＸ値設定
                Select Case strValue(4).ToString
                    Case "R1"
                        If strValue(1).ToString = "1" Then
                            fncGetMaxSol = 16
                        Else
                            fncGetMaxSol = 32
                        End If
                    Case "T10", "T51"
                        fncGetMaxSol = 18
                    Case "T20"
                        fncGetMaxSol = 16
                    Case "T30", "T53"
                        fncGetMaxSol = 24

                        'シリアル伝送要素追加により変更 2016/12/09 変更 松原
                    Case "T8G1", "T8D1", "T8C1", "T7ECP1", "T7EC1", "T7ENP1",
                         "T7EN1", "T8DP1", "T7EB1", "T7EP1", "T7EBP1", "T7EPP1",
                         "T7D1", "T7DP1"
                        'RM1609016 "T7ENP1" 追加
                        'RM1612013 "T8DP1" 追加
                        'RM1612014 "T7EN1" 追加
                        'RM1708015 "T7EB1", "T7EBP1", "T7EP1", "T7EPP1" 追加
                        'RM1709014 "T7D1", "T7DP1" 追加
                        If strOptionY Is Nothing Then
                            fncGetMaxSol = 16
                        ElseIf strOptionY = "Y01" Then
                            fncGetMaxSol = 12
                        ElseIf strOptionY = "Y02" Then
                            fncGetMaxSol = 8
                        End If
                    Case "T8G2", "T8D2", "T7ECP2", "T7EC2",
                         "T7EN2", "T7ENP2", "T8DP2", "T7EB2", "T7EBP2", "T7EP2", "T7EPP2",
                         "T7D2", "T7DP2"
                        'RM1612013 "T8DP2" 追加
                        'RM1612014 "T7EN2", "T7ENP2" 追加
                        'RM1708015 "T7EB2", "T7EBP2" 追加
                        'RM1709014 "T7D2", "T7DP2" 追加
                        fncGetMaxSol = 32
                    Case "T8G7", "T8D7", "T7ECB7", "T7ECPB7",
                         "T7ENB7", "T7ENPB7", "T7EBB7", "T7EBPB7", "T7EPB7", "T7EPPB7",
                         "T7DB7", "T7DPB7"
                        'RM1612014 "T7ENB7", "T7ENPB7" 追加
                        'RM1708015 "T7EBB7", "T7EBPB7", "T7EPB7", "T7EPPB7" 追加
                        'RM1709014 "T7DB7", "T7DPB7" 追加
                        If strOptionY Is Nothing Then
                            fncGetMaxSol = 16
                        Else
                            Select Case strOptionY
                                Case "Y10", "Y20", "Y30", "Y40"
                                    fncGetMaxSol = 16
                                Case "Y11", "Y21", "Y31", "Y41"
                                    fncGetMaxSol = 12
                                Case "Y12", "Y22", "Y32", "Y42"
                                    fncGetMaxSol = 8
                            End Select
                        End If
                    Case "T8M6", "T8C6"
                        If strOptionY Is Nothing Then
                            fncGetMaxSol = 8
                        Else
                            Select Case strOptionY
                                Case "Y10", "Y20"
                                    fncGetMaxSol = 8
                                Case "Y01", "Y11", "Y21"
                                    fncGetMaxSol = 4
                            End Select
                        End If
                    Case "T8MA"
                        fncGetMaxSol = 4
                End Select
            Case 18
                Dim strDen As String = String.Empty
                If strValue(3).ToString.Trim = "R" Then
                    strDen = strValue(5)
                Else
                    strDen = strValue(4)
                End If
                If Strings.Left(strDen.ToString, 1) = "T" Then
                    Select Case Strings.Left(strDen, 2)
                        Case "T1"
                            Select Case Strings.Left(strDen, 3)
                                Case "T10"
                                    fncGetMaxSol = 16
                                Case "T11"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T3"
                            If Strings.Left(strDen, 3) = "T30" Then
                                fncGetMaxSol = 24
                            End If
                        Case "T5"
                            Select Case Strings.Left(strDen, 3)
                                Case "T50"
                                    fncGetMaxSol = 16
                                Case "T51"
                                    fncGetMaxSol = 18
                                Case "T52"
                                    fncGetMaxSol = 8
                                Case "T53"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T6"
                            If strDen.Length = 4 Then
                                If strDen.Substring(3, 1) = "1" Then
                                    fncGetMaxSol = 16
                                Else
                                    fncGetMaxSol = 8
                                End If
                            Else
                                fncGetMaxSol = 16
                            End If
                        Case "T7"
                            If strDen.Length >= 4 Then
                                If strDen.Substring(3, 1) = "1" Then
                                    fncGetMaxSol = 16
                                Else
                                    If strDen.Length >= 5 Then
                                        If strDen.Substring(4, 1) = "1" Then
                                            fncGetMaxSol = 16 '電線接続"T7SP1"を選択時
                                        Else
                                            fncGetMaxSol = 8
                                        End If
                                    Else
                                        fncGetMaxSol = 8
                                    End If
                                End If
                            Else
                                fncGetMaxSol = 8
                            End If
                        Case "T8"
                            Select Case Right(strDen, 1)
                                Case "2"
                                    fncGetMaxSol = 32
                                Case "1"
                                    fncGetMaxSol = 16
                            End Select
                        Case Else
                            fncGetMaxSol = 16
                    End Select
                End If

            Case 19
                Dim strOptionT As String
                Select Case strValue(0)
                    Case "MN3EX0", "MN4EX0"
                        strOptionT = strValue(4)
                    Case "MN3Q0", "MT3Q0"
                        strOptionT = strValue(5)
                    Case Else
                        strOptionT = strValue(6)
                End Select

                If Left(strOptionT, 1) = "T" Then
                    Select Case Left(strOptionT, 2)
                        Case "T1"
                            Select Case Left(strOptionT, 3)
                                Case "T10"
                                    fncGetMaxSol = 16
                                Case "T11"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T3"
                            Select Case Left(strOptionT, 3)
                                Case "T30"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T5"
                            Select Case Left(strOptionT, 3)
                                Case "T50"
                                    fncGetMaxSol = 16
                                Case "T51"
                                    fncGetMaxSol = 18
                                Case "T52"
                                    fncGetMaxSol = 8
                                Case "T53"
                                    fncGetMaxSol = 24
                            End Select
                        Case "T8"
                            If strOptionT.EndsWith("1") Then
                                fncGetMaxSol = 16
                            Else
                                fncGetMaxSol = 32
                            End If
                        Case "T7"
                            If strOptionT.EndsWith("1") Then
                                fncGetMaxSol = 16
                            Else
                                fncGetMaxSol = 32
                            End If
                        Case "T6"
                            If strOptionT.EndsWith("1") Then
                                fncGetMaxSol = 16
                            Else
                                fncGetMaxSol = 8
                            End If
                        Case "TM"
                            Select Case strOptionT
                                Case "TM1A"
                                    fncGetMaxSol = 10
                                Case "TM1C"
                                    fncGetMaxSol = 5
                                Case "TM52"
                                    fncGetMaxSol = 8
                            End Select

                        Case Else


                    End Select
                End If
        End Select
    End Function

    ''' <summary>
    ''' 形番から重複するハイフンを除去する
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncHypenCut(ByVal strKataban As String) As String
        Dim sbKataban As New StringBuilder(60)
        Dim bolHypenFlg As Boolean = False
        Dim intLoopCnt As Integer

        fncHypenCut = ""
        Try
            For intLoopCnt = 1 To strKataban.Length
                If Mid(strKataban, intLoopCnt, 1) = CdCst.Sign.Hypen Then
                    If bolHypenFlg = True Then
                        '1桁前がハイフンの場合は次へ
                    Else
                        '形番生成
                        sbKataban.Append(Mid(strKataban, intLoopCnt, 1))
                    End If

                    'ハイフンフラグＯＮ
                    bolHypenFlg = True
                Else
                    '形番生成
                    sbKataban.Append(Mid(strKataban, intLoopCnt, 1))

                    'ハイフンフラグＯＦＦ
                    bolHypenFlg = False
                End If
            Next

            fncHypenCut = sbKataban.ToString

            '形番の右側がハイフンの場合は除去する
            If Left(fncHypenCut, 1) = CdCst.Sign.Hypen Then
                fncHypenCut = Mid(fncHypenCut, 2, fncHypenCut.Length)
            End If
            '形番の左側がハイフンの場合は除去する
            If Right(fncHypenCut, 1) = CdCst.Sign.Hypen Then
                fncHypenCut = Left(fncHypenCut, fncHypenCut.Length - 1)
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            sbKataban = Nothing
        End Try
    End Function

    ''' <summary>
    ''' スイッチ数取得
    ''' </summary>
    ''' <param name="strSwitchQty">スイッチ数</param>
    ''' <returns></returns>
    ''' <remarks>スイッチ数を判定し返却する</remarks>
    Public Shared Function fncSwitchQtyGet(ByVal strSwitchQty As String) As Integer
        Try
            'スイッチ数判定
            Select Case Left(strSwitchQty, 1)
                Case "R", "L", "H", "X"
                    fncSwitchQtyGet = 1
                Case "D"
                    fncSwitchQtyGet = 2
                Case "T"
                    fncSwitchQtyGet = 3
                Case Else
                    fncSwitchQtyGet = CInt(strSwitchQty)
            End Select
        Catch ex As Exception
            fncSwitchQtyGet = 1
        End Try
    End Function

    ''' <summary>
    ''' 異電圧判定
    ''' </summary>
    ''' <param name="strVoltage">電圧</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd">営業所コード</param>
    ''' <returns>True:異電圧、Flase:標準電圧</returns>
    ''' <remarks>
    ''' 海外店対応
    ''' 海外ユーザーの場合、異電圧加算を行わない
    ''' </remarks>
    Public Shared Function fncVoltageIsStandard(ByVal strVoltage As String, ByVal strCountryCd As String, _
                                         ByVal strOfficeCd As String) As Boolean
        '初期値
        fncVoltageIsStandard = True
        Try
            '海外ユーザは異電圧加算しない
            If (strCountryCd <> CdCst.CountryCd.DefaultCountry) Or _
               (strCountryCd = CdCst.CountryCd.DefaultCountry And _
                strOfficeCd = CdCst.OfficeCd.Overseas) Then

                Select Case strVoltage
                    'AC110V/AC220V/AC120V/AC240V　海外では標準電圧扱い
                    Case CdCst.PowerSupply.Const5, CdCst.PowerSupply.Const6, _
                         CdCst.PowerSupply.Const7, CdCst.PowerSupply.Const8
                        '標準電圧
                        fncVoltageIsStandard = False
                End Select
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' フル形番から第一ハイフン前を取得
    ''' </summary>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncMdlKtbnGet(ByVal strKataban As String) As String
        fncMdlKtbnGet = String.Empty
        Try
            '機種形番設定
            If strKataban.IndexOf(CdCst.Sign.Hypen) > 0 Then
                fncMdlKtbnGet = Mid(strKataban, 1, strKataban.IndexOf(CdCst.Sign.Hypen))
            Else
                fncMdlKtbnGet = strKataban
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 簡易仕様書かどうかを判定する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc">形番情報取得クラス</param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns>簡易仕様書ならば「true」それ以外は「false」</returns>
    ''' <remarks>
    ''' 引当仕様書テーブルから機種番号を取得して簡易仕様書かどうかを判断する
    ''' 特定機種のミックスマニホールドは簡易仕様書でなくても仕様書№を表示させるためにtrueを戻す
    ''' </remarks>
    Public Shared Function fncJudgeSimpleSpec(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                        strUserId As String, strSessionId As String) As Boolean
        fncJudgeSimpleSpec = False

        Dim dtResult As New DataTable
        Try
            'モード番号の取得
            dtResult = KatabanDAL.fncSelectModeNo(objCon, strUserId, strSessionId)

            If dtResult.Rows.Count > 0 Then
                'モード番号が「0」の場合は簡易マニホールド
                If Len(Trim(dtResult.Rows(0)("model_no"))) = 0 Then
                    fncJudgeSimpleSpec = True
                Else
                    fncJudgeSimpleSpec = False
                End If
            Else
                fncJudgeSimpleSpec = False
            End If

            'モード番号記録されない場合は機種で判断すること
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN4KB1", "MN4KB2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).PadRight(2, " ").Substring(0, 2) = "80" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "M"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "N"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "MN3Q0", "MT3Q0"
                    fncJudgeSimpleSpec = False
                Case "M4HA1", "M4HA2", "M4JA1", "M4JA2"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                    'RM1805001_4Rシリーズ追加
                Case "M4RD1", "M4RD2", "M4RE1", "M4RE2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "MN4GDX12", "MN4GEX12"
                    fncJudgeSimpleSpec = False
                Case "B"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "M4SA0", "M4SB0"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "M3KA1", "M4KA1", "M4KA2", "M4KA3", "M4KA4", _
                     "M4KB1", "M4KB2", "M4KB3", "M4KB4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Then
                                fncJudgeSimpleSpec = True
                            End If
                        Case "M"
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                                fncJudgeSimpleSpec = True
                            End If
                    End Select
                Case "M4F0", "M4F1", "M4F2", "M4F3", "M4F4", "M4F5", "M4F6", "M4F7"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
                Case "M3MA0", "M3MB0", "M3PA1", "M3PA2", "M3PB1", "M3PB2", "M4L2", "M4LB2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncJudgeSimpleSpec = True
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 簡易仕様書のミックスマニホールド数量情報を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns>ミックスマニホールド数量情報</returns>
    ''' <remarks>引当仕様書構成テーブルから簡易仕様書のミックスマニホールド数量情報を取得</remarks>
    Public Shared Function fncGetMixManifoldInfo(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                        strUserId As String, strSessionId As String) As Object
        Dim strOptKatabanInfo() As String
        Dim intQtyInfo() As Integer
        Dim bolMPFlg As Boolean = False

        Dim intQty2 As Integer = 0
        Dim intQty3 As Integer = 0
        Dim intQty4 As Integer = 0
        Dim dt As New DataTable
        fncGetMixManifoldInfo = Nothing
        Try

            ReDim strOptKatabanInfo(0)
            ReDim intQtyInfo(0)

            If objKtbnStrc.strcSelection.strSeriesKataban = "M" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Case "4"
                        ReDim intQtyInfo(8)
                    Case "3"
                        ReDim intQtyInfo(3)
                End Select
            End If

            '選択したマニホールド情報を取得
            dt = KatabanDAL.fncSelectManifoldInfo(objCon, strUserId, strSessionId)

            If dt.Rows.Count > 0 Then
                'マニホールド画面の記号欄に記載がある分だけ繰り返す
                For Each dr In dt.Rows
                    If objKtbnStrc.strcSelection.strSeriesKataban = "N" Then
                        'ミックスマニホールド数量情報のエリア定義
                        If dr("spec_strc_seq_no") = 1 Then
                            ReDim Preserve intQtyInfo(3)
                        End If
                        '電磁弁のみ集計対象
                        If dr("spec_strc_seq_no") >= 3 AndAlso _
                        dr("spec_strc_seq_no") <= 12 Then

                            'チェック有無
                            Dim setQuantity As Integer
                            setQuantity = dr("quantity")
                            If setQuantity > 0 Then
                                '形番の４桁目にて集計
                                Select Case Mid(dr("option_kataban"), 4, 1)
                                    Case "2"
                                        intQtyInfo(1) = intQtyInfo(1) + setQuantity
                                    Case "3"
                                        intQtyInfo(2) = intQtyInfo(2) + setQuantity
                                    Case "4"
                                        intQtyInfo(3) = intQtyInfo(3) + setQuantity
                                End Select
                            End If
                        End If

                    Else
                        ReDim Preserve strOptKatabanInfo(UBound(strOptKatabanInfo) + 1)
                        strOptKatabanInfo(UBound(strOptKatabanInfo)) = _
                        IIf(IsDBNull(dr("option_kataban")), "", dr("option_kataban"))

                        If strOptKatabanInfo(UBound(strOptKatabanInfo)) = "" And (objKtbnStrc.strcSelection.strSeriesKataban = "M" Or objKtbnStrc.strcSelection.strSeriesKataban = "MN4KB1" Or objKtbnStrc.strcSelection.strSeriesKataban = "MN4KB2") Then
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "M"
                                Case "MN4KB1", "MN4KB2"
                                Case Else
                                    Exit For
                            End Select
                        Else
                            If strOptKatabanInfo(UBound(strOptKatabanInfo)) = "" Then
                                Exit For
                            End If

                        End If

                        '「M4F2」「M4F3」の特殊条件
                        'マニホールド画面の記号欄に「MP」は表示されないが数量情報にゼロを補って桁数を4桁固定にする処理
                        If (objKtbnStrc.strcSelection.strSeriesKataban = "M4F2" Or objKtbnStrc.strcSelection.strSeriesKataban = "M4F3") _
                        And InStr(1, strOptKatabanInfo(UBound(strOptKatabanInfo)), "MP") <> 0 Then
                            bolMPFlg = True
                        End If

                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "M"
                                If strOptKatabanInfo(UBound(strOptKatabanInfo)) <> "" And Not IsDBNull(dr("position_info")) Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "4"
                                            'Series_Kataban=M（M4SA1など）のとき
                                            'option_katabanがNULLでなく、position_infoがNULLでない場合のみマニホールド形番に追加する
                                            If strOptKatabanInfo(UBound(strOptKatabanInfo)).Length >= 6 Then
                                                Select Case Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 5, 2)
                                                    Case "19"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "3S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(6) = intQtyInfo(6) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "4S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(1) = intQtyInfo(1) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "11"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "3S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(7) = intQtyInfo(7) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "29"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "4S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(2) = intQtyInfo(2) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "39"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "4S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(3) = intQtyInfo(3) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "49"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "4S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(4) = intQtyInfo(4) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "59"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "4S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(5) = intQtyInfo(5) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                End Select
                                            Else
                                                'MP
                                                If strOptKatabanInfo(UBound(strOptKatabanInfo)).Trim = "MP" Then
                                                    If dr("quantity") > 0 Then
                                                        intQtyInfo(8) = intQtyInfo(8) + _
                                                        IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                    End If
                                                End If
                                            End If
                                        Case "3"
                                            If strOptKatabanInfo(UBound(strOptKatabanInfo)).Length >= 6 Then
                                                Select Case Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 5, 2)
                                                    Case "19"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "3S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(1) = intQtyInfo(1) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                    Case "11"
                                                        If Mid(strOptKatabanInfo(UBound(strOptKatabanInfo)), 1, 2) = "3S" Then
                                                            If dr("quantity") > 0 Then
                                                                intQtyInfo(2) = intQtyInfo(2) + _
                                                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                            End If
                                                        End If
                                                End Select
                                            Else
                                                'MP
                                                If strOptKatabanInfo(UBound(strOptKatabanInfo)).Trim = "MP" Then
                                                    If dr("quantity") > 0 Then
                                                        intQtyInfo(3) = intQtyInfo(3) + _
                                                        IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                                                    End If
                                                End If
                                            End If
                                    End Select
                                End If
                            Case "MN4KB1", "MN4KB2"
                                If IIf(IsDBNull(dr("attribute_symbol")), "", dr("attribute_symbol")) = "B3" And strOptKatabanInfo(UBound(strOptKatabanInfo)) <> "" Then
                                    '切替位置区分がミックス対象
                                    If Left(objKtbnStrc.strcSelection.strOpSymbol(1), 1) = "8" Then
                                        'ミックスマニホールド数量情報のエリア定義
                                        ReDim Preserve intQtyInfo(5)

                                        '電磁弁のみ集計対象
                                        If dr("spec_strc_seq_no") >= 9 AndAlso _
                                        dr("spec_strc_seq_no") <= 14 Then
                                            'チェック有無
                                            Dim setQuantity As Integer
                                            setQuantity = dr("quantity")
                                            If setQuantity > 0 Then
                                                Select Case Mid(dr("option_kataban"), 6, 1)
                                                    Case "1"
                                                        intQtyInfo(1) = intQtyInfo(1) + setQuantity
                                                    Case "2"
                                                        intQtyInfo(2) = intQtyInfo(2) + setQuantity
                                                    Case "3"
                                                        intQtyInfo(3) = intQtyInfo(3) + setQuantity
                                                    Case "4"
                                                        intQtyInfo(4) = intQtyInfo(4) + setQuantity
                                                    Case "5"
                                                        intQtyInfo(5) = intQtyInfo(5) + setQuantity
                                                End Select
                                            End If
                                        End If
                                    End If
                                End If
                            Case Else
                                ReDim Preserve intQtyInfo(UBound(intQtyInfo) + 1)
                                intQtyInfo(UBound(intQtyInfo)) = _
                                IIf(IsDBNull(dr("quantity")), "", dr("quantity"))
                        End Select
                    End If
                Next
            End If

            If (objKtbnStrc.strcSelection.strSeriesKataban = "M4F2" Or objKtbnStrc.strcSelection.strSeriesKataban = "M4F3") _
            And bolMPFlg = False Then
                ReDim Preserve intQtyInfo(UBound(intQtyInfo))
                intQtyInfo(UBound(intQtyInfo)) = 0
            End If

            '「B*P512(3,4)」の特殊条件
            'マニホールド画面の記号欄に「M4」は表示されないが数量情報の3桁目にゼロを補って桁数を4桁固定にする処理
            If objKtbnStrc.strcSelection.strSeriesKataban = "B" _
            AndAlso InStr(1, strOptKatabanInfo(3), "P514") = 0 Then
                ReDim Preserve intQtyInfo(4)
                intQtyInfo(4) = intQtyInfo(3)
                intQtyInfo(3) = 0
            End If

            fncGetMixManifoldInfo = intQtyInfo
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 特定の形番の時、適用個数を""にする。
    ''' </summary>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks>FRL中国生産対応で適用個数の仕様が不確定のための処置</remarks>
    Public Shared Function subJapanChinaAmount(ByVal strKataban As String) As Boolean

        ''Web系形引のまま流用する、VB6はCCIテーブルの内容
        'If strKataban = "F1000-6-W" Or _
        'strKataban = "F1000-8-W" Or _
        'strKataban = "M1000-6-W" Or _
        'strKataban = "M1000-8-W" Or _
        'strKataban = "R1000-6-W" Or _
        'strKataban = "R1000-8-W" Or _
        'strKataban = "L1000-6-W" Or _
        'strKataban = "L1000-8-W" Or _
        'strKataban = "W1000-6-W" Or _
        'strKataban = "W1000-8-W" Or _
        'strKataban = "C1000-6-W" Or _
        'strKataban = "C1000-8-W" Or _
        'strKataban = "F3000-8-W" Or _
        'strKataban = "F3000-8-W-F" Or _
        'strKataban = "F3000-10-W" Or _
        'strKataban = "F3000-10-W-F" Or _
        'strKataban = "M3000-8-W" Or _
        'strKataban = "M3000-8-W-F1" Or _
        'strKataban = "M3000-10-W" Or _
        'strKataban = "M3000-10-W-F1" Or _
        'strKataban = "R3000-8-W" Or _
        'strKataban = "R3000-10-W" Or _
        'strKataban = "L3000-8-W" Or _
        'strKataban = "L3000-10-W" Or _
        'strKataban = "W3000-8-W" Then
        '    Return True
        'End If

        'If strKataban = "W3000-8-W-F" Or _
        'strKataban = "W3000-10-W" Or _
        'strKataban = "W3000-10-W-F" Or _
        'strKataban = "C3000-8-W" Or _
        'strKataban = "C3000-8-W-F" Or _
        'strKataban = "C3000-10-W" Or _
        'strKataban = "C3000-10-W-F" Or _
        'strKataban = "F4000-10-W" Or _
        'strKataban = "F4000-10-W-F" Or _
        'strKataban = "F4000-15-W" Or _
        'strKataban = "F4000-15-W-F" Or _
        'strKataban = "M4000-10-W" Or _
        'strKataban = "M4000-10-W-F1" Or _
        'strKataban = "M4000-15-W" Or _
        'strKataban = "M4000-15-W-F1" Or _
        'strKataban = "R4000-10-W" Or _
        'strKataban = "R4000-15-W" Or _
        'strKataban = "L4000-10-W" Or _
        'strKataban = "L4000-15-W" Or _
        'strKataban = "W4000-10-W" Or _
        'strKataban = "W4000-10-W-F" Or _
        'strKataban = "W4000-15-W" Or _
        'strKataban = "W4000-15-W-F" Or _
        'strKataban = "C4000-10-W" Or _
        'strKataban = "C4000-10-W-F" Then
        '    Return True
        'End If

        'If strKataban = "C4000-15-W" Or _
        'strKataban = "C4000-15-W-F" Or _
        'strKataban = "F8000-20-W" Or _
        'strKataban = "F8000-20-W-F" Or _
        'strKataban = "F8000-25-W" Or _
        'strKataban = "F8000-25-W-F" Or _
        'strKataban = "M8000-20-W" Or _
        'strKataban = "M8000-20-W-F1" Or _
        'strKataban = "M8000-25-W" Or _
        'strKataban = "M8000-25-W-F1" Or _
        'strKataban = "R8000-20-W" Or _
        'strKataban = "R8000-25-W" Or _
        'strKataban = "L8000-20-W" Or _
        'strKataban = "L8000-25-W" Or _
        'strKataban = "W8000-20-W" Or _
        'strKataban = "W8000-20-W-F" Or _
        'strKataban = "W8000-25-W" Or _
        'strKataban = "W8000-25-W-F" Or _
        'strKataban = "C8000-20-W" Or _
        'strKataban = "C8000-20-W-F" Or _
        'strKataban = "C8000-25-W" Or _
        'strKataban = "C8000-25-W-F" Then
        '    Return True
        'End If

        ''2008.7.30 追加分
        'If strKataban = "F1000-6-W-BW" Or _
        'strKataban = "F1000-8-W-BW" Or _
        'strKataban = "M1000-6-W-BW" Or _
        'strKataban = "M1000-8-W-BW" Or _
        'strKataban = "R1000-6-W-BW" Or _
        'strKataban = "R1000-6-W-B3W" Or _
        'strKataban = "R1000-8-W-BW" Or _
        'strKataban = "R1000-8-W-B3W" Or _
        'strKataban = "L1000-6-W-BW" Or _
        'strKataban = "L1000-8-W-BW" Or _
        'strKataban = "W1000-6-W-BW" Or _
        'strKataban = "W1000-6-W-B3W" Or _
        'strKataban = "W1000-8-W-BW" Or _
        'strKataban = "W1000-8-W-B3W" Or _
        'strKataban = "F3000-8-W-BW" Or _
        'strKataban = "F3000-8-W-F-BW" Or _
        'strKataban = "F3000-10-W-BW" Or _
        'strKataban = "F3000-10-W-F-BW" Or _
        'strKataban = "M3000-8-W-BW" Or _
        'strKataban = "M3000-8-W-F1-BW" Or _
        'strKataban = "M3000-10-W-BW" Or _
        'strKataban = "M3000-10-W-F1-BW" Or _
        'strKataban = "R3000-8-W-BW" Or _
        'strKataban = "R3000-8-W-B3W" Or _
        'strKataban = "R3000-10-W-BW" Then
        '    Return True
        'End If

        'If strKataban = "R3000-10-W-B3W" Or _
        'strKataban = "L3000-8-W-BW" Or _
        'strKataban = "L3000-10-W-BW" Or _
        'strKataban = "W3000-8-W-BW" Or _
        'strKataban = "W3000-8-W-B3W" Or _
        'strKataban = "W3000-8-W-F-BW" Or _
        'strKataban = "W3000-10-W-BW" Or _
        'strKataban = "W3000-10-W-B3W" Or _
        'strKataban = "W3000-10-W-F-BW" Or _
        'strKataban = "W3000-10-W-F-B3W" Or _
        'strKataban = "F4000-10-W-BW" Or _
        'strKataban = "F4000-10-W-F-BW" Or _
        'strKataban = "F4000-15-W-BW" Or _
        'strKataban = "F4000-15-W-F-BW" Or _
        'strKataban = "M4000-10-W-BW" Or _
        'strKataban = "M4000-10-W-F1-BW" Or _
        'strKataban = "M4000-15-W-BW" Or _
        'strKataban = "M4000-15-W-F1-BW" Or _
        'strKataban = "R4000-10-W-BW" Or _
        'strKataban = "R4000-10-W-B3W" Or _
        'strKataban = "R4000-15-W-BW" Or _
        'strKataban = "R4000-15-W-B3W" Or _
        'strKataban = "L4000-10-W-BW" Or _
        'strKataban = "L4000-15-W-BW" Or _
        'strKataban = "W4000-10-W-BW" Then
        '    Return True
        'End If

        'If strKataban = "W4000-10-W-B3W" Or _
        'strKataban = "W4000-10-W-F-BW" Or _
        'strKataban = "W4000-10-W-F-B3W" Or _
        'strKataban = "W4000-15-W-BW" Or _
        'strKataban = "W4000-15-W-B3W" Or _
        'strKataban = "W4000-15-W-F-BW" Or _
        'strKataban = "W4000-15-W-F-B3W" Or _
        'strKataban = "F8000-20-W-BW" Or _
        'strKataban = "F8000-20-W-F-BW" Or _
        'strKataban = "F8000-25-W-BW" Or _
        'strKataban = "F8000-25-W-F-BW" Or _
        'strKataban = "M8000-20-W-BW" Or _
        'strKataban = "M8000-20-W-F1-BW" Or _
        'strKataban = "M8000-25-W-BW" Or _
        'strKataban = "M8000-25-W-F1-BW" Or _
        'strKataban = "R8000-20-W-BW" Or _
        'strKataban = "R8000-25-W-BW" Or _
        'strKataban = "L8000-20-W-BW" Or _
        'strKataban = "L8000-25-W-BW" Or _
        'strKataban = "W8000-20-W-BW" Or _
        'strKataban = "W8000-20-W-F-BW" Or _
        'strKataban = "W8000-25-W-BW" Or _
        'strKataban = "W8000-25-W-F-BW" Then
        '    Return True
        'End If

        If strKataban = "AMDZ13R-6UP-Z0N4H" Or _
            strKataban = "AMDZ13R-8BUP-Z0N4H" Or _
            strKataban = "AMD313R-10UP-00N4H" Or _
            strKataban = "AMD313R-10BUP-00N4H" Or _
            strKataban = "AMD313R-12UP-00N4H" Or _
            strKataban = "AMD313R-15BUP-00N4H" Or _
            strKataban = "AMD413R-20BUP-00N4H" Or _
            strKataban = "AMD513R-25UP-00N4H" Or _
            strKataban = "AMD513R-25BUP-00N4H" Or _
            strKataban = "MMD302-10BUP-8-U" Or _
            strKataban = "MMD302-10BUP-8-P" Or _
            strKataban = "MMD302-15BUP-10-U" Or _
            strKataban = "MMD302-15BUP-10-P" Or _
            strKataban = "MMD402-20BUP-16-U" Or _
            strKataban = "MMD402-20BUP-16-P" Or _
            strKataban = "MMD502-25BUP-20-U" Or _
            strKataban = "MMD502-25BUP-20-P" Or _
            strKataban = "AGD11R-4RM" Or _
            strKataban = "AGD11R-4R" Or _
            strKataban = "AGD11R-4S" Or _
            strKataban = "AGD21R-6RM" Or _
            strKataban = "AGD21R-6R" Or _
            strKataban = "AGD21R-6S" Or _
            strKataban = "PGM-30-4R" Or _
            strKataban = "PGM-30-4RM" Or _
            strKataban = "PGM-H-60-4R" Or _
            strKataban = "PGM-H-60-4RM" Or _
            strKataban = "AVB217-16K-4" Or _
            strKataban = "AVB317-25K-4" Or _
            strKataban = "AVB417-40K-4" Or _
            strKataban = "AVB517-50K-4" Or _
            strKataban = "MVB217-16K" Or _
            strKataban = "MVB317-25K" Or _
            strKataban = "MVB417-40K" Or _
            strKataban = "MVB517-50K" Or _
            strKataban = "VG-05F" Or _
            strKataban = "VG-05P" Or _
            strKataban = "AMD313R-10BUP-00N4F" Or _
            strKataban = "AMD313R-15BUP-00N4F" Or _
            strKataban = "AMD413R-20BUP-00N4F" Or _
            strKataban = "AMD513R-25BUP-00N4F" Or _
            strKataban = "MMD302-3BT-A" Or _
            strKataban = "MMD302-4BT-A" Or _
            strKataban = "MMD402-6BT-A" Or _
            strKataban = "MMD502-8BT-A" Then
            Return True
        End If

        Return False

    End Function

    ''' <summary>
    ''' セレクト品検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="htSelInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectCatalogInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                          ByRef htSelInfo As Hashtable) As Boolean
        Dim dt As New DataTable
        Dim dalKatabanTmp As New KatabanDAL

        fncSelectCatalogInfo = False

        Try
            htSelInfo = New Hashtable
            dt = dalKatabanTmp.fncSelectCatalogInfo(objCon, strKataban)
            If dt.Rows.Count > 0 Then
                htSelInfo("DispKosu") = ClsCommon.fncIsInputed(dt.Rows(0)("DispKosu").ToString, "")
                htSelInfo("DispNoki") = ClsCommon.fncIsInputed(dt.Rows(0)("DispNoki").ToString, "")
                htSelInfo("Kosu") = ClsCommon.fncIsInputed(dt.Rows(0)("Kosu").ToString, "")
                htSelInfo("Noki") = ClsCommon.fncIsInputed(dt.Rows(0)("Noki").ToString, "")
                htSelInfo("MsgKbn") = ClsCommon.fncIsInputed(dt.Rows(0)("MsgKbn").ToString, "")
                fncSelectCatalogInfo = True
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' セレクト品検索(M4GB用)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKatas">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectCatalogInfo4G(objCon As SqlConnection, ByVal strKatas() As String, ByVal intCounts() As Double) As Boolean
        'CHANGED BY YGY 20141023
        Dim strKatabans As System.Collections.Generic.List(Of String) = CdCst.strSelectM4GCheckData

        For inti As Integer = 1 To 18
            If Not strKatas(inti).Equals(String.Empty) And intCounts(inti) <> 0 Then
                If Not strKatabans.Contains(strKatas(inti)) Then
                    Return False
                End If
            End If
        Next
        fncSelectCatalogInfo4G = True
        'Dim dt As New DataTable
        'Dim dalKatabanTmp As New KatabanDAL

        'fncSelectCatalogInfo4G = False

        'Try
        '    dt = dalKatabanTmp.fncSelectCatalogInfo4G(objCon, strKataban)

        '    If dt.Rows.Count > 0 Then
        '        fncSelectCatalogInfo4G = True
        '        'SelectM4G = True
        '    End If
        'Catch ex As Exception
        '    WriteErrorLog("E001", ex)
        'End Try
    End Function

    ''' <summary>
    ''' 中国セレクト品の検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strLanguage"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncELKatabanCheck_Kaigai(objCon As SqlConnection, ByVal strCountryCd As String, _
                                                    ByVal strLanguage As String, ByVal strKataban As String) As DataTable
        Dim dtResult As New DataTable
        Dim dalKatabanTmp As New KatabanDAL

        Try
            dtResult = dalKatabanTmp.fncELKatabanCheck_Kaigai(objCon, strCountryCd, strLanguage, strKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' ユーザー特殊メッセージ登録かどうか
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strCountryCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSpecialUserMessage(objCon As SqlConnection, ByVal strSeries As String, ByVal strCountryCd As String, ByVal strLabelKind As String) As Boolean

        Dim blnResult As Boolean = False
        Dim dt As New DataTable
        Dim dalKatabanTmp As New KatabanDAL

        dt = dalKatabanTmp.fncSpecialUserMessage(objCon, strSeries, strCountryCd, strLabelKind)

        If dt.Rows.Count > 0 Then
            blnResult = True
        End If

        Return blnResult

    End Function

End Class
