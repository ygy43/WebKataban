Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KatabanDAL
    ''' <summary>
    ''' ＥＬ品判定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strElFlg">ELフラグ（1:EL品/0:中国圧力容器輸出不可商品）</param>
    ''' <returns></returns>
    ''' <remarks>引当てた形番がＥＬ品かどうかチェックする</remarks>
    Public Function fncSelectELKataban(objCon As SqlConnection, ByVal strKataban As String, ByVal strElFlg As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban ")
            sbSql.Append(" FROM    kh_el_kataban_mst ")
            sbSql.Append(" WHERE   @Kataban Like kataban ")
            sbSql.Append(" AND   el_flg = @ElFlg ")            '2013/08/27 ADD

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@ElFlg", SqlDbType.VarChar, 1).Value = strElFlg   '2013/08/27 ADD
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
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
    Public Function fncSelectQtyUnitInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                   ByVal strLanguageCd As String, ByRef strQtyUnitNm As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  c.qty_unit_nm, ")
            sbSql.Append("         b.qty_unit_nm as default_unit_nm,")
            sbSql.Append("         d.sales_unit,")
            sbSql.Append("         d.sap_base_unit,")
            sbSql.Append("         d.quantity_per_sales_unit,")
            sbSql.Append("         d.order_lot")
            sbSql.Append(" FROM    kh_qty_unit a")
            sbSql.Append(" INNER JOIN  kh_qty_unit_nm_mst b")
            sbSql.Append(" ON      a.qty_unit_cd           = b.qty_unit_cd ")
            sbSql.Append(" AND     b.language_cd           = @DefaultLangCd ")
            sbSql.Append(" AND     b.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     b.out_effective_date    > @StandardDate ")
            sbSql.Append(" LEFT  JOIN  kh_qty_unit_nm_mst c")
            sbSql.Append(" ON      a.qty_unit_cd           = c.qty_unit_cd ")
            sbSql.Append(" AND     c.language_cd           = @LanguageCd ")
            sbSql.Append(" AND     c.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     c.out_effective_date    > @StandardDate ")
            sbSql.Append(" LEFT JOIN kh_qty_unit_mst AS d on a.qty_unit_cd = d.qty_unit_cd")
            sbSql.Append(" WHERE   a.kataban               = @Kataban ")
            sbSql.Append(" AND     a.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date    > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@LanguageCd", SqlDbType.Char, 3).Value = strLanguageCd
                .Parameters.Add("@DefaultLangCd", SqlDbType.Char, 3).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
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
    Public Function fncSelectStockInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                 ByVal strLanguageCd As String, ByVal strStockPlaceCd As String, _
                                 ByRef intStockQty As Integer, ByRef intShipmentQty As Integer, _
                                 ByRef strStockContent As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kh_stock.stock_place_cd, ")
            sbSql.Append("         kh_stock.stock_qty, ")
            sbSql.Append("         kh_stock.shipment_qty, ")
            sbSql.Append("         kh_stock_content.stock_content ")
            sbSql.Append(" FROM    kh_stock ")
            sbSql.Append(" INNER JOIN  kh_stock_content ")
            sbSql.Append(" ON      kh_stock.stock_cd                    = kh_stock_content.stock_cd ")
            sbSql.Append(" WHERE   kh_stock.kataban                     = @Kataban ")
            sbSql.Append(" AND     kh_stock_content.language_cd         = @LanguageCd ")
            sbSql.Append(" AND     kh_stock.stock_place_cd              = @PlaceCd ")
            sbSql.Append(" AND     kh_stock.in_effective_date          <= @StandardDate ")
            sbSql.Append(" AND     kh_stock.out_effective_date          > @StandardDate ")
            sbSql.Append(" AND     kh_stock_content.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     kh_stock_content.out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@LanguageCd", SqlDbType.Char, 3).Value = strLanguageCd
                .Parameters.Add("@PlaceCd", SqlDbType.Char, 4).Value = strStockPlaceCd
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' モード番号の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectModeNo(ByVal objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  model_no ")
            sbSql.Append(" FROM    kh_sel_spec ")
            sbSql.Append(" WHERE   user_id             = @UserId ")
            sbSql.Append(" AND     session_id          = @SessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' 選択したマニホールドの情報を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectManifoldInfo(ByVal objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String)
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  spec_strc_seq_no, ")
            sbSql.Append("         option_kataban, ")
            sbSql.Append("         quantity, ")
            sbSql.Append("         attribute_symbol, ")
            sbSql.Append("         position_info ")
            sbSql.Append(" FROM    kh_sel_spec_strc ")
            sbSql.Append(" WHERE   user_id             = @UserId ")
            sbSql.Append(" AND     session_id          = @SessionId ")
            sbSql.Append(" ORDER BY spec_strc_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' 電圧情報取得
    ''' </summary>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd">営業所コード</param>
    ''' <returns></returns>
    ''' <remarks>電圧情報テーブルを読み込み電圧情報を取得する</remarks>
    Public Shared Function fncSelectVoltageInfo(objCon As SqlConnection, ByVal strPortSize As String, _
                                                ByVal strCoil As String, ByVal strSeriesKataban As String, _
                                                ByVal strKeyKataban As String, ByVal strVoltageDiv As String, _
                                                Optional ByRef strCountryCd As String = Nothing, _
                                                Optional ByRef strOfficeCd As String = Nothing) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  b.std_voltage, ")
            sbSql.Append("         b.std_voltage_flag ")
            sbSql.Append(" FROM    kh_voltage  a ")
            sbSql.Append(" INNER JOIN  kh_std_voltage_mst  b ")
            sbSql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sbSql.Append(" AND     a.port_size           = b.port_size ")
            sbSql.Append(" AND     a.coil                = b.coil ")
            sbSql.Append(" AND     a.voltage_div         = b.voltage_div ")
            sbSql.Append(" WHERE   a.series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban         = @KeyKataban ")
            If Not strPortSize Is Nothing Then
                sbSql.Append(" AND     a.port_size           = @PortSize ")
            End If
            If Not strCoil Is Nothing Then
                sbSql.Append(" AND     a.coil                = @Coil ")
            End If
            sbSql.Append(" AND     a.voltage_div         = @VoltageDiv ")
            sbSql.Append(" AND     a.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     b.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     b.out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                If Not strPortSize Is Nothing Then
                    .Parameters.Add("@PortSize", SqlDbType.VarChar, 4).Value = strPortSize
                End If
                If Not strCoil Is Nothing Then
                    .Parameters.Add("@Coil", SqlDbType.VarChar, 2).Value = strCoil
                End If
                If strVoltageDiv = CdCst.PowerSupply.Div1 Then
                    .Parameters.Add("@VoltageDiv", SqlDbType.VarChar, 1).Value = CdCst.PowerSupply.AC
                Else
                    .Parameters.Add("@VoltageDiv", SqlDbType.VarChar, 1).Value = CdCst.PowerSupply.DC
                End If
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' ストローク取得
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intBoreSize">口径</param>
    ''' <param name="intStroke">ストローク</param>
    ''' <returns></returns>
    ''' <remarks>ストロークのサイズを調整する</remarks>
    Public Shared Function fncSelectStrokeSize(objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                     ByVal intBoreSize As Integer, _
                                     ByVal intStroke As Integer) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  b.std_stroke ")
            sbSql.Append(" FROM    kh_stroke  a ")
            sbSql.Append(" INNER JOIN  kh_std_stroke_mst  b ")
            sbSql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sbSql.Append(" AND     a.bore_size           = b.bore_size ")
            sbSql.Append(" WHERE   a.series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     a.bore_size           = @BoreSize ")
            sbSql.Append(" AND     a.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     b.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     b.out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     a.country_cd  = @countrycd ")
            sbSql.Append(" ORDER BY  b.std_stroke DESC ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = objKtbnStrc.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = objKtbnStrc.strcSelection.strKeyKataban
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = intBoreSize
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("@countrycd", SqlDbType.VarChar, 3).Value = objKtbnStrc.strcSelection.strMadeCountry
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult

    End Function

    ''' <summary>
    ''' セレクト品検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectCatalogInfo(objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  Kataban, ")
            sbSql.Append("         DispKosu, ")
            sbSql.Append("         DispNoki, ")
            sbSql.Append("         Kosu, ")
            sbSql.Append("         Noki, ")
            sbSql.Append("         MsgKbn ")
            sbSql.Append(" FROM    kh_select_catalog ")
            sbSql.Append(" WHERE   Kataban =@Kataban")
            sbSql.Append(" ORDER BY  Noki,Kataban")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 100).Value = strKataban
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' セレクト品検索(M4GB用)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectCatalogInfo4G(objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  Kataban ")
            sbSql.Append(" FROM    kh_select_M4G ")
            sbSql.Append(" WHERE   @Kataban like Kataban")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 100).Value = strKataban
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult

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
    Public Function fncELKatabanCheck_Kaigai(objCon As SqlConnection, ByVal strCountryCd As String, _
                                                    ByVal strLanguage As String, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  country_cd, kataban, kosu, kosu_cd, B.disp_nm as kosu_nm, nouki, nouki_cd, C.disp_nm as nouki_nm ")
            sbSql.Append(" FROM    kh_select_catalog_fc_mst A ")
            sbSql.Append(" LEFT JOIN  kh_select_catalog_fc_nm_mst B ")
            sbSql.Append(" ON  A.kosu_cd = B.disp_cd ")
            sbSql.Append(" AND  B.language_cd = @language_cd ")
            sbSql.Append(" LEFT JOIN  kh_select_catalog_fc_nm_mst C ")
            sbSql.Append(" ON  A.nouki_cd = C.disp_cd ")
            sbSql.Append(" AND  C.language_cd = @language_cd ")
            sbSql.Append(" WHERE   A.country_cd      = @country_cd ")
            sbSql.Append(" AND     A.kataban         = @kataban ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@country_cd", SqlDbType.VarChar, 3).Value = strCountryCd
                .Parameters.Add("@language_cd", SqlDbType.VarChar, 2).Value = strLanguage
                .Parameters.Add("@kataban", SqlDbType.VarChar, 30).Value = strKataban
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' インドユーザー特殊メッセージ
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries">シリーズ</param>
    ''' <param name="strCountryCd">ユーザー国コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSpecialUserMessage(objCon As SqlConnection, ByVal strSeries As String, ByVal strCountryCd As String, ByVal strLabelKind As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  country_cd, item_search_key, register_person, register_datetime, current_person, current_datetime ")
            sbSql.Append(" FROM    kh_select_limitation_mst ")
            sbSql.Append(" WHERE   country_cd              = @country_cd ")
            sbSql.Append(" AND     label_kind              = @label_kind ")
            sbSql.Append(" AND     item_search_key         = @item_search_key ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@country_cd", SqlDbType.VarChar, 3).Value = strCountryCd
                .Parameters.Add("@label_kind", SqlDbType.VarChar, 3).Value = strLabelKind
                .Parameters.Add("@item_search_key", SqlDbType.VarChar, 30).Value = strSeries
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function
End Class
