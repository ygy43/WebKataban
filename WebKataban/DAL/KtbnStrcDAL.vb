Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KtbnStrcDAL
    ''' <summary>
    ''' 引当積上単価構成追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intDispSeqNo">表示順序</param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所コード</param>
    ''' <param name="intListPrice">定価</param>
    ''' <param name="intRegPrice">登録店価格</param>
    ''' <param name="intSsprice">SS店価格</param>
    ''' <param name="intBsprice">BS店価格</param>
    ''' <param name="intGsprice">GS店価格</param>
    ''' <param name="intPsprice">PS店価格</param>
    ''' <param name="decAmount">数量</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <remarks>引当積上単価構成テーブルに追加する</remarks>
    Public Sub subAccPriceStrcMnt(objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String, _
                                  ByVal intDispSeqNo As Integer, ByVal strKataban As String, _
                                  ByVal strKatabanCheckDiv As String, ByVal strPlaceCd As String, _
                                  ByVal intListPrice As Decimal, ByVal intRegPrice As Decimal, _
                                  ByVal intSsprice As Decimal, ByVal intBsprice As Decimal, _
                                  ByVal intGsprice As Decimal, ByVal intPsprice As Decimal, _
                                  ByVal decAmount As Decimal, _
                                  ByVal strCurrency As String, ByVal strMadeCountry As String)
        Dim objCmd As New SqlCommand
        Try

            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelAccPrcStrcIns, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@DispSeqNo", SqlDbType.Int).Value = intDispSeqNo
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@KatabanCheckDiv", SqlDbType.Char, 1).Value = IIf(IsNothing(strKatabanCheckDiv), "", strKatabanCheckDiv)
                .Parameters.Add("@PlaceCd", SqlDbType.VarChar, 4).Value = IIf(IsNothing(strPlaceCd), "", strPlaceCd)
                .Parameters.Add("@ListPrice", SqlDbType.Money).Value = intListPrice
                .Parameters.Add("@RegPrice", SqlDbType.Money).Value = intRegPrice
                .Parameters.Add("@Ssprice", SqlDbType.Money).Value = intSsprice
                .Parameters.Add("@Bsprice", SqlDbType.Money).Value = intBsprice
                .Parameters.Add("@Gsprice", SqlDbType.Money).Value = intGsprice
                .Parameters.Add("@Psprice", SqlDbType.Money).Value = intPsprice
                .Parameters.Add("@Amount", SqlDbType.Decimal).Value = decAmount
                .Parameters.Add("@CurrencyCd", SqlDbType.VarChar, 3).Value = strCurrency    'Add by Zxjike 2013/05/16
                .Parameters.Add("@CountryCd", SqlDbType.VarChar, 3).Value = strMadeCountry  'Add by Zxjike 2013/06/07
            End With

            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番更新処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strRodEndOption">ロッド先端特注仕様</param>
    ''' <param name="strOtherOption">オプション外仕様</param>
    ''' <param name="strPositionOption">簡易仕様書設置位置仕様</param>
    ''' <remarks>引当シリーズ形番テーブルのオプション情報を更新する</remarks>
    Public Sub subSelSrsKtbnOptionUpd(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                      ByVal strSessionId As String, _
                                      Optional ByVal strRodEndOption As String = "", _
                                      Optional ByVal strOtherOption As String = "", _
                                      Optional ByVal strPositionOption As String = "")
        Dim objCmd As New SqlCommand
        Try
            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSrsKtbnOptionUpd, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.Char, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@RodEndOption", SqlDbType.VarChar, 60).Value = strRodEndOption
                .Parameters.Add("@OtherOption", SqlDbType.VarChar, 60).Value = strOtherOption
                .Parameters.Add("@PositionOption", SqlDbType.VarChar, 60).Value = strPositionOption
            End With
            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当形番構成追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strElementDiv">要素区分</param>
    ''' <param name="strStructureDiv">構成区分</param>
    ''' <param name="strAdditionDiv">付加区分</param>
    ''' <param name="strHyphenDiv">継続ハイフン有無区分</param>
    ''' <param name="strKtbnStrcNm">形番構成名称</param>
    ''' <param name="strplace_lvl"></param>
    ''' <remarks>引当形番構成テーブルにデータを追加する</remarks>
    Public Sub subSelKtbnStrcIns(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                 ByVal strSessionId As String, ByVal intKtbnStrcSeqNo As Integer, _
                                 ByVal strElementDiv As String, ByVal strStructureDiv As String, _
                                 ByVal strAdditionDiv As String, ByVal strHyphenDiv As String, _
                                 ByVal strKtbnStrcNm As String, ByVal strplace_lvl As Integer)
        Dim objCmd As New SqlCommand

        Try
            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelKtbnStrcIns, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure

                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@KtbnStrcSeqNo", SqlDbType.Int).Value = intKtbnStrcSeqNo
                .Parameters.Add("@ElementDiv", SqlDbType.VarChar, 150).Value = strElementDiv
                .Parameters.Add("@StructureDiv", SqlDbType.Char, 1).Value = strStructureDiv
                .Parameters.Add("@AdditionDiv", SqlDbType.Char, 1).Value = strAdditionDiv
                .Parameters.Add("@HyphenDiv", SqlDbType.Char, 1).Value = strHyphenDiv
                .Parameters.Add("@KtbnStrcNm", SqlDbType.NVarChar, 150).Value = strKtbnStrcNm
                .Parameters.Add("@place_lvl", SqlDbType.Int).Value = strplace_lvl
            End With
            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try

    End Sub

    ''' <summary>
    ''' 引当シリーズ形番更新処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所コード</param>
    ''' <param name="strCostCalcNo">原価積算No.</param>
    ''' <param name="intListPrice">定価</param>
    ''' <param name="intRegPrice">登録店価格</param>
    ''' <param name="intSsPrice">SS店価格</param>
    ''' <param name="intBsPrice">BS店価格</param>
    ''' <param name="intGsPrice">GS店価格</param>
    ''' <param name="intPsPrice">PS店価格</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <remarks>引当シリーズ形番テーブルのデータを更新する</remarks>
    Public Sub subSelSrsKtbnPriceUpd(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String, ByVal strKatabanCheckDiv As String, _
                                     ByVal strPlaceCd As String, ByVal strCostCalcNo As String, _
                                     ByVal intListPrice As Decimal, ByVal intRegPrice As Decimal, _
                                     ByVal intSsPrice As Decimal, ByVal intBsPrice As Decimal, _
                                     ByVal intGsPrice As Decimal, ByVal intPsPrice As Decimal, _
                                     ByVal strCurrency As String, ByVal strMadeCountry As String)
        Dim objCmd As New SqlCommand

        Try
            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSrsKtbnPriUpd, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure

                ' 定義
                .Parameters.Add("@UserId", SqlDbType.Char, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@KatabanCheckDiv", SqlDbType.Char, 1).Value = strKatabanCheckDiv
                .Parameters.Add("@PlaceCd", SqlDbType.VarChar, 4).Value = strPlaceCd
                .Parameters.Add("@CostCalcNo", SqlDbType.VarChar, 7).Value = IIf(strCostCalcNo = "", DBNull.Value, strCostCalcNo)
                .Parameters.Add("@ListPrice", SqlDbType.Money).Value = intListPrice
                .Parameters.Add("@RegPrice", SqlDbType.Money).Value = intRegPrice
                .Parameters.Add("@Ssprice", SqlDbType.Money).Value = intSsPrice
                .Parameters.Add("@Bsprice", SqlDbType.Money).Value = intBsPrice
                .Parameters.Add("@Gsprice", SqlDbType.Money).Value = intGsPrice
                .Parameters.Add("@Psprice", SqlDbType.Money).Value = intPsPrice
                .Parameters.Add("@CurrencyCd", SqlDbType.VarChar, 3).Value = strCurrency    'add by Zxjike 2013/05/16
                .Parameters.Add("@CountryCd", SqlDbType.VarChar, 3).Value = strMadeCountry  'add by Zxjike 2013/06/07
            End With
            '実行
            objCmd.ExecuteNonQuery()

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当形番構成更新処理
    ''' </summary>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strStroptionSymbol">オプション記号</param>
    ''' <param name="PlaceLvl"></param>
    ''' <param name="objCon_2"></param>
    ''' <remarks></remarks>
    Public Shared Sub subSelKtbnStrcUpd(ByVal strUserId As String, _
                                 ByVal strSessionId As String, ByVal intKtbnStrcSeqNo As Integer, _
                                 ByVal strStroptionSymbol As String, ByVal PlaceLvl As Long, _
                                 Optional ByVal objCon_2 As SqlConnection = Nothing)
        Dim objCmd As SqlCommand
        Dim bolClose As Boolean = False
        Try
            If objCon_2 Is Nothing Then
                bolClose = True
                objCon_2 = New SqlConnection
                objCon_2 = New SqlClient.SqlConnection(My.Settings.connkhdb)
                objCon_2.Open()
            End If

            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelKtbnStrcUpd, objCon_2)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@KtbnStrcSeqNo", SqlDbType.Int).Value = intKtbnStrcSeqNo
                .Parameters.Add("@OptionSymbol", SqlDbType.VarChar, 150).Value = strStroptionSymbol
                .Parameters.Add("@place_lvl", SqlDbType.Int).Value = PlaceLvl
            End With
            objCmd.ExecuteNonQuery() '実行
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
            If bolClose Then
                If Not objCon_2 Is Nothing Then If Not objCon_2.State = ConnectionState.Closed Then objCon_2.Close()
                objCon_2 = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当形番構成追加処理
    ''' </summary>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <remarks></remarks>
    Public Function fncSelectAccPriceStrc(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                          ByVal strSessionId As String) As DataTable

        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban_check_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         ls_price, ")
            sbSql.Append("         rg_price, ")
            sbSql.Append("         ss_price, ")
            sbSql.Append("         bs_price, ")
            sbSql.Append("         gs_price, ")
            sbSql.Append("         ps_price, ")
            sbSql.Append("         amount, ")
            sbSql.Append("         currency_cd, ")
            sbSql.Append("         country_cd ")
            sbSql.Append(" FROM    kh_sel_acc_prc_strc ")
            sbSql.Append(" WHERE   user_id    = @UserId ")
            sbSql.Append(" AND     session_id = @SessionId ")
            sbSql.Append(" ORDER BY  disp_seq_no ")

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
    ''' 引当シリーズ形番テーブルのデータを更新する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strKataban">形番</param>
    ''' <remarks></remarks>
    Public Sub subSelSrsKtbnFullKtbnUpd(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                        ByVal strSessionId As String, ByVal strKataban As String)
        Dim objCmd As New SqlCommand
        Try
            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSrsKtbnFullKtbnUpd, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.Char, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
            End With
            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try

    End Sub

    ''' <summary>
    ''' 引当積上単価構成削除処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <remarks></remarks>
    Public Sub subAccPriceStrcDel(ByVal objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String)
        Dim objCmd As SqlCommand
        Try
            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelAccPrcStrcDel, objCon)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With
            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番を読み込み単価情報を取得し返却する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelSrsKtbn(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                  ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  division          , series_kataban    , ")
            sbSql.Append("         key_kataban       , hyphen_div        , ")
            sbSql.Append("         price_no          , spec_no           , ")
            sbSql.Append("         full_kataban      , goods_nm          , ")
            sbSql.Append("         kataban_check_div , place_cd          , ")
            sbSql.Append("         cost_calc_no      , ls_price          , ")
            sbSql.Append("         rg_price          , ss_price          , ")
            sbSql.Append("         bs_price          , gs_price          , ")
            sbSql.Append("         ps_price          , factor            , ")
            sbSql.Append("         unit_price        , amount            , ")
            sbSql.Append("         rod_end_option    , other_option      , ")
            sbSql.Append("         position_option   , currency_cd, country_cd ")
            sbSql.Append(" FROM    kh_sel_srs_ktbn ")
            sbSql.Append(" WHERE   user_id    = @UserId ")
            sbSql.Append(" AND     session_id = @SessionId ")

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
    ''' 引当ロッド先端特注取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーID</param>
    ''' <param name="strSessionId">セッションID</param>
    ''' <returns></returns>
    ''' <remarks>引当ロッド先端特注を読み込みデータがあればTrueを返却する</remarks>
    Public Function fncSelRodSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As Boolean
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        fncSelRodSelect = False
        Try
            'SQL Query生成
            sbSql.Append(" SELECT      * ")
            sbSql.Append(" FROM        kh_sel_rod_end_order ")
            sbSql.Append(" WHERE       user_id    = @UserId ")
            sbSql.Append(" AND         session_id = @SessionId ")

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

            If dtResult.Rows.Count > 0 Then
                fncSelRodSelect = True
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当ロッド先端特注取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーID</param>
    ''' <param name="strSessionId">セッションID</param>
    ''' <returns></returns>
    ''' <remarks>引当ロッド先端特注を読み込みデータがあればTrueを返却する</remarks>
    Public Function fncSelKtbnStrcSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  option_symbol, ")
            sbSql.Append("         element_div, ")
            sbSql.Append("         structure_div, ")
            sbSql.Append("         addition_div, ")
            sbSql.Append("         hyphen_div, ")
            sbSql.Append("         ktbn_strc_nm,place_lvl ")
            sbSql.Append(" FROM    kh_sel_ktbn_strc ")
            sbSql.Append(" WHERE   user_id    = @PrmUserId ")
            sbSql.Append(" AND     session_id = @PrmSessionId ")
            sbSql.Append(" ORDER BY  ktbn_strc_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@PrmUserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@PrmSessionId", SqlDbType.NVarChar, 88).Value = strSessionId
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
    ''' 引当積上単価構成取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelAccPriceStrcSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban           , kataban_check_div , ")
            sbSql.Append("         place_cd          , ls_price          , ")
            sbSql.Append("         rg_price          , ss_price          , ")
            sbSql.Append("         bs_price          , gs_price          , ")
            sbSql.Append("         ps_price          , amount            , currency_cd,country_cd ")
            sbSql.Append(" FROM    kh_sel_acc_prc_strc ")
            sbSql.Append(" WHERE   user_id    = @PrmUserId ")
            sbSql.Append(" AND     session_id = @PrmSessionId ")
            sbSql.Append(" ORDER BY  disp_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@PrmUserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@PrmSessionId", SqlDbType.NVarChar, 88).Value = strSessionId
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
    ''' 引当積上単価構成取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelSpecSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  model_no, ")
            sbSql.Append("         wiring_spec, ")
            sbSql.Append("         din_rail_length ")
            sbSql.Append(" FROM    kh_sel_spec ")
            sbSql.Append(" WHERE   user_id    = @UserId ")
            sbSql.Append(" AND     session_id = @SessionId ")

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
    ''' 引当積上単価構成取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelSpecStrcSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  attribute_symbol, ")
            sbSql.Append("         option_kataban, ")
            sbSql.Append("         cxa_kataban, ")
            sbSql.Append("         cxb_kataban, ")
            sbSql.Append("         position_info, ")
            sbSql.Append("         quantity ")
            sbSql.Append(" FROM    kh_sel_spec_strc ")
            sbSql.Append(" WHERE   user_id    = @UserId ")
            sbSql.Append(" AND     session_id = @SessionId ")
            sbSql.Append(" ORDER BY  spec_strc_seq_no ")

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
    ''' 引当ロッド先端特注WF標準寸法取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelRodWFSelect(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                     ByVal strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  normal_value ")
            sbSql.Append(" FROM    kh_sel_rod_end_order ")
            sbSql.Append(" WHERE   user_id    = @UserId ")
            sbSql.Append(" AND     session_id = @SessionId ")
            sbSql.Append(" AND     external_form = '" & CdCst.RodEndCstmOrder.FrmWF & "'")

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
End Class
