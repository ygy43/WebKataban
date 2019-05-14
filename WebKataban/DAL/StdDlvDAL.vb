Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class StdDlvDAL
    ''' <summary>
    ''' 引当オプション外特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>引当オプション外特注テーブルからデータを削除する</remarks>
    Public Function fncStandardDateEx(ByVal objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append(" SELECT ExceptionCode, shipment_qty  ")
            sbSql.Append(" FROM    StandardDateEX ")
            sbSql.Append(" WHERE   Kataban = @Kataban")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 30).Value = strKataban
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 引当オプション外特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>引当オプション外特注テーブルからデータを削除する</remarks>
    Public Function fncSeriesTableSelect(ByVal objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append(" SELECT  Series, ")
            sbSql.Append("         Tnt_Main, ")
            sbSql.Append("         Tnt_Main_Tel, ")
            sbSql.Append("         Tnt_Sub, ")
            sbSql.Append("         Tnt_Sub_Tel ")
            sbSql.Append(" FROM    SeriesTable ")
            sbSql.Append(" WHERE   @Symbol LIKE Symbol + '%' ")
            sbSql.Append(" ORDER BY  LEN(Symbol) DESC ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Symbol", SqlDbType.VarChar, 60).Value = strKataban
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 引当オプション外特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>引当オプション外特注テーブルからデータを削除する</remarks>
    Public Function fncItemNameTableSelect(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append(" SELECT  Name1      , Hyphen1    , ")
            sbSql.Append("         Name2      , Hyphen2    , ")
            sbSql.Append("         Name3      , Hyphen3    , ")
            sbSql.Append("         Name4      , Hyphen4    , ")
            sbSql.Append("         Name5      , Hyphen5    , ")
            sbSql.Append("         Name6      , Hyphen6    , ")
            sbSql.Append("         Name7      , Hyphen7    , ")
            sbSql.Append("         Name8      , Hyphen8    , ")
            sbSql.Append("         Name9      , Hyphen9    , ")
            sbSql.Append("         Name10     , Hyphen10   , ")
            sbSql.Append("         Name11     , Hyphen11   , ")
            sbSql.Append("         Name12     , Hyphen12   , ")
            sbSql.Append("         Name13     , Hyphen13   , ")
            sbSql.Append("         Name14     , Hyphen14   , ")
            sbSql.Append("         Name15     , Hyphen15   , ")
            sbSql.Append("         Name16     , Hyphen16   , ")
            sbSql.Append("         Name17     , Hyphen17   , ")
            sbSql.Append("         Name18     , Hyphen18   , ")
            sbSql.Append("         Name19     , Hyphen19   , ")
            sbSql.Append("         Name20     , Hyphen20   , ")
            sbSql.Append("         Name21     , Hyphen21   , ")
            sbSql.Append("         Name22     , Hyphen22   , ")
            sbSql.Append("         Name23     , Hyphen23   , ")
            sbSql.Append("         Name24 ")
            sbSql.Append(" FROM    ItemNameTable ")
            sbSql.Append(" WHERE   Series = @Series ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Series", SqlDbType.VarChar, 30).Value = strSeriesKataban
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetAllSymbol(objCon As SqlConnection, ByVal strSeriesKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand
            'SQL文生成
            sbSql.Append(" SELECT  Series, ")
            sbSql.Append("         ItemNo, ")
            sbSql.Append("         Symbol, ")
            sbSql.Append("         LEN(Symbol) AS LENGTH ")
            sbSql.Append(" FROM    SeparatorTable ")
            sbSql.Append(" WHERE   Series = @Series ")
            sbSql.Append(" ORDER BY  ItemNo,Symbol")

            objCmd.CommandText = sbSql.ToString
            With objCmd
                .Parameters.Add("@Series", SqlDbType.VarChar, 30).Value = strSeriesKataban
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 適用個数取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncStandardQuantity(objCon As SqlConnection, ByVal strSeriesKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand
            'SQL文生成
            sbSql.Append(" SELECT  Quantity ")
            sbSql.Append(" FROM    StandardQuantity ")
            sbSql.Append(" WHERE   Series = @Series ")

            objCmd.CommandText = sbSql.ToString

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Series", SqlDbType.VarChar, 30).Value = strSeriesKataban
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 納期情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncExpDate(objCon As SqlConnection, ByVal strSeriesKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL文生成
            sbSql.Append(" SELECT  Symbol1  , Symbol2  , ")
            sbSql.Append("         Symbol3  , Symbol4  , ")
            sbSql.Append("         Symbol5  , Symbol6  , ")
            sbSql.Append("         Symbol7  , Symbol8  , ")
            sbSql.Append("         Symbol9  , Symbol10 , ")
            sbSql.Append("         Symbol11 , Symbol12 , ")
            sbSql.Append("         Symbol13 , Symbol14 , ")
            sbSql.Append("         Symbol15 , Symbol16 , ")
            sbSql.Append("         Symbol17 , Symbol18 , ")
            sbSql.Append("         Symbol19 , Symbol20 , ")
            sbSql.Append("         Symbol21 , Symbol22 , ")
            sbSql.Append("         Symbol23 , Symbol24 , ")
            sbSql.Append("         Quantity , Date ")
            sbSql.Append(" FROM    expDateTable ")
            sbSql.Append(" WHERE   Series = @Series ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Series", SqlDbType.VarChar, 30).Value = strSeriesKataban
            End With

            '実行
            objAdp = New SqlDataAdapter(objCmd)
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
