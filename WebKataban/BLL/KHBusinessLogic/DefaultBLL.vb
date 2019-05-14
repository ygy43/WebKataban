Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class DefaultBLL

    Private Property dllDefault As New DefaultDAL

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        dllDefault = New DefaultDAL
    End Sub

    ''' <summary>
    ''' 一ヶ月前の臨時テーブル情報をクリア
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objConBase"></param>
    ''' <remarks></remarks>
    Public Sub subDelErrHistory(ByVal objCon As SqlConnection, ByVal objConBase As SqlConnection)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del = New DataTable

        Dim strDate As String = Now.AddMonths(-1).ToString("yyyy/MM/dd")

        Try
            'KH_LOGINテーブルのデータを削除
            dllDefault.subDeleteFromLoginByLoginDate(objConBase, strDate)

            'kh_sel_acc_prc_strcテーブルのデータを削除
            dllDefault.subDeleteFromSelAccPrcStrcByRegDate(objCon, strDate)

            'kh_sel_ktbn_strcテーブルのデータを削除
            dllDefault.subDeleteFromSelKtbnStrcByRegDate(objCon, strDate)

            'kh_sel_outofop_orderテーブルのデータを削除
            dllDefault.subDeleteFromSelOutofopOrderByRegDate(objCon, strDate)

            'kh_sel_rod_end_orderテーブルのデータを削除
            dllDefault.subDeleteFromSelRodEndOrderByRegDate(objCon, strDate)

            'kh_sel_specテーブルのデータを削除
            dllDefault.subDeleteFromSelSpecByRegDate(objCon, strDate)

            'kh_sel_spec_strcテーブルのデータを削除
            dllDefault.subDeleteFromSelSpecStrcByRegDate(objCon, strDate)

            'kh_sel_srs_ktbnテーブルのデータを削除
            dllDefault.subDeleteFromSelSrsKtbnByRegDate(objCon, strDate)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objAdp = Nothing
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' メニュー情報取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strLangCd"></param>
    ''' <param name="strAuthorityCd"></param>
    ''' <param name="strMenuId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncMenuMstSelect(ByVal objConBase As SqlConnection, ByVal strLangCd As String, _
                                            ByVal strAuthorityCd As String, _
                                            Optional ByVal strMenuId As String = Nothing) As DataTable
        fncMenuMstSelect = New DataTable

        Try
            fncMenuMstSelect = dllDefault.fncMenuMstSelect(objConBase, strLangCd, strAuthorityCd, strMenuId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 言語選択欄の内容を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectLanguageList(ByVal objConBase As SqlConnection, ByVal strLang As String) As DataTable
        fncSelectLanguageList = New DataTable

        Try
            fncSelectLanguageList = dllDefault.fncSelectLanguageList(objConBase, strLang)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

End Class
