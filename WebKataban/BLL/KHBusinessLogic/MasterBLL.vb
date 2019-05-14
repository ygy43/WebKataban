Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class MasterBLL

    ''' <summary>
    ''' ユーザー情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strDate"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_UserMstList(objCon As SqlConnection, strUserID As String, _
                                              strDate As String, strLanguage As String, intStartIndex As Integer, _
                                              intEndIndex As Integer, Optional likeflg As Boolean = True) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_UserMstList = New DataTable

        Try
            fncSQL_UserMstList = dalMasterTmp.fncSQL_UserMstList(objCon, strUserID, strDate, strLanguage, intStartIndex, intEndIndex, likeflg)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' ユーザーマスタ総数の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strDate"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_UserMstCount(objCon As SqlConnection, strUserID As String, _
                                              strDate As String, strLanguage As String) As Integer
        Dim dalMasterTmp As New MasterDAL
        fncSQL_UserMstCount = 0

        Try
            fncSQL_UserMstCount = dalMasterTmp.fncSQL_UserMstCount(objCon, strUserID, strDate, strLanguage)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 国別生産品情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_CountryItemMstList(objCon As SqlConnection, _
                                                     strCountryCd As String, strKataban As String) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_CountryItemMstList = New DataTable

        Try
            fncSQL_CountryItemMstList = dalMasterTmp.fncSQL_CountryItemMstList(objCon, strCountryCd, strKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 現地定価
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_RateMstList_L(objCon As SqlConnection, _
                                                strCountryCd As String, strKataban As String) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_RateMstList_L = New DataTable

        Try
            fncSQL_RateMstList_L = dalMasterTmp.fncSQL_RateMstList_L(objCon, strCountryCd, strKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 購入価格
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strMadeCountryCd"></param>
    ''' <param name="strSaleCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_RateMstList_N(objCon As SqlConnection, strMadeCountryCd As String, _
                                                strSaleCountryCd As String, strKataban As String) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_RateMstList_N = New DataTable

        Try
            fncSQL_RateMstList_N = dalMasterTmp.fncSQL_RateMstList_N(objCon, strMadeCountryCd, strSaleCountryCd, strKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 国マスタに存在するかチェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_CountryMst(objCon As SqlConnection) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_CountryMst = New DataTable

        Try
            fncSQL_CountryMst = dalMasterTmp.fncSQL_CountryMst(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 国マスタの取得（Sort_no順）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetAllCountryMst(objCon As SqlConnection) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncGetAllCountryMst = New DataTable

        Try
            fncGetAllCountryMst = dalMasterTmp.fncGetAllCountryMst(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 営業所情報の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_OfficeMst(objCon As SqlConnection) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_OfficeMst = New DataTable

        Try
            fncSQL_OfficeMst = dalMasterTmp.fncSQL_OfficeMst(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ユーザクラスの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_UserClassMst(objCon As SqlConnection, strLanguage As String) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_UserClassMst = New DataTable

        Try
            fncSQL_UserClassMst = dalMasterTmp.fncSQL_UserClassMst(objCon, strLanguage)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ドロップダウンリスト作成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_CountryCodeList(objCon As SqlConnection, strLanguage As String) As DataTable
        Dim dalMasterTmp As New MasterDAL
        fncSQL_CountryCodeList = New DataTable

        Try
            fncSQL_CountryCodeList = dalMasterTmp.fncSQL_CountryCodeList(objCon, strLanguage)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 国コードによりベースコードの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCod"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetBaseCdByCountryCd(objCon As SqlConnection, strCountryCod As String) As String
        fncGetBaseCdByCountryCd = String.Empty

        Dim dalMasterTmp As New MasterDAL
        Dim dt As New DataTable

        Try
            dt = dalMasterTmp.fncGetBaseCdByCountryCd(objCon, strCountryCod)

            If dt.Rows.Count > 0 Then
                fncGetBaseCdByCountryCd = dt.Rows(0).Item("base_cd")
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

End Class
