Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHSeriesSearch

#Region "プロパティ"
    Private dsKataban As DataSet
    Private strHeader As String
    Private strSrsKata As String
    Private strSearchDv As String
    Private strLangCd As String
    Private intRange As Integer
    Private strMinKataban As String
    Private strCKDMinKata As String
    Private strCountryCd As String
    Private strResultTypeCd As ResultType

    Public Enum ResultType As Integer
        ''' <summary>初期値</summary>
        ''' <remarks></remarks>
        [None] = 0
        ''' <summary>正常終了</summary>
        ''' <remarks></remarks>
        Success = 1
        ''' <summary>最大件数オーバー</summary>
        ''' <remarks></remarks>
        MaxCountOver = -1
    End Enum

#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal strSearchDvValue As String, ByVal strLangCdValue As String, ByVal strSrsKataValue As String, _
                   ByVal intRangeValue As Integer, ByVal strMinKatabanValue As String, ByVal strCountryCdValue As String)
        Me.strSearchDvValue = strSearchDvValue  '検索区分を取得する
        Me.strLangCdValue = strLangCdValue
        Me.strSrsKataValue = strSrsKataValue
        Me.intRangeValue = intRangeValue
        Me.strMinKatabanValue = strMinKatabanValue
        Me.strCountryCdValue = strCountryCdValue
        Me.strResultTypeCdValue = ResultType.None
    End Sub

    ''' <summary>
    ''' 検索結果の取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property dsKatabanValue() As DataSet
        Get
            Return Me.dsKataban
        End Get
        Set(ByVal value As DataSet)
            Me.dsKataban = value
        End Set
    End Property

    ''' <summary>
    ''' ヘッダー情報の取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strHeaderValue() As String
        Get
            Return Me.strHeader
        End Get
        Set(ByVal value As String)
            Me.strHeader = value
        End Set
    End Property

    ''' <summary>
    ''' 検索条件(型番)の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strSrsKataValue() As String
        Get
            Return Me.strSrsKata
        End Get
        Set(ByVal value As String)
            Me.strSrsKata = value
        End Set
    End Property

    ''' <summary>
    ''' 検索区分の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strSearchDvValue() As String
        Get
            Return Me.strSearchDv
        End Get
        Set(ByVal value As String)
            Me.strSearchDv = value
        End Set
    End Property

    ''' <summary>
    ''' 言語区分の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strLangCdValue() As String
        Get
            Return Me.strLangCd
        End Get
        Set(ByVal value As String)
            Me.strLangCd = value
        End Set
    End Property

    ''' <summary>
    ''' 【プロパティ】strRangeVa取得行数の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property intRangeValue() As Integer
        Get
            Return Me.intRange
        End Get
        Set(ByVal value As Integer)
            Me.intRange = value
        End Set
    End Property

    ''' <summary>
    ''' 形番の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strMinKatabanValue() As String
        Get
            Return Me.strMinKataban
        End Get
        Set(ByVal value As String)
            Me.strMinKataban = value
        End Set
    End Property

    ''' <summary>
    ''' 国コードの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strCountryCdValue() As String
        Get
            Return Me.strCountryCd
        End Get
        Set(ByVal value As String)
            Me.strCountryCd = value
        End Set
    End Property

    ''' <summary>
    ''' データ取得結果
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property strResultTypeCdValue() As String
        Get
            Return Me.strResultTypeCd
        End Get
        Set(ByVal value As String)
            Me.strResultTypeCd = value
        End Set
    End Property

End Class
