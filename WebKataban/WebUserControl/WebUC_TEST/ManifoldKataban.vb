Public Class ManifoldKataban
    Private _kataban As String
    Private _siyousyo As String
    Private _kataplace As String
    Private _katacheck As String
    Private _gsprice As String

    '形番
    Public Property KATABAN As String
        Get
            Return _kataban
        End Get
        Set(value As String)
            _kataban = value
        End Set
    End Property

    '仕様書
    Public Property SIYOUSYO As String
        Get
            Return _siyousyo
        End Get
        Set(value As String)
            _siyousyo = value
        End Set
    End Property

    '出荷場所
    Public Property KATAPLACE As String
        Get
            Return _kataplace
        End Get
        Set(value As String)
            _kataplace = value
        End Set
    End Property

    'ﾁｪｯｸ区分
    Public Property KATACHECK As String
        Get
            Return _katacheck
        End Get
        Set(value As String)
            _katacheck = value
        End Set
    End Property

    'GS価格
    Public Property GSPRICE As String
        Get
            Return _gsprice
        End Get
        Set(value As String)
            _gsprice = value
        End Set
    End Property

    '初期化
    Public Sub New()
        _kataban = String.Empty
        _siyousyo = String.Empty
        _kataplace = String.Empty
        _katacheck = String.Empty
        _gsprice = String.Empty
    End Sub

    '初期化
    Public Sub New(ByVal listKataban As Object)
        If listKataban.GetType.Name.Equals("kh_shiyou_testRow") Then
            _kataban = listKataban.item("KATABAN").ToString
            _siyousyo = listKataban.item("ID").ToString
            _kataplace = listKataban.item("SHIPPLACE").ToString
            _katacheck = listKataban.item("CHECKKBN").ToString
            _gsprice = listKataban.item("GSPRICE").ToString
        ElseIf listKataban.GetType.Name.Equals("ManifoldKataban") Then
            _kataban = listKataban.KATABAN
            _siyousyo = listKataban.SIYOUSYO
            _kataplace = listKataban.KATAPLACE
            _katacheck = listKataban.KATACHECK
            _gsprice = listKataban.GSPRICE
        End If
    End Sub

    Public Overrides Function ToString() As String
        'Return "形番：" & _kataban & ControlChars.Tab & "仕様書：" & _siyousyo & ControlChars.Tab & "出荷場所：" & _kataplace & ControlChars.Tab & "ﾁｪｯｸ区分：" & _katacheck & ControlChars.Tab & "GS価格：" & _gsprice
        Return _kataban & ControlChars.Tab & _siyousyo & ControlChars.Tab & _kataplace & ControlChars.Tab & _katacheck & ControlChars.Tab & _gsprice
    End Function
End Class
