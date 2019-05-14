Imports WebKataban.ClsCommon

Public Class CtlCharText
    Inherits System.Web.UI.WebControls.TextBox

#Region " Method "

    '********************************************************************************************
    '*【関数名】
    '*   CtlCharText_Init
    '*【処理】
    '*   初期処理
    '********************************************************************************************
    Private Sub CtlCharText_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        Try
            'クライアントサイドスクリプト(JavaScript)を追加
            Call Me.subJavaScriptSet()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   subJavaScriptSet
    '*【処理】
    '*   クライアントサイドスクリプト(JavaScript)を追加する
    '********************************************************************************************
    Private Sub subJavaScriptSet()
        Try
            ' フォーカス取得時の処理
            Call Me.subOnFocus()
            ' フォーカス喪失時の処理
            Call Me.subOnBlur()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   subOnFocus
    '*【処理】
    '*   フォーカス取得時の処理
    '********************************************************************************************
    Private Sub subOnFocus()
        Dim strJS As String
        Try
            strJS = ""
            strJS = strJS & "if (this.readOnly==false) { "
            strJS = strJS & " fncCmnTextOnFocus( this ); }"
            Me.Attributes.Add("onfocus", strJS)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   subOnBlur
    '*【処理】
    '*   フォーカス喪失時の処理
    '********************************************************************************************
    Private Sub subOnBlur()
        Dim strJS As String
        Try
            strJS = ""
            strJS = strJS & "if (this.readOnly==false) { "
            strJS = strJS & "    this.value = fncTextTrim( this.value ); "
            '文字変換(大文字・小文字)
            Select Case Me.CharCasing
                Case enCharCasing.Lower
                    strJS = strJS & " this.value = fncTextLCase( this.value ); "
                Case enCharCasing.Upper
                    strJS = strJS & " this.value = fncTextUCase( this.value ); "
            End Select
            '文字変換（ひらがな＞カナ）
            If ToKatakana Then
                strJS = strJS & " this.value = fncTextToKana( this.value ); "
            End If
            '文字変換(カナ)
            Select Case Me.ToKana
                Case enZenkakuHankaku.toHankaku
                    '全角＞半角
                    strJS = strJS & " this.value = fncTextToHankaku( this.value ); "
                Case enZenkakuHankaku.toZenkaku
                    '半角＞全角
            End Select
            '文字変換(アルファベット)
            Select Case Me.ToAlp
                Case enZenkakuHankaku.toHankaku
                    '全角＞半角
                    strJS = strJS & " this.value = fncTextToHankakuAlp( this.value ); "
                Case enZenkakuHankaku.toZenkaku
                    '半角＞全角
            End Select
            '文字変換(数字)
            Select Case Me.ToNum
                Case enZenkakuHankaku.toHankaku
                    '全角＞半角
                    strJS = strJS & " this.value = fncTextToHankakuNum( this.value ); "
                Case enZenkakuHankaku.toZenkaku
                    '半角＞全角
            End Select
            '文字変換(記号)
            Select Case Me.ToKigou
                Case enZenkakuHankaku.toHankaku
                    '全角＞半角
                    strJS = strJS & " this.value = fncTextToHankakuKigou( this.value ); "
                Case enZenkakuHankaku.toZenkaku
                    '半角＞全角
            End Select
            strJS = strJS & " if ( fncCheckChar( this.value, "
            strJS = strJS & CStr(Me.MinByte) & ", "
            strJS = strJS & CStr(Me.MaxByte) & ", "
            strJS = strJS & CStr(Me.AllowFullSizeChar).ToLower & " )"
            strJS = strJS & "==false ) { "
            strJS = strJS & " this.style.backgroundColor = '#FF0000'; "
            'NGの場合の処理
            If Me.CheckNgProc <> "" Then
                strJS = strJS & Me.CheckNgProc
            End If
            strJS = strJS & " } else { "
            strJS = strJS & " fncCmnTextOnBlur( this ); "
            'OKの場合の処理
            If Me.CheckOkProc <> "" Then
                strJS = strJS & Me.CheckOkProc
            End If
            strJS = strJS & " } "
            strJS = strJS & " } "
            Me.Attributes.Add("onblur", strJS)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub
#End Region
#Region " Property "

    '**********************************************************************************************
    '*【プロパティ】MinByte
    '*  最小桁数の設定・取得
    '**********************************************************************************************
    Public Property MinByte() As Integer
        Get
            Return CType(ViewState("MinByte"), Integer)
        End Get
        Set(ByVal Value As Integer)
            ViewState("MinByte") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】MaxByte
    '*  最大桁数の設定・取得
    '**********************************************************************************************
    Public Property MaxByte() As Integer
        Get
            Return CType(ViewState("MaxByte"), Integer)
        End Get
        Set(ByVal Value As Integer)
            ViewState("MaxByte") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】AllowFullSizeChar
    '*  全角入力の可・不可の設定・取得
    '**********************************************************************************************
    Public Property AllowFullSizeChar() As Boolean
        Get
            Return CType(ViewState("AllowFullSizeChar"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AllowFullSizeChar") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】CharCasing
    '*  大文字・小文字の制御
    '**********************************************************************************************
    Public Property CharCasing() As enCharCasing
        Get
            Return CType(ViewState("CharCasing"), Integer)
        End Get
        Set(ByVal Value As enCharCasing)
            ViewState("CharCasing") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】CheckOkProc
    '*  入力チェックＯＫ時に行うJavaScript処理
    '**********************************************************************************************
    Public Property CheckOkProc() As String
        Get
            Return CType(ViewState("CheckOkProc"), String)
        End Get
        Set(ByVal Value As String)
            ViewState("CheckOkProc") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】CheckNgProc
    '*  入力チェックＮＧ時に行うJavaScript処理
    '**********************************************************************************************
    Public Property CheckNgProc() As String
        Get
            Return CType(ViewState("CheckNgProc"), String)
        End Get
        Set(ByVal Value As String)
            ViewState("CheckNgProc") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrFlg
    '*  エラーフラグのＯＮ／ＯＦＦ
    '**********************************************************************************************
    Public WriteOnly Property ErrFlg() As Boolean
        'Get
        '    'none
        'End Get
        Set(ByVal Value As Boolean)
            If Value Then
                Call SetAttributes(Me, 7)
            Else
                Call SetAttributes(Me, 6)
            End If
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ToKatakana
    '*  ひらがな＞カナ変換の制御
    '**********************************************************************************************
    Public Property ToKatakana() As Boolean
        Get
            Return CType(ViewState("ToKatakana"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("ToKatakana") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ToKana
    '*  カナ変換の制御
    '**********************************************************************************************
    Public Property ToKana() As enZenkakuHankaku
        Get
            Return CType(ViewState("ToKana"), enZenkakuHankaku)
        End Get
        Set(ByVal Value As enZenkakuHankaku)
            ViewState("ToKana") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ToAlp
    '*  アルファベット変換の制御
    '**********************************************************************************************
    Public Property ToAlp() As enZenkakuHankaku
        Get
            Return CType(ViewState("ToAlp"), enZenkakuHankaku)
        End Get
        Set(ByVal Value As enZenkakuHankaku)
            ViewState("ToAlp") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ToNum
    '*  数字変換の制御
    '**********************************************************************************************
    Public Property ToNum() As enZenkakuHankaku
        Get
            Return CType(ViewState("ToNum"), enZenkakuHankaku)
        End Get
        Set(ByVal Value As enZenkakuHankaku)
            ViewState("ToNum") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ToKigou
    '*  数字変換の制御
    '**********************************************************************************************
    Public Property ToKigou() As enZenkakuHankaku
        Get
            Return CType(ViewState("ToKigou"), enZenkakuHankaku)
        End Get
        Set(ByVal Value As enZenkakuHankaku)
            ViewState("ToKigou") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property
#End Region

End Class

Public Enum enCharCasing
    Normal = 0
    Upper = 1
    Lower = 2
End Enum

Public Enum enZenkakuHankaku
    Normal = 0
    toZenkaku = 1
    toHankaku = 2
End Enum