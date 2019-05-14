Imports WebKataban.ClsCommon

Public Class CtlNumText
    Inherits System.Web.UI.WebControls.TextBox

#Region " Method "

    '********************************************************************************************
    '*【関数名】
    '*   CtlNumText_Init
    '*【処理】
    '*   初期処理
    '********************************************************************************************
    Private Sub CtlNumText_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
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
            strJS = strJS & " if ( this.readOnly==false ) { "
            '数値編集を解除する
            If Me.DispComma = True Then
                If Me.EditDiv = enEditDiv.Normal Then
                    'カンマを除去
                    strJS = strJS & " this.value = fncRemoveComma( this.value ); "
                Else
                    'ドットを除去
                    strJS = strJS & " this.value = fncRemoveDot( this.value );"
                End If
            End If
            strJS = strJS & " fncCmnTextOnFocus( this ); "
            strJS = strJS & " } "
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
            strJS = strJS & " if ( this.readOnly==false ) { "
            If Me.EditDiv = enEditDiv.Normal Then
                'カンマ編集
                strJS = strJS & " this.value = fncRemoveComma( this.value ); "
            Else
                'ドット編集
                strJS = strJS & " this.value = fncRemoveDot( this.value );"
            End If
            strJS = strJS & " if ( fncCheckNum( this.value, "
            strJS = strJS & CStr(Me.IntLen) & ", "
            strJS = strJS & LCase(CStr(Me.AllowMinus)) & ", "
            strJS = strJS & LCase(CStr(Me.AllowZero)) & " ,"
            strJS = strJS & LCase(CStr(Me.DenyNull)) & " ,"
            strJS = strJS & LCase(CStr(Me.EditDiv)) & " )"
            strJS = strJS & "==false ) { "
            strJS = strJS & " this.style.backgroundColor = '#FF0000'; "

            'NGの場合の処理
            If Me.CheckNgProc <> "" Then
                strJS = strJS & Me.CheckNgProc
            End If

            'strJS = strJS & "alert('" & strErrorMessage & "');"
            strJS = strJS & " } else { "
            strJS = strJS & " fncCmnTextOnBlur( this ); "

            'OKの場合の処理
            If Me.CheckOkProc <> "" Then
                strJS = strJS & Me.CheckOkProc
            End If

            '数値の少数部まるめ処理
            If Len(CStr(Me.DecLen)) <> 0 Then
                If Me.EditDiv = enEditDiv.Normal Then
                    strJS = strJS & " this.value = fncRound( this.value, " & CStr(Me.DecLen) & ", '.' ); "
                Else
                    strJS = strJS & " this.value = fncRound( this.value, " & CStr(Me.DecLen) & ", ',' ); "
                End If
            End If

            strJS = strJS & " } "

            '数値を編集する
            If Me.DispComma = True Then
                If Me.EditDiv = enEditDiv.Normal Then
                    'カンマ編集
                    strJS = strJS & " this.value = fncSetComma( this.value ); "
                Else
                    'ドット編集
                    strJS = strJS & " this.value = fncSetDot( this.value ); "
                End If
            End If

            strJS = strJS & " } "
            Me.Attributes.Add("onblur", strJS)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub
#End Region

#Region " Property "
    ''**********************************************************************************************
    ''*【プロパティ】StrLanguage
    ''*  エラーメッセージの表示言語
    ''**********************************************************************************************
    'Public Property strErrorMessage() As String
    '    Get
    '        Return CType(ViewState("strErrorMessage"), String)
    '    End Get
    '    Set(ByVal Value As String)
    '        ViewState("strErrorMessage") = Value
    '        Call Me.subJavaScriptSet()
    '    End Set
    'End Property

    '**********************************************************************************************
    '*【プロパティ】IntLen
    '*  整数部の最大桁数の設定・取得
    '**********************************************************************************************
    Public Property IntLen() As Integer
        Get
            Return CType(ViewState("IntLen"), Integer)
        End Get
        Set(ByVal Value As Integer)
            ViewState("IntLen") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】DecLen
    '*  少数部の最大桁数の設定・取得
    '**********************************************************************************************
    Public Property DecLen() As Integer
        Get
            Return CType(ViewState("DecLen"), Integer)
        End Get
        Set(ByVal Value As Integer)
            ViewState("DecLen") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】AllowMinus
    '*  マイナス入力の可・不可の設定・取得
    '**********************************************************************************************
    Public Property AllowMinus() As Boolean
        Get
            Return CType(ViewState("AllowMinus"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AllowMinus") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】AllowZero
    '*  ゼロ入力の可・不可の設定・取得
    '**********************************************************************************************
    Public Property AllowZero() As Boolean
        Get
            Return CType(ViewState("AllowZero"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AllowZero") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property


    '**********************************************************************************************
    '*【プロパティ】DenyNull
    '*  ＮＵＬＬ入力禁止の設定・取得
    '**********************************************************************************************
    Public Property DenyNull() As Boolean
        Get
            Return CType(ViewState("DenyNull"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("DenyNull") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】DispComma
    '*  カンマ編集のＯＮ／ＯＦＦ
    '**********************************************************************************************
    Public Property DispComma() As Boolean
        Get
            Return CType(ViewState("DispComma"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("DispComma") = Value
            Call Me.subJavaScriptSet()
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】EditDiv
    '*  数値編集区分
    '**********************************************************************************************
    Public Property EditDiv() As enEditDiv
        Get
            Return CType(ViewState("EditDiv"), Integer)
        End Get
        Set(ByVal Value As enEditDiv)
            ViewState("EditDiv") = Value
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
    '*  エラーフラグ
    '**********************************************************************************************
    Public WriteOnly Property ErrFlg() As Boolean
        'Get
        '    'none
        'End Get
        Set(ByVal Value As Boolean)
            If Value Then
                Call SetAttributes(Me, 5)
            Else
                Call SetAttributes(Me, 4)
            End If
        End Set
    End Property

#End Region

End Class

Public Enum enEditDiv
    Normal = 0
    Other = 1
End Enum
