/*******************************************************************************
【関数名】  : f_RodEnd_PortChk
【概要】    : ポートクッションニードルラジオチェック
【引数】    : strRdoid    :チェックされたラジオボタンのID
【戻り値】  : 無し
*******************************************************************************/
function f_OutOfOption_PortChk(strName, strRdoid) {
    var wk = ''
    switch (strRdoid.substring(0, 5)) {
        case 'rdoRK':
            document.getElementById(strName + "_HdnPortPlace1").value = strRdoid.substring(5, 6);
            break;
        case 'rdoRC':
            document.getElementById(strName + "_HdnPortPlace2").value = strRdoid.substring(5, 6);
            break;
        case 'rdoHK':
            document.getElementById(strName + "_HdnPortPlace3").value = strRdoid.substring(5, 6);
            break;
        case 'rdoHC':
            document.getElementById(strName + "_HdnPortPlace4").value = strRdoid.substring(5, 6);
            break;
    }
    //連結
    wk = document.getElementById(strName + "_HdnPortPlace1").value;
    wk = wk + document.getElementById(strName + "_HdnPortPlace2").value;
    wk = wk + document.getElementById(strName + "_HdnPortPlace3").value;
    wk = wk + document.getElementById(strName + "_HdnPortPlace4").value;
    //連結後の文字列をポート・クッションニードルテキストに表示
    document.getElementById(strName + "_txtPortCuchon").value = wk;
}

/*******************************************************************************
【関数名】  : f_RodEnd_TieRodChk
【概要】    : ポートクッションニードルラジオチェック
【引数】    : strRdoid    :チェックされたラジオボタン行番号
【戻り値】  : 無し
*******************************************************************************/
function f_OutOfOption_TieRodChk(strName, strRdoNo) {
    //タイロッド延長寸法の選択ラジオを保存
    document.getElementById(strName + "_HdnTieRodRdio").value = strRdoNo
}

/*******************************************************************************
【関数名】  : f_OKOnClick
【概要】    : [OK]ボタンonclick処理(Bottomフレーム)
【引数】    : 無し
【戻り値】  : 無し
*******************************************************************************/
function f_OutOfOption_OK(strName) {
    //以下、placelvl判定用の隠し項目に値を入れる処理を追加  2017/04/10 変更
    //ポート・クッションニードル
    if (!!document.getElementById(strName + "_cmbPortCushon")) {
        document.getElementById(strName + "_HdnSelPortCushon").value = document.getElementById(strName + "_cmbPortCushon").selectedIndex;
        document.getElementById(strName + "_HdnValPortCushon").value = document.getElementById(strName + "_cmbPortCushon").value;
    } else {
        document.getElementById(strName + "_HdnSelPortCushon").value = '0'
    }
    //ポート・クッションニードル位置
    if (!!document.getElementById(strName + "_txtPortCuchon")) {
        document.getElementById(strName + "_HdnSelPortPlace").value = document.getElementById(strName + "_txtPortCuchon").value;
    }
    //ポート２箇所
    if (!!document.getElementById(strName + "_cmbPort")) {
        document.getElementById(strName + "_HdnSelPort").value = document.getElementById(strName + "_cmbPort").selectedIndex;
        document.getElementById(strName + "_HdnValPort").value = document.getElementById(strName + "_cmbPort").value;
    } else {
        document.getElementById(strName + "_HdnSelPort").value = '0'
    }
    //ポートサイズダウン
    if (!!document.getElementById(strName + "_cmbPortSize")) {
        document.getElementById(strName + "_HdnSelPortSize").value = document.getElementById(strName + "_cmbPortSize").selectedIndex;
        document.getElementById(strName + "_HdnValPortSize").value = document.getElementById(strName + "_cmbPortSize").value;
    } else {
        document.getElementById(strName + "_HdnSelPortSize").value = '0'
    }
    //支持金具回転
    if (!!document.getElementById(strName + "_cmbMounting")) {
        document.getElementById(strName + "_HdnSelMounting").value = document.getElementById(strName + "_cmbMounting").selectedIndex;
        document.getElementById(strName + "_HdnValMounting").value = document.getElementById(strName + "_cmbMounting").value;
    } else {
        document.getElementById(strName + "_HdnSelMounting").value = '0'
    }
    //トラニオン位置指定
    if (!!document.getElementById(strName + "_txtTrunnion")) {
        document.getElementById(strName + "_HdnSelTrunnion").value = document.getElementById(strName + "_txtTrunnion").value;
    }
    //二山ナックル・二山クレビス
    if (!!document.getElementById(strName + "_cmbClevis")) {
        document.getElementById(strName + "_HdnSelClevis").value = document.getElementById(strName + "_cmbClevis").selectedIndex;
        document.getElementById(strName + "_HdnValClevis").value = document.getElementById(strName + "_cmbClevis").value;
    } else {
        document.getElementById(strName + "_HdnSelClevis").value = '0'
    }
    //タイロッド延長寸法ラジオ
    if (!!document.getElementById(strName + "_HdnTieRodRdio")) {
        document.getElementById(strName + "_HdnSelTieRod").value = document.getElementById(strName + "_HdnTieRodRdio").value;
    }
    //標準寸法
    if (!!document.getElementById(strName + "_lblDefault")) {
        document.getElementById(strName + "_HdnSelTieRodDefault").value = document.getElementById(strName + "_lblDefault").innerHTML;
    }
    //タイロッド材質SUS
    if (!!document.getElementById(strName + "_cmbSUS")) {
        document.getElementById(strName + "_HdnSelSUS").value = document.getElementById(strName + "_cmbSUS").selectedIndex;
        document.getElementById(strName + "_HdnValSUS").value = document.getElementById(strName + "_cmbSUS").value;
    } else {
        document.getElementById(strName + "_HdnSelSUS").value = '0'
    }
    //ジャバラ
    if (!!document.getElementById(strName + "_cmbJM")) {
        document.getElementById(strName + "_HdnSelJM").value = document.getElementById(strName + "_cmbJM").selectedIndex;
        document.getElementById(strName + "_HdnValJM").value = document.getElementById(strName + "_cmbJM").value;
    } else {
        document.getElementById(strName + "_HdnSelJM").value = '0'
    }
    //フッ素ゴム
    if (!!document.getElementById(strName + "_cmbFluoroRub")) {
        document.getElementById(strName + "_HdnSelFluoroRub").value = document.getElementById(strName + "_cmbFluoroRub").selectedIndex;
        document.getElementById(strName + "_HdnValFluoroRub").value = document.getElementById(strName + "_cmbFluoroRub").value;
    } else {
        document.getElementById(strName + "_HdnSelFluoroRub").value = '0'
    }
    //特注寸法
    if (!!document.getElementById(strName + "_txtTieRodCstm")) {
        //特注寸法テキストが存在する場合
        document.getElementById(strName + "_HdnSeltxtTieRodCstm").value = document.getElementById(strName + "_txtTieRodCstm").value;
    } else {
        //特注寸法テキストが存在しない場合、特注寸法コンボより設定
        document.getElementById(strName + "_HdnSelcmbTieRodCstm").value = document.getElementById(strName + "_cmbTieRodCstm").selectedIndex;
    }
}
