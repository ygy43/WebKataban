/*******************************************************************************
【関数名】  : f_Tanka_File
【概要】    : 単価のファイル出力
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_Tanka_File(strName) {
    //保存した情報をクリア
    document.getElementById(strName + "HidPriceForFile").value = "";
    //価格リストをクリア
    document.getElementById(strName + "HidPriceList").value = "";

    //掛率、単価、数量、金額、消費税と合計
    f_SetSelectInfo(strName);

    //価格リストの保存
    f_SetPriceList(strName);
}

/*******************************************************************************
【関数名】  : f_SetSelectInfo
【概要】    : 画面から入力した情報を取得して保存
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_SetSelectInfo(strName) {
    //掛率、単価、数量、金額、消費税と合計
    var strSelectInfo;
    //掛率    
    strSelectInfo = document.getElementById(strName + "txt_Rate").value + "|";
    //単価
    strSelectInfo = strSelectInfo + document.getElementById(strName + "TextUnitPrice").value + "|";
    //数量
    strSelectInfo = strSelectInfo + document.getElementById(strName + "TextCnt").value + "|";
    //金額
    strSelectInfo = strSelectInfo + document.getElementById(strName + "TextMoney").value;
    //消費税と合計
    if (document.getElementById(strName + "TextTax")) {
        strSelectInfo = strSelectInfo + "|" + document.getElementById(strName + "TextTax").value + "|";
        strSelectInfo = strSelectInfo + document.getElementById(strName + "TextAmount").value;
    } else {
        strSelectInfo += "|" + "|";
    }

    //選択情報の保存
    document.getElementById(strName + "HidPriceForFile").value = strSelectInfo;
}

/*******************************************************************************
【関数名】  : f_SetPriceList
【概要】    : 価格リストを取得して保存
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_SetPriceList(strName) {
    //価格リスト
    var list = "";

    $("#" + strName + "GVPrice" + " tr").each(function () {
        if (!this.rowIndex) return;

        var firstColumn = $(this).find("td:first");
        //タイトル
        var title = firstColumn.html();
        //価格
        var price = firstColumn.next().html();
        //価格区分
        var columnKbn = $(this).find("td:last").children(0).val();

        list += columnKbn + ":" + title + ":" + price + "|";

    });

    //価格リストの保存
    document.getElementById(strName + "HidPriceList").value = list;

}

/*******************************************************************************
【関数名】  : fncChangePlace
【概要】    : 出荷場所の変換（C11,P21）
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncChangePlace(strName, strMsg, strNewPlace, strPlace, strFRLMsg, strStockPlaceCdJa, strStockPlaceCdOs, strGLCMsg) {
    if (document.getElementById(strName + "cmbPlace")) {
        if (strMsg == ("")) {
            if (strStockPlaceCdOs != "") {
                if (confirm(strGLCMsg)) {
                    var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                    for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                        if (dropDownStrageEvaluation.options[i].value == (strStockPlaceCdOs)) {
                            dropDownStrageEvaluation.selectedIndex = i;
                            break;
                        }
                    }
                    // document.getElementById(strName + "HidNewPlace").value = strStockPlaceCdOs;
                    
                } else {
                    var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                    for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                        if (dropDownStrageEvaluation.options[i].value == (strPlace)) {
                            dropDownStrageEvaluation.selectedIndex = i;
                            break;
                        }
                    }
              // document.getElementById(strName + "HidNewPlace").value = "";
                }
            }

            if (strStockPlaceCdJa != "") {
                if (confirm(strGLCMsg)) {
                    var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                    for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                        if (dropDownStrageEvaluation.options[i].value == (strStockPlaceCdJa)) {
                            dropDownStrageEvaluation.selectedIndex = i;
                            break;
                        }
                    }
                   // document.getElementById(strName + "HidNewPlace").value = strStockPlaceCdJa;
                } else {
                    var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                    for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                        if (dropDownStrageEvaluation.options[i].value == (strPlace)) {
                            dropDownStrageEvaluation.selectedIndex = i;
                            break;
                        }
                    }
                   // document.getElementById(strName + "HidNewPlace").value = "";
                }
            }
        } else {
            if (confirm(strMsg)) {
                var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                    if (dropDownStrageEvaluation.options[i].value == (strNewPlace)) {
                        dropDownStrageEvaluation.selectedIndex = i;
                        break;
                    }
                }

               // document.getElementById(strName + "HidNewPlace").value = strNewPlace;

                if (strStockPlaceCdOs != "") {
                    if (confirm(strGLCMsg)) {
                        var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                        for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                            if (dropDownStrageEvaluation.options[i].value == (strStockPlaceCdOs)) {
                                dropDownStrageEvaluation.selectedIndex = i;
                                break;
                            }
                        }
                        // document.getElementById(strName + "HidNewPlace").value = strStockPlaceCdOs;
                    } else {
                        var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                        for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                            if (dropDownStrageEvaluation.options[i].value == (strNewPlace)) {
                                dropDownStrageEvaluation.selectedIndex = i;
                                break;
                            }
                        }
                    }
                }
            } else {
                var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                    if (dropDownStrageEvaluation.options[i].value == (strPlace)) {
                        dropDownStrageEvaluation.selectedIndex = i;
                        break;
                    }
                }
                //document.getElementById(strName + "cmbPlace").value = strPlace;
                document.getElementById(strName + "HidNewPlace").value = "";
                dropDownPlace.disabled = true;
                if (strStockPlaceCdJa != "") {
                    if (confirm(strGLCMsg)) {
                        var dropDownStrageEvaluation = document.getElementById(strName + "cmbStrageEvaluation");
                        for (var i = 0; i < dropDownStrageEvaluation.options.length; i++) {
                            if (dropDownStrageEvaluation.options[i].value == (strStockPlaceCdJa)) {
                                dropDownStrageEvaluation.selectedIndex = i;
                                break;
                            }
                        }
                       // document.getElementById(strName + "HidNewPlace").value = strStockPlaceCdJa;
                    }
                }
                if (strFRLMsg == "") {
                } else {
                    alert(strFRLMsg);
                }
            }
        }
    }

    //  dropDownPlace.disabled = true;
    if (dropDownStrageEvaluation != undefined) {
        dropDownStrageEvaluation.disabled = true;
    }
    if (document.getElementById(strName + "Label10") != null) {
        document.getElementById(strName + "Label10").style.visibility = 'hidden';
    }
    //document.getElementById(strName + "td").
    //document.getElementById(strName + "Label8").style.visibility = 'hidden';
}

/*******************************************************************************
【関数名】  : fnclblCheck
【概要】    : チェック区分TD非表示
【引数】    : strClsName         <String>   クラス名
【戻り値】  : 無し
*******************************************************************************/
function fnclblCheck(strClsName) {
    //document.getElementById(strID).style.visibility = 'hidden';

    var el1 = document.getElementsByClassName(strClsName);
    for (var i = 0; i < el1.length; ++i) {
        el1[i].style.visibility = 'hidden';
    };

}

/*******************************************************************************
【関数名】  : fncShowFrlMessage
【概要】    : FRLメッセージを表示する
【引数】    : strFrlMessage         <String>     FRLメッセージ
【戻り値】  : 無し
*******************************************************************************/
function fncShowFrlMessage(strFrlMessage) {
    if (strFrlMessage == "") {
    } else {
        alert(strFrlMessage);
    }
}

/*******************************************************************************
【関数名】  : fncShowFrlMessage
【概要】    : FRLメッセージを表示する
【引数】    : strFrlMessage         <String>     FRLメッセージ
【戻り値】  : 無し
*******************************************************************************/
function fncShowFrlMessage(strFrlMessage) {
    if (strFrlMessage == "") {
    } else {
        alert(strFrlMessage);
    }
}

/*******************************************************************************
【関数名】  : fncTanka_onKeyUp
【概要】    : キーアップイベント
【引数】    : strName         <String>    画面名
: strValue        <String>　　TextBox区分
【戻り値】  : 無し
*******************************************************************************/
function fncTanka_onKeyUp(strName, strValue) {
    var objEvent = window.event;

    if (objEvent.keyCode == 68) {
        if (objEvent.shiftKey == true) {
            var strName = ""
        }
    }
}

/*******************************************************************************
【関数名】  : fncTanka_onKeyDown
【概要】    : キーダウンイベント
【引数】    : strName         <String>    画面名
            : strValue        <String>　　TextBox区分
【戻り値】  : 無し
*******************************************************************************/
function fncTanka_onKeyDown(strName, strValue) {
    var objEvent = window.event;
    if (objEvent.keyCode == 13) {
        /* エンターキー押下時 */
        if (strValue == "Rate") {
            //単価ボックスへfocusを移動
            document.getElementById(strName + "TextUnitPrice").focus();
        } else if (strValue == "UnitPrice") {
            //数量ボックスへfocusを移動
            document.getElementById(strName + "TextCnt").focus();
        } else if (strValue == "Cnt") {
            //ボタンへfocusを移動
            //if (document.getElementById(strName + "btnOK")) { document.getElementById(strName + "btnOK").focus(); }
            document.getElementById(strName + "TextMoney").focus();
        }

        event.Returnvalue = false;
        return false;
    }
}

/*******************************************************************************
【関数名】  : f_UnitPriceCal
【概要】    : 単価計算
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_UnitPriceCal(strName, EditDiv) {
    var intRate;
    var intUnitPrice;
    var intRateUnitPrice;

    //単価
    intGPrice = document.getElementById(strName + "SelUnitValue").value;
    //通貨
    strCurrency = document.getElementById(strName + "SelCurrValue").value;

    //掛率(整数と少数の区切り(ドット)のみ)
    intRate = document.getElementById(strName + "txt_Rate").value;
    //単価計算(整数と少数の区切り(ドット)のみ)
    intUnitPrice = intGPrice * intRate;

    if (intGPrice != 0) {
        if (EditDiv == "0") {
            document.getElementById(strName + "txt_Rate").value = fncSetComma(fncRound(intRate, "4", "."));
            //document.getElementById(strName + "TextUnitPrice").value = fncSetComma(f_JudgeRound("fncRoundDown", intUnitPrice + 0.9, f_DecLen(0), "."));
            intUnitPrice = fncFloatMultiplication(intGPrice, intRate);
            document.getElementById(strName + "TextUnitPrice").value = fncSetComma(f_JudgeRound("fncRoundUp", intUnitPrice, f_DecLen(0, strCurrency), ".", strCurrency));

            document.getElementById(strName + "TextRateUnitPrice").value = fncSetComma(fncRound(intUnitPrice, f_DecLen(2, strCurrency), "."));
        } else {
            //カンマ編集・丸め編集
            intRateUnitPrice = f_CommaReplace(fncRound(intUnitPrice, f_DecLen(1, strCurrency), ","));
            intUnitPrice = f_CommaReplace(f_JudgeRound("fncRoundUp", intUnitPrice, f_DecLen(0, strCurrency), ",", strCurrency));
            document.getElementById(strName + "txt_Rate").value = fncSetDot(f_CommaReplace(intRate));
            document.getElementById(strName + "TextUnitPrice").value = fncSetDot(intUnitPrice);
            document.getElementById(strName + "TextRateUnitPrice").value = fncSetDot(intRateUnitPrice);
        }
    }
}

/*******************************************************************************
【関数名】  : f_RateCal
【概要】    : 掛け率計算
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_RateCal(strName, EditDiv) {
    var intRate;
    var intUnitPrice;
    var intRateUnitPrice;

    //単価
    intGPrice = document.getElementById(strName + "SelUnitValue").value;
    //通貨
    strCurrency = document.getElementById(strName + "SelCurrValue").value;

    if (intGPrice != 0) {
        if (EditDiv == "0") {
            //単価(整数と少数の区切り(ドット)のみ)
            intUnitPrice = fncRemoveComma(document.getElementById(strName + "TextUnitPrice").value);
            //掛率(整数と少数の区切り(ドット)のみ)
            intRate = intUnitPrice / intGPrice;
            //掛率(カンマ編集・丸め)
            document.getElementById(strName + "txt_Rate").value = fncSetComma(fncRoundUp(intRate, "4", "."));
            //単価(カンマ編集・丸め)
            document.getElementById(strName + "TextUnitPrice").value = fncSetComma(f_JudgeRound("fncRoundUp", intUnitPrice, f_DecLen(0, strCurrency), ".", strCurrency));
            //掛単価(カンマ編集・丸め)
            intRateUnitPrice = intGPrice * fncRoundUp(intRate, "3", ".");
            document.getElementById(strName + "TextRateUnitPrice").value = fncSetComma(fncRound(intRateUnitPrice, f_DecLen(2, strCurrency), "."));
        } else {
            //単価(整数と少数の区切り(ドット)のみ)
            intUnitPrice = fncRemoveDot(frmM.TextUnitPrice.value);
            //掛率(整数と少数の区切り(ドット)のみ)
            intRate = fncRoundUp(intUnitPrice / intGPrice, "3", ",");
            //掛率(カンマ編集・丸め)
            document.getElementById(strName + "txt_Rate").value = fncSetDot(f_CommaReplace(intRate));
            //単価(カンマ編集・丸め)
            document.getElementById(strName + "TextUnitPrice").value = fncSetDot(f_JudgeRound("fncRoundUp", intUnitPrice, f_DecLen(0, strCurrency), ",", strCurrency));
            //掛単価(カンマ編集・丸め)
            intRateUnitPrice = f_CommaReplace(fncRound(intGPrice * intRate, f_DecLen(2, strCurrency), ","));
            document.getElementById(strName + "TextRateUnitPrice").value = fncSetDot(intRateUnitPrice);
        }
    }
}

/*******************************************************************************
【関数名】  : f_MoneyCal
【概要】    : 合計計算
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_MoneyCal(strName, EditDiv) {
    var intRate;
    var intUnitPrice;
    var intCnt;
    var intRateUnitPrice;
    var intMoney;
    var intTax;
    var intAmount;
    intGPrice = document.getElementById(strName + "SelUnitValue").value;
    strCurrency = document.getElementById(strName + "SelCurrValue").value;
    if (intGPrice != 0) {
        if (EditDiv == "0") {
            //掛率(整数と少数の区切り(ドット)のみ)
            intRate = fncRemoveComma(document.getElementById(strName + "txt_Rate").value);

            //単価(整数と少数の区切り(ドット)のみ)
            intUnitPrice = fncRemoveComma(document.getElementById(strName + "TextUnitPrice").value);

            //数量(整数と少数の区切り(ドット)のみ)
            intCnt = fncRemoveComma(document.getElementById(strName + "TextCnt").value);

            if ((intCnt != null) && (intCnt != "")) {
                //金額 算出
                intMoney = fncRound(intUnitPrice * intCnt, f_DecLen(0, strCurrency), ".");
                //ベトナムの場合百位切捨てにする
                if (strCurrency == "VND") {
                    intMoney = Math.floor(intMoney / 1000) * 1000;
                }
                document.getElementById(strName + "TextMoney").value = fncSetComma(intMoney);

                //消費税 算出
                intTax = fncRoundDown(intMoney * 0.08, f_DecLen(0, strCurrency), ".");
                if (document.getElementById(strName + "TextTax")) { document.getElementById(strName + "TextTax").value = fncSetComma(intTax); }

                //合計 算出
                intAmount = fncRound(Number(intMoney) + Number(intTax), f_DecLen(0, strCurrency), ".");
                if (document.getElementById(strName + "TextAmount")) { document.getElementById(strName + "TextAmount").value = fncSetComma(String(intAmount)); }
            }
        } else {
            //掛率(整数と少数の区切り(ドット)のみ)
            intRate = f_DotReplace(fncRemoveDot(document.getElementById(strName + "txt_Rate").value));

            //単価(整数と少数の区切り(ドット)のみ)
            intUnitPrice = f_DotReplace(fncRemoveDot(document.getElementById(strName + "TextUnitPrice").value));

            //数量(整数と少数の区切り(ドット)のみ)
            intCnt = f_DotReplace(fncRemoveDot(document.getElementById(strName + "TextCnt").value));

            if ((intCnt != null) && (intCnt != "")) {
                //金額 算出
                intMoney = fncRound(intUnitPrice * intCnt, f_DecLen(0, strCurrency), ",");
                //ベトナムの場合百位切捨てにする
                if (strCurrency == "VND") {
                    intMoney = Math.floor(intMoney / 1000) * 1000;
                }
                document.getElementById(strName + "TextMoney").value = fncSetDot(intMoney);

                //消費税 算出
                intTax = fncRoundDown(intMoney * 0.08, f_DecLen(0, strCurrency), ",");
                if (document.getElementById(strName + "TextTax")) { document.getElementById(strName + "TextTax").value = fncSetDot(intTax); }

                //合計 算出
                intAmount = fncRound(Number(intMoney) + Number(intTax), f_DecLen(0, strCurrency), ",");
                if (document.getElementById(strName + "TextAmount")) { document.getElementById(strName + "TextAmount").value = fncSetDot(String(intAmount)); }
            }
        }
    }
}

/*******************************************************************************
【関数名】  : f_ShowPriceDetail
【概要】    : 価格詳細画面へ遷移
【戻り値】  : 無し
*******************************************************************************/
function f_ShowPriceDetail(strName) {
    if (document.getElementById(strName + "HidPriceDetail") != null) {
        if (document.getElementById(strName + "HidPriceDetail").value == "1") {
            document.getElementById(strName + "HidPriceDetail").value = "2";
            var timer = document.getElementById(strName + "btnClick");
            if (timer) { timer.click(); }
        }
    }
}

/********************************************************
calls the login/cad preview/cad generation process
********************************************************/
function call_cadenas(language, jsondata) {
    window.open('WebUserControl/Web2Cad.html?language=' + language + '&jsondata=' + jsondata, "CadWindow");
}
