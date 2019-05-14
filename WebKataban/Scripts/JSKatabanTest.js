/*******************************************************************************
【関数名】  : KatOutGotFocus
【概要】    : フォカスイベント
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function KatOutGotFocus(strName, strID) {
    //20141014 YGY FIREFOX対応
    if ($('#' + strName + 'PnlText').attr("disabled") != true) {
        var intID;

        intID = parseInt(strID.replace("txt", ""));
        if (document.getElementById(strName + "HidGotID").value != intID) {
            document.getElementById(strName + "HidDblClick").value = "0";
            //20141014 YGY FIREFOX対応
            $('#' + strName + 'PnlText').attr("disabled", true);

            document.getElementById(strName + "HidGotID").value = intID;
            document.getElementById(strName + strID).style.backgroundColor = "#FFCC33";
            var timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        }
    }
}

/*******************************************************************************
【関数名】  : KatOutLostFocus
【概要】    : フォカス失うイベント
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function KatOutLostFocus(strName, strID) {
    //20141014 YGY FIREFOX対応
    if ($('#' + strName + 'PnlText').attr("disabled") != true) {
        var intID;

        intID = parseInt(strID.replace("txt", ""), 10);
        document.getElementById(strName + "HidLostID").value = intID;
    }
}

/*******************************************************************************
【関数名】  : KatOutDblClick
【概要】    : ダブルクリックイベント(オプションリストから選択リストに移動する)
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function KatOutDblClick(strName, strID) {
    var intID;
    var strValue;
    if (document.getElementById(strID)) {
        if (document.getElementById(strID).cells(0).innerText !== undefined) {
            strValue = document.getElementById(strID).cells(0).innerText.replace(/(^\s+)|(\s+$)/g, "");
        } else {
            strValue = document.getElementById(strID).cells(0).textContent.replace(/(^\s+)|(\s+$)/g, "");
        }

        if (strValue.length <= 0) { strValue = "無記号"; }
        intID = document.getElementById(strName + "HidLostID").value;
        if (document.getElementById(strName + "txt" + intID).value == "") {
            document.getElementById(strName + "txt" + intID).value = 1;
        } else {
            document.getElementById(strName + "txt" + intID).value = parseInt(document.getElementById(strName + "txt" + intID).value) + 1;
        }
        document.getElementById(strName + "HidListValue").value = strValue;
        document.getElementById(strName + "HidDblClick").value = "1";
        if (document.getElementById(strName + "txt" + (parseInt(intID) + 1)) != null) {
            //document.getElementById(strName + "GVDetail").style.display = "none";
            //document.getElementById(strName + "Panel5").style.display = "none";
            document.getElementById(strName + "txt" + parseInt(intID)).focus();
            var timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        } else {//なければ、OKボタンへ
            document.getElementById(strName + "btnOutPut").focus();
            //document.getElementById(strName + "GVDetail").style.display = "none";
            //document.getElementById(strName + "Panel5").style.display = "none";
        }
    }
}

/*******************************************************************************
【関数名】  : KatOutSelDblClick
【概要】    : ダブルクリックイベント(選択リストからオプションリストに戻る)
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function KatOutSelDblClick(strName, strID) {
    var intID;
    var strValue;
    if (document.getElementById(strID)) {
        if (document.getElementById(strID).cells(0).innerText !== undefined) {
            strValue = document.getElementById(strID).cells(0).innerText.replace(/(^\s+)|(\s+$)/g, "");
        } else {
            strValue = document.getElementById(strID).cells(0).textContent.replace(/(^\s+)|(\s+$)/g, "");
        }

        intID = document.getElementById(strName + "HidLostID").value;
        if (document.getElementById(strName + "txt" + intID).value == "") {
            document.getElementById(strName + "txt" + intID).value = "";
        } else {
            document.getElementById(strName + "txt" + intID).value = parseInt(document.getElementById(strName + "txt" + intID).value) - 1;
        }
        document.getElementById(strName + "HidListValue").value = strValue;
        document.getElementById(strName + "HidDblClick").value = "2";
        if (document.getElementById(strName + "txt" + (parseInt(intID) + 1)) != null) {
            document.getElementById(strName + "GVSelect").style.display = "none";
            document.getElementById(strName + "txt" + parseInt(intID)).focus();
            var timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        } else {//なければ、OKボタンへ
            document.getElementById(strName + "btnOutPut").focus();
            document.getElementById(strName + "GVSelect").style.display = "none";
        }
    }
}

/*******************************************************************************
【関数名】  : KatOutConfirm
【概要】    : ダブルクリックイベント
【引数】    : strMsg           <String>    メッセージ
            　strName          <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function KatOutConfirm(strMsg, strName) {
    if (document.getElementById(strName + "btnNext")) {
        var timer;
        if (confirm(strMsg)) {
            document.getElementById(strName + "HidDblClick").value = "3";
            timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        } else {
            timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        }
    }
}

/*******************************************************************************
【関数名】  : MFHistorytest
【概要】    : ダブルクリックイベント
【引数】    : strMsg           <String>    メッセージ
            　strName          <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function MFHistorytest(strName, strKeyName, strKey) {
    if (document.getElementById(strName)) {
        var timer = document.getElementById(strName);
        if (timer) {
            if (document.getElementById(strKeyName)) {
                document.getElementById(strKeyName).value = strKey;
            }
            timer.click();
        }
    }
}

