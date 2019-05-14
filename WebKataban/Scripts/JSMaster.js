/*******************************************************************************
【関数名】  : UserMasterCellClick
【概要】    : ユーザマスタ画面の選択イベント
【引数】    : strName         <String>    画面名
strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function UserMasterCellClick(strName, strID, strTableKey) {//ユーザーマスタ
    var strOldID = document.getElementById(strName + "_HidSelID").value;
    document.getElementById(strName + "_HidTableKey").value = strTableKey;
    if (strOldID != "") {
        if (document.getElementById(strOldID)) {
            document.getElementById(strOldID + "_pnlData").style.backgroundColor = "#CACAFF";
            document.getElementById(strOldID + "_Cell0").style.backgroundColor = "#FFFFCA";
            document.getElementById(strOldID + "_Cell0").style.color = "black";
        }
    }
    if (document.getElementById(strID + "_Cell0")) {
        document.getElementById(strID + "_pnlData").style.backgroundColor = "blue";
        document.getElementById(strID + "_Cell0").style.backgroundColor = "blue";
        document.getElementById(strID + "_Cell0").style.color = "white";
        document.getElementById(strName + "_HidSelID").value = strID;
        document.getElementById(strName + "_Button5").disabled = true;
        document.getElementById(strName + "_Button6").disabled = false;
        document.getElementById(strName + "_Button7").disabled = false;

        //更新する時に選択肢を有効にする
        SetCheckBoxEnabled(strName, "_pnlEditInput1");
        SetCheckBoxEnabled(strName, "_pnlEditInput2");
        SetCheckBoxEnabled(strName, "_pnlEditInput3");

        var strlblID;
        for (inti = 0; inti <= 7; inti++) {
            for (intj = 2; intj <= 4; intj++) {
                strlblID = "_Label" + intj + inti;
                if (document.getElementById(strName + strlblID)) {
                    document.getElementById(strName + strlblID).checked = document.getElementById(strID + strlblID).checked;
                    if (document.getElementById(strName + strlblID).checked) {
                        document.getElementById(strName + strlblID).parentNode.style.backgroundColor = "#CACAFF";
                        document.getElementById(strName + strlblID).parentNode.style.color = "red";
                    } else {
                        document.getElementById(strName + strlblID).parentNode.style.backgroundColor = "#C7EDCC";
                        document.getElementById(strName + strlblID).parentNode.style.color = "black";
                    }
                }
            }
        }
        for (inti = 5; inti <= 15; inti++) {
            strlblID = "_txtEdit" + inti;
            if (document.getElementById(strName + strlblID)) {
                if (inti == 5) { document.getElementById(strName + strlblID).disabled = true; }
                if (inti == 15) {
                    var e = document.getElementById(strName + strlblID);
                    var strUser = e.options[e.selectedIndex].value;
                    var flag = false;
                    for (var i = 0; i < e.options.length; i++) {
                        var option = e.options[i];
                        if (document.getElementById(strID + "_txtEdit16").innerText !== undefined) {
                            if (option.value == document.getElementById(strID + "_txtEdit16").innerText) {
                                if (flag) {
                                    option.selected = false;
                                } else {
                                    option.selected = true;
                                    flag = true;
                                }
                            } else {
                                option.selected = false;
                            }
                        } else {
                            if (option.value == document.getElementById(strID + "_txtEdit16").textContent) {
                                if (flag) {
                                    option.selected = false;
                                } else {
                                    option.selected = true;
                                    flag = true;
                                }
                            } else {
                                option.selected = false;
                            }
                        }
                    }
                } else {
                    if (document.getElementById(strID + strlblID).innerText !== undefined) {
                        document.getElementById(strName + strlblID).value = document.getElementById(strID + strlblID).innerText;
                    } else {
                        document.getElementById(strName + strlblID).value = document.getElementById(strID + strlblID).textContent;
                    }
                }
            }
        }
        //端末情報の設定
        if (document.getElementById(strName + "_pnlWebLog")) {
            if (document.getElementById(strID + "_HdnWebPassNo")) {
                document.getElementById(strName + "_txtPassword").value = document.getElementById(strID + "_HdnWebPassNo").value;
                document.getElementById(strName + "_txtMacAddress").value = document.getElementById(strID + "_HdnMacNo").value;
                document.getElementById(strName + "_txtSerial").value = document.getElementById(strID + "_HdnSerialNo").value;
                document.getElementById(strName + "_txtLastUsedTime").value = document.getElementById(strID + "_HdnLastDateNo").value;
            } else {
                document.getElementById(strName + "_txtPassword").value = "";
                document.getElementById(strName + "_txtMacAddress").value = "";
                document.getElementById(strName + "_txtSerial").value = "";
                document.getElementById(strName + "_txtLastUsedTime").value = "";
                //document.getElementById(strName + "_txtLastUsedTime").value = new Date().Format("yyyy/MM/dd hh:mm:ss");
            }
        }
    }
}

/*******************************************************************************
【関数名】  : CountryMasterCellClick
【概要】    : 国マスタ画面の選択イベント
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function CountryMasterCellClick(strName, strID, strTableKey) {//国別生産品マスタと掛率マスタ
    var strOldID = document.getElementById(strName + "_HidSelID").value;
    document.getElementById(strName + "_HidTableKey").value = strTableKey;
    if (strOldID != "") {
        if (document.getElementById(strOldID)) {
            var strRowID = strOldID.substr(strOldID.length - 1, 1);
            var strBKcolor;
            if (strRowID % 2 == 0) {
                strBKcolor = "white";
            } else {
                strBKcolor = "#CACAFF";
            }
            document.getElementById(strOldID + "_pnlData").style.backgroundColor = strBKcolor;
            document.getElementById(strOldID + "_Cell0").style.backgroundColor = strBKcolor;
            document.getElementById(strOldID + "_Cell0").style.color = "black";
        }
    }
    if (document.getElementById(strID + "_Cell0")) {
        document.getElementById(strID + "_pnlData").style.backgroundColor = "blue";
        document.getElementById(strID + "_Cell0").style.backgroundColor = "blue";
        document.getElementById(strID + "_Cell0").style.color = "white";
        document.getElementById(strName + "_HidSelID").value = strID;
        document.getElementById(strName + "_Button5").disabled = true;
        document.getElementById(strName + "_Button6").disabled = false;
        document.getElementById(strName + "_Button7").disabled = false;

        var strlblID;
        for (inti = 3; inti <= 12; inti++) {
            strlblID = "_txtEdit" + inti;
            if (document.getElementById(strName + strlblID)) {
                if (inti == 3 || inti == 4) { document.getElementById(strName + strlblID).disabled = true; } //国別生産品情報
                if (document.getElementById(strName + "_txtEdit3") == null) {
                    if (inti == 5 || inti == 6 || inti == 7) { document.getElementById(strName + strlblID).disabled = true; }
                }
                if (document.getElementById(strID + strlblID).innerText !== undefined) {
                    document.getElementById(strName + strlblID).value = document.getElementById(strID + strlblID).innerText;
                } else {
                    document.getElementById(strName + strlblID).value = document.getElementById(strID + strlblID).textContent;
                }
            }
        }
    }
}

/*******************************************************************************
【関数名】  : RadioClick
【概要】    : 購入価格と現地定価
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function RadioClick(strName, intMode) {
    if (intMode == 1) {
        if (document.getElementById(strName + "RadioButton1")) {
            if (document.getElementById(strName + "RadioButton1").checked) {
                document.getElementById(strName + "RadioButton2").checked = false;
                document.getElementById(strName + "HidRateDiv").value = "0";
            }
        }
    }
    if (intMode == 2) {
        if (document.getElementById(strName + "RadioButton2")) {
            if (document.getElementById(strName + "RadioButton2").checked) {
                document.getElementById(strName + "RadioButton1").checked = false;
                document.getElementById(strName + "HidRateDiv").value = "1";
            }
        }
    }
}

/*******************************************************************************
【関数名】  : SetCheckBoxEnabled
【概要】    : 権限選択肢を有効にする
【戻り値】  : 無し
*******************************************************************************/
function SetCheckBoxEnabled(strName, strPanelName) {
    try {
        document.getElementById(strName + strPanelName).disabled = false;
        var controls = document.getElementById(strName + strPanelName).childNodes;
        for (var i = 0; i < controls.length - 1; i++) {
            controls[i].disabled = false;
        }
        var controls = document.getElementById(strName + strPanelName).getElementsByTagName("input");
        for (var j = 0; j < controls.length; j++) {
            controls[j].disabled = false;
        }
    } catch (err) {
        alert(err.Message);
    }
}

/*******************************************************************************
【関数名】  : SetGSClientInfo
【概要】    : 国内GSの場合自動的に端末情報をセットする
【戻り値】  : 無し
*******************************************************************************/
function SetGSClientInfo(strName, strClientID) {
    var drplist = document.getElementById(strName + '_' + strClientID);
    var strSelected = drplist.options[drplist.selectedIndex].value;

    if (strSelected == 14) {
        if (document.getElementById(strName + '_txtPassword') !== null) {
            if (document.getElementById(strName + '_txtPassword').innerText !== undefined) {
                document.getElementById(strName + '_txtPassword').innerText = 'WebforGS';
                document.getElementById(strName + '_txtMacAddress').innerText = '40:61:86:2B:A6:AE';
                document.getElementById(strName + '_txtSerial').innerText = '99001';
            } else {
                document.getElementById(strName + '_txtPassword').textContent = 'WebforGS';
                document.getElementById(strName + '_txtMacAddress').textContent = '40:61:86:2B:A6:AE';
                document.getElementById(strName + '_txtSerial').textContent = '99001';
            }
        }
    }

}