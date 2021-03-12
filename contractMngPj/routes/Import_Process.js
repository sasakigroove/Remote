// 拡張機能の読み込み
var express = require('express');
var router = express.Router();
//HashMapの読み込み
var HashMap = require('hashmap');

var coCk = require('./Common_Check.js');    //Common_Check.js読み込み
var coGt = require('./Common_Get.js');      //Common_Get.js読み込み
var coIt = require('./Common_import.js');   //Common_Get.js読み込み
let XLSX = require('xlsx');

// 初期画面用メソッド
router.get('/', function (req, res, next) {
    var today = new Date();
    var inputHsh = new HashMap();
    inputHsh.set('path', '');
    inputHsh.set('Year', today.toFormat("YYYY"));
    inputHsh.set('Month', today.toFormat("MM"));
    var data = { hsh: inputHsh };
    res.render('import-page', data);
}
);

//画面情報受け取り
router.post('/proccess', (req, res, next) => {
    console.log('取り込み処理開始!!');
    var inputHsh = new HashMap();
    inputHsh.set('path', req.body.path);
    inputHsh.set('Year', req.body.Year);
    inputHsh.set('Month', req.body.Month);
    mainJob(inputHsh, res);
});

//メイン処理呼び出し・画面返却
async function mainJob(paramHsh, res) {
    var path = paramHsh.get('path');                                    //入力パス
    var date = paramHsh.get('Year') + "/" + paramHsh.get('Month');      //画面で入力された年と月を結合して年月とする
    var resultHsh = new HashMap();                                      //画面に返却するものを処理から受け取る

//取り込み処理メイン処理
    resultHsh = await Import_Model(path, date);

//画面入力・選択保持
    resultHsh.set('path',   paramHsh.get('path'));
    resultHsh.set('Year',   paramHsh.get('Year'));
    resultHsh.set('Month',  paramHsh.get('Month'));
    var data = { hsh: resultHsh };
    res.render('import-page', data);
    console.log("取り込み処理終了!!");
}

//取り込みメイン処理
async function Import_Model(path, date) {
    var resultHsh = new HashMap();
    resultHsh.set("errorResult", '0');   //エラー有無の判定用
    try {

//エラーチェック
        resultHsh = await errorCheck(path, date);

    //エラーが無ければ取り込み処理
        if (resultHsh.get("errorResult") == "0") {
            let workbook = XLSX.readFile(path, { cellDates: true });    //入力パス情報読み込み
            let sheet_name_list = workbook.SheetNames;                  //シート名(全部)
            var sheetRows = await coGt.getSheetMaster();                //シートマスタ情報検索取得

        //シートマスタのレコード数とExcelのシート数が一致したら取り込み処理を行う
            if (sheet_name_list.length == sheetRows.length) {

//取り込み処理
                resultHsh = await importProcess(path, date);
            }

        //取り込み処理でエラーが無ければ取り込み完了文言を指定
            if (resultHsh.get("errorResult") == "0") {
                var yearMonth = String(date).split('/');
                resultHsh.set('completed', yearMonth[0] + '月' + yearMonth[1] + '月分の取り込みが完了しました');
            } else {
                resultHsh.set('completed', '');
            }
        } else {
            resultHsh.set('completed', '');
        }
    } catch (e) {
        console.log(e.stack);
        console.log('予期せぬエラー発生');
        resultHsh.set('popUpError', '予期せぬエラーが発生しました');
    }
    return resultHsh;
}


//エラーチェック
async function errorCheck(path, date) {
    var errorHsh            = new HashMap();     //返却するエラー文言の全て
    var excelErrorHsh       = new HashMap();     //エクセル上で起きるエラー文言の全て
    var duplicationError    = new HashMap();     //営業データの重複エラー文言
    var inputError          = '';
    errorHsh.set("errorResult", '0');            //エラー有無の判定用
    try {

//画面必須項目チェック
        inputError = coCk.requiredCheck(path);
        errorHsh.set('error1', inputError);
        if (inputError != '') {
            errorHsh.set("errorResult", "1");
        } else {

//ファイルの存在チェック
            inputError = coCk.fileExistenceCheck(path);
            errorHsh.set('error1', inputError);
            if (inputError != "") {
                errorHsh.set("errorResult", "1");
            } else {

//ファイル形式チェック
                inputError = coCk.fileFormatCheck(path);
                errorHsh.set('error1', inputError);
                if (inputError != "") {
                    errorHsh.set("errorResult", "1");
                } else {
                    let workbook = XLSX.readFile(path, { cellDates: true }); //ファイルの読み込み
                    let sheet_name_list = workbook.SheetNames;               //シート名(全シート)       
                    var sheetRows = await coGt.getSheetMaster();             //シートマスタ情報検索

//ファイルシート数チェック
                    if (sheetRows.length != sheet_name_list.length) {
                        errorHsh.set("error1", 'シートが不足しているまたは、余計なシートが存在します');
                        errorHsh.set("errorResult", "1");
                    }
                }
            }
        }

//選択年月チェック
        inputError = coCk.selectionCheck(date);
        errorHsh.set("error2", inputError);
        if (inputError != '') {
            errorHsh.set("errorResult", "1");
        }

    //画面の入力エラーチェックが無ければExcel内の入力エラーチェックをする
        if (errorHsh.get("errorResult") == "0") {
            let workbook = XLSX.readFile(path, { cellDates: true });   //ファイルの読み込み
            let sheet_name_list = workbook.SheetNames;                 //シート名宣言
            errorHsh.set('sheetCount', sheet_name_list.length);        //エラーがある時に画面でエラーを表示させるためのシートの数
            var sheetRows = await coGt.getSheetMaster();               //シートマスタ情報取得

    //シートの項目数エラー文言に先に全て空文字を入れておく
            for (var z = 1; z <= sheet_name_list.length; z++) {
                errorHsh.set("sheetError" + z, "");
            }

    //シートマスタのレコード数とExcelファイルのシートが一致したらエラーチェックをする
            if (sheetRows.length == sheet_name_list.length) {

    //シートマスタのデータ数分繰り返し///////////////////////////////////////////////////////
                for (var a = 0; a < sheetRows.length; a++) {
                    var sheetOKNG = false;  //シート名が一致したかの判定用

    //ファイルのシート数分繰り返し/////////////////////////////////////////////////////////////
                    for (var b = 0; b < sheet_name_list.length; b++) {

        //シートマスタにあるシート名とファイルのシート名が一致⇒エラーチェック
                        if (sheetRows[a].sheet_name == sheet_name_list[b]) {
                            sheetOKNG = true;       //一致した事を示す

                            let sheetInformation    = workbook.Sheets[sheet_name_list[b]];             //シートの情報を代入
                            var terminalColumn      = sheetInformation['!ref'].match(/\:[A-Z+]/);
                            terminalColumn          = terminalColumn.toString().replace("\:", "");     //Excel入力されている末端列を取得

//ファイルの項目数チェック
                            inputError = await coCk.itemsCountCheck(sheet_name_list[b], sheetRows[a].summary_item, terminalColumn);

                            var summary = await coGt.getSummary(sheet_name_list[b]);                   //集計情報登録項目の検索取得
                            errorHsh.set("sheetError" + (a + 1), inputError);
                            if (inputError != "") {
                                errorHsh.set("errorResult", "1");
                            } else {

    //データ数取得(末端行-シート毎の余分な行数)
                                let endCol = sheetInformation['!ref'].match(/\:[A-Z+]([0-9]+)/)[1];         //末端行数          
                                var dataCount = 0;                                                          //営業データ数
                            //シート毎に減算する数値が変わる
                                if (summary == 1) {
                                    dataCount = endCol - 17;
                                } else if (summary == 2) {
                                    dataCount = endCol - 10;
                                } else if (summary == 0) {
                                    dataCount = endCol - 5;
                                }

                                console.log("データ行数:"+dataCount);

//ファイルの営業データ情報の正入力判別
                                excelErrorHsh = await excelCheck(sheet_name_list[b], sheetInformation, dataCount, date, endCol);
                                errorHsh.set('dataCount' + (a + 1), dataCount);         //各シートのデータ数
                                errorHsh.set("excelError" + (a + 1), excelErrorHsh);    //各シートのエラー文言
                                if (excelErrorHsh.get('errorResult') == '1') {
                                    errorHsh.set("errorResult", "1");
                                } else {

//データ重複チェック
                                    duplicationError = await coCk.duplicationCheck(sheet_name_list[b], sheetInformation, dataCount);
                                    errorHsh.set("duplicationError" + (a + 1), duplicationError);
                                    if (duplicationError != '') {
                                        errorHsh.set("errorResult", "1");
                                    }
                                }
                            }
                            break;      //エラーチェックができたら次のシート情報のエラーチェックに移る
                        }
                    }

//シート名チェック
                    if (!sheetOKNG) {
                        errorHsh.set("sheetError" + (a + 1), 'ファイルに' + sheetRows[a].sheet_name + 'シートが存在しないまたは、不正なシートが存在します。');
                        errorHsh.set("errorResult", "1");
                    }

                    errorHsh.set('summaryCount' + (a + 1), sheetRows[a].summary_item);     //エラー文言、シート毎の項目数判定用
                }
            }
        }
    } catch (e) {
        console.log(e.stack);
        console.log('エラーチェックで予期せぬエラー発生');
        throw e;
    }
    return errorHsh;
}

//Excel-正入力チェック
async function excelCheck(sheetName, sheetInfo, dataCount, date, endCol) {
    var returnHsh = new HashMap();
    returnHsh.set("errorResult", "0");
    var errorHsh = new HashMap();           //行数分のエラー情報全て
    errorHsh.set("errorResult", "0");       //エラー有無の判定用
    var inputError = '';
    var errorCheck = 0;                     //エラーチェック番号⇒必須項目チェックでエラーのある行/全て未入力の行は他のエラーチェックをしない

    try {

    //データ数分繰り返し///////////////////////////////////////////////////////////////////
        for (var x = 5; x < (dataCount + 5); x++) {
            errorHsh = new HashMap();           //行数分のエラー情報全て
            errorHsh.set("errorResult", "0");   //最初にまだエラー発生してないことを宣言

//Excel-必須項目チェック
            var itemHsh = new HashMap();        //チェックする項目
            var checkHsh = new HashMap();       //チェックする値

    //入力必須項目の項目名
            itemHsh.set(1, '案件元');
            itemHsh.set(2, '案件元営業');
            itemHsh.set(3, '所属会社');
            itemHsh.set(4, '所属会社営業');
            itemHsh.set(5, '技術者');
            itemHsh.set(6, '売上');

    //入力必須項目のデータの値
            checkHsh.set(1, sheetInfo['F' + x]);
            checkHsh.set(2, sheetInfo['G' + x]);
            checkHsh.set(3, sheetInfo['H' + x]);
            checkHsh.set(4, sheetInfo['I' + x]);
            checkHsh.set(5, sheetInfo['J' + x]);
            checkHsh.set(6, sheetInfo['L' + x]);

            var inputCount = 0;

        // 所属会社がA-STARか？

        if(sheetInfo['H' + x] != null && sheetInfo['H' + x].v == "A-STAR"){
            // A-STARの場合は必須項目チェック無しとする
        }else{

        //必須項目が全部入力有りか
            for (var a = 1; a <= 6; a++) {
                if (checkHsh.get(a) == null) {
                    inputCount++;
                }
            }

        //全部入力有り➡正常➡次のエラーチェックへ
            if (inputCount == 0) {
                errorCheck = 1;
                errorHsh.set("itemError0", '');

        //必須項目の入力が一つもない場合➡その後の入力チェックは無しで行う➡ここではなにもしない
            } else if (inputCount == 6) {
                errorCheck = 0;

        //どれかの入力がない⇒エラー文言を指定
            } else {
                errorCheck = 0;
                errorHsh.set("itemError0", sheetName + '-' + x + '行目-案件元・案件元営業・所属会社・所属会社営業・技術者名・売上は必須入力項目です');
                errorHsh.set("errorResult", "1");
            }

            var summary = await coGt.getSummary(sheetName);     //集計情報登録項目取得

        //項目毎のエラー文言の初期値を空文字で指定
            if (summary == 1 || summary == 0) {
                for (var i = 1; i <= 15; i++) {
                    errorHsh.set("itemError" + i, '');
                }
            } else if (summary == 2) {
                for (var i = 1; i <= 18; i++) {
                    errorHsh.set("itemError" + i, '');
                }
            }
        } 
        // A-STARの場合の必須項目チェック終わり

            //必須項目チェックが正常の時
            if (errorCheck == 1) {

                //1.稼働開始日のエラーチェック
                if (sheetInfo['B' + x] != null) {
                    //日付妥当性チェック
                    inputError = await coCk.dateValidityCheck(sheetName, '稼働開始日', sheetInfo['B' + x].v, x);
                    errorHsh.set('itemError1', inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //2.稼働終了日のエラーチェック
                if (sheetInfo['C' + x] != null) {
                    //日付妥当性チェック
                    inputError = await coCk.dateValidityCheck(sheetName, '稼働終了日', sheetInfo['C' + x].v, x);
                    errorHsh.set("itemError2", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {

                        //選択年月が履歴情報か最新情報かの判定取得
                        var dateDecision = coGt.getHistoryLatest(date);

                        //稼働終了日が過去日かチェック
                        if (dateDecision == 1) {
                            inputError = await coCk.operatingEndDateCheck(sheetName, '稼働終了日', sheetInfo['C' + x].v, x);
                            errorHsh.set("itemError2", inputError);
                            if (inputError != "") {
                                errorHsh.set("errorResult", "1");
                            }
                        }
                    }
                }


                if (errorHsh.get('errorResult') == '0') {
                    //稼働年月日の大小チェック
                    if (sheetInfo['B' + x] != null && sheetInfo['C' + x] != null) {
                        inputError = await coCk.dateRangeCheck(sheetName, sheetInfo['B' + x].v, sheetInfo['C' + x].v, x);
                        errorHsh.set("itemError2", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }


                //3.エンドのエラーチェック
                if (sheetInfo['D' + x] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, 'エンド', sheetInfo['D' + x].v, 150, x);
                    errorHsh.set("itemError3", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //4.元請のエラーチェック
                if (sheetInfo['E' + x] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, '元請', sheetInfo['E' + x].v, 150, x);
                    errorHsh.set("itemError4", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //5.案件元のエラーチェック
                if (sheetInfo['F' + x] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, '案件元', sheetInfo['F' + x].v, 150, x);
                    errorHsh.set("itemError5", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //6.案件元営業のエラーチェック
                if (sheetInfo['G' + x] != null) {
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '案件元営業', sheetInfo['G' + x].v, 30, x);
                    errorHsh.set("itemError6", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //全角文字チェック
                        inputError = await coCk.fullWidthCheck(sheetName, '案件元営業', sheetInfo['G' + x].v, x);
                        errorHsh.set("itemError6", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //7.所属会社のエラーチェック
                if (sheetInfo['H' + x] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, '所属会社', sheetInfo['H' + x].v, 60, x);
                    errorHsh.set("itemError7", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //8.所属会社営業のエラーチェック
                if (sheetInfo['I' + x] != null) {
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '所属会社営業', sheetInfo['I' + x].v, 30, x);
                    errorHsh.set("itemError8", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //全角文字チェック
                        inputError = await coCk.fullWidthCheck(sheetName, '所属会社営業', sheetInfo['I' + x].v, x);
                        errorHsh.set("itemError8", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //9.技術者のエラーチェック
                if (sheetInfo['J' + x] != null) {
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '技術者名', sheetInfo['J' + x].v, 30, x);
                    errorHsh.set("itemError9", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //全角文字チェック
                        inputError = await coCk.fullWidthCheck(sheetName, '技術者名', sheetInfo['J' + x].v, x);
                        errorHsh.set("itemError9", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //10.スキルのエラーチェック
                if (sheetInfo['K' + x] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, 'スキル', sheetInfo['K' + x].v, 60, x);
                    errorHsh.set("itemError10", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //11.売上のエラーチェック
                if (sheetInfo['L' + x] != null) {
                    var checkWord = Math.round(sheetInfo['L' + x].v);
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '売上', checkWord, 11, x);
                    errorHsh.set("itemError11", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //半角数字チェック
                        inputError = await coCk.halfWidthNumberCheck(sheetName, '売上', checkWord, x);
                        errorHsh.set("itemError11", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //12.仕入のエラーチェック
                if (sheetInfo['M' + x] != null) {
                    var checkWord = Math.round(sheetInfo['M' + x].v);
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '仕入', checkWord, 11, x);
                    errorHsh.set("itemError12", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //半角数字チェック
                        inputError = await coCk.halfWidthNumberCheck(sheetName, '仕入', checkWord, x);
                        errorHsh.set("itemError12", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //13.法定福利のエラーチェック
                if (sheetInfo['N' + x] != null) {
                    var checkWord = Math.round(sheetInfo['N' + x].v);
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '法定福利', checkWord, 11, x);
                    errorHsh.set("itemError13", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //半角数字チェック
                        inputError = await coCk.halfWidthNumberCheck(sheetName, '法定福利', checkWord, x);
                        errorHsh.set("itemError13", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //14.粗利益のエラーチェック
                if (sheetInfo['O' + x] != null) {
                    var checkWord = Math.round(sheetInfo['O' + x].v);
                    //文字数チェック
                    inputError = await coCk.lengthCheck(sheetName, '粗利益', checkWord, 11, x);
                    errorHsh.set("itemError14", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //半角数字チェック
                        inputError = await coCk.halfWidthMinusCheck(sheetName, '粗利益', checkWord, x);
                        errorHsh.set("itemError14", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //15.粗利率のエラーチェック
                if (sheetInfo['P' + x] != null && sheetInfo['L' + x] != null) {
                    if (sheetInfo['L' + x].v != 0) {
                        var checkWord = Math.round(sheetInfo['P' + x].v * 1000) / 10;
                        //半角文字チェック
                        inputError = await coCk.halfWidthCheck(sheetName, '粗利率', checkWord, x);
                        errorHsh.set("itemError15", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        } else {
                            //数値範囲チェック      
                            if (checkWord > 100.0) {
                                errorHsh.set("itemError15", sheetName + '-粗利率-' + x + '行目-100%以下で入力してください');
                                console.log(sheetName + '-粗利率-' + x + '行目-100%以下で入力してください');
                                errorHsh.set("errorResult", "1");
                            }
                        }
                    }
                }

                //データ項目数取得
                var summary = await coGt.getSummary(sheetName);

                //営業一覧シートでのエラーチェック
                if (summary == '2') {

                    //16.件数のエラーチェック
                    if (sheetInfo['Q' + x] != null) {
                        var checkWord = Math.round(sheetInfo['Q' + x].v * 10) / 10;
                        //半角文字チェック
                        inputError = await coCk.halfWidthCheck(sheetName, '件数', checkWord, x);
                        errorHsh.set("itemError16", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        } else {
                            //数値範囲チェック      
                            if (checkWord >= 100000.0) {
                                errorHsh.set("itemError16", sheetName + '-件数-' + x + '行目-十万件未満で入力してください');
                                errorHsh.set("errorResult", "1");
                            }
                        }

                    }

                    //17.実売上のエラーチェック
                    if (sheetInfo['R' + x] != null) {
                        var checkWord = Math.round(sheetInfo['R' + x].v);
                        //文字数チェック
                        inputError = await coCk.lengthCheck(sheetName, '実売上', checkWord, 11, x);
                        errorHsh.set("itemError17", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        } else {
                            //小数チェック
                            inputError = await coCk.halfWidthCheck(sheetName, '実売上', checkWord, x);
                            errorHsh.set("itemError17", inputError);
                            if (inputError != "") {
                                errorHsh.set("errorResult", "1");
                            }
                        }
                    }


                    //18.実粗利のエラーチェック
                    if (sheetInfo['S' + x] != null) {
                        var checkWord = Math.round(sheetInfo['S' + x].v);
                        //文字数チェック
                        inputError = await coCk.lengthCheck(sheetName, '実粗利', checkWord, 11, x);
                        errorHsh.set("itemError18", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        } else {
                            //小数チェック
                            inputError = await coCk.halfWidthMinusCheck(sheetName, '実粗利', checkWord, x);
                            errorHsh.set("itemError18", inputError);
                            if (inputError != "") {
                                errorHsh.set("errorResult", "1");
                            }
                        }
                    }
                }
            }

            //一行分のエラー文言を入れる
            if (errorHsh.get("errorResult") == "1") {

                returnHsh.set('errorResult', "1");
            }
            returnHsh.set('salesDataError' + x, errorHsh);
        }


        //データ項目数取得
        var summary = await coGt.getSummary(sheetName);
        errorHsh = new HashMap();
        errorHsh.set('errorResult', "0");
        //0以外⇒集計情報のエラーチェック
        if (summary != 0) {

            //1.平均売上
            if (sheetInfo['M' + (dataCount + 7)] != null) {
                var checkWord = Math.round(sheetInfo['M' + (dataCount + 7)].v);
                //データ数チェック
                inputError = await coCk.dataCountCheck(sheetName, '平均売上', checkWord, 11, dataCount + 7);
                errorHsh.set("itemError1", inputError);
                if (inputError != "") {
                    errorHsh.set("errorResult", "1");
                } else {
                    //小数チェック
                    inputError = await coCk.halfWidthMinusCheck(sheetName, '平均売上', checkWord, dataCount + 7);
                    errorHsh.set("itemError1", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }
            }


            //2.平均利益
            if (sheetInfo['M' + (dataCount + 8)] != null) {
                var checkWord = Math.round(sheetInfo['M' + (dataCount + 8)].v);
                //データ数チェック
                inputError = await coCk.dataCountCheck(sheetName, '平均利益', checkWord, 11, dataCount + 8);
                errorHsh.set("itemError2", inputError);
                if (inputError != "") {
                    errorHsh.set("errorResult", "1");
                } else {
                    //半角数字チェック
                    inputError = await coCk.halfWidthMinusCheck(sheetName, '平均利益', Math.floor(sheetInfo['M' + (dataCount + 8)].v), dataCount + 8);
                    errorHsh.set("itemError2", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }
            }

            //3.稼働件数
            if (sheetInfo['M' + (dataCount + 9)] != null) {
                var checkWord = Math.round(sheetInfo['M' + (dataCount + 9)].v * 10) / 10;
                //半角少数チェックチェック
                inputError = await coCk.halfWidthCheck(sheetName, '稼働件数', checkWord, dataCount + 9);
                errorHsh.set("itemError3", inputError);
                if (inputError != "") {
                    errorHsh.set("errorResult", "1");
                } else {
                    //数値の範囲チェック
                    if (checkWord >= 100000.0) {
                        errorHsh.set("itemError3", sheetName + '-稼働件数-' + (dataCount + 9) + '行目-十万件未満で入力してください');
                        errorHsh.set("errorResult", "1");
                    }
                }
            }

            //稼働一覧ページの集計情報
            if (summary == 1) {

                //4.末落ち件数のエラーチェック
                if (sheetInfo['M' + (dataCount + 10)] != null) {

                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, '末落ち件数', sheetInfo['M' + (dataCount + 10)].v, 6, dataCount + 10);
                    errorHsh.set("itemError4", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }

                //5.新規受注件数のエラーチェック
                if (sheetInfo['M' + (dataCount + 11)] != null) {
                    var checkWord = Math.round(sheetInfo['M' + (dataCount + 11)].v * 10) / 10;
                    //半角少数チェック
                    inputError = await coCk.halfWidthCheck(sheetName, '新規受注件数', checkWord, dataCount + 11);
                    errorHsh.set("itemError5", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //データ数チェック
                        inputError = await coCk.dataCountCheck(sheetName, '新規受注件数', checkWord, 6, dataCount + 11);
                        errorHsh.set("itemError5", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                } else {
                    errorHsh.set("itemError5", '稼働一覧-新規受注件数-' + (dataCount + 11) + '行目-必須入力項目です');
                    errorHsh.set("errorResult", "1");
                }

                //6.新規受注件数目標のエラーチェック
                if (sheetInfo['M' + (dataCount + 12)] != null) {
                    var checkWord = Math.round(sheetInfo['M' + (dataCount + 12)].v * 10) / 10;
                    //半角少数チェック
                    inputError = await coCk.halfWidthCheck(sheetName, '新規受注件数目標', checkWord, dataCount + 12);
                    errorHsh.set("itemError6", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //データ数チェック
                        inputError = await coCk.dataCountCheck(sheetName, '新規受注件数目標', checkWord, 6, dataCount + 12);
                        errorHsh.set("itemError6", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                } else {
                    errorHsh.set("itemError6", '稼働一覧-新規受注件数目標-' + (dataCount + 12) + '行目-必須入力項目です');
                    errorHsh.set("errorResult", "1");
                }


                //7.受注件数GAPのエラーチェック
                if (sheetInfo['M' + (dataCount + 13)] != null) {
                    var checkWord = Math.round(sheetInfo['M' + (dataCount + 13)].v * 10) / 10;

                    //半角少数(マイナス考慮)チェック/////////////////////////////////////////////////////////////////////////////////////////////////////////////////(12-22追加項目)
                    inputError = await coCk.halfWidthMinusCheck(sheetName, '受注件数GAP', checkWord, dataCount + 13);
                    errorHsh.set("itemError7", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    } else {
                        //データ数チェック
                        inputError = await coCk.dataCountCheck(sheetName, '受注件数GAP', checkWord, 6, dataCount + 13);
                        errorHsh.set("itemError7", inputError);
                        if (inputError != "") {
                            errorHsh.set("errorResult", "1");
                        }
                    }
                }

                //8.営業別受注件数のエラーチェック
                var inputErrorCount = 8;
                //営業別受注件数の数分繰り返し(営業別受注件数の一番上の行数から末端行まで)
                for (var s = (dataCount + 14); s <= endCol; s++) {
                    //項目名でエラーがあるか
                    var inputError = await coCk.salesOrderCountCheck(sheetName, sheetInfo['L' + s]);
                    errorHsh.set("itemError" + inputErrorCount, inputError);
                    if (inputError != "") {
                        errorHsh.set('errorResult', "1");
                    } else {
                        //営業別受注件数の入力チェッック
                        if (sheetInfo['L' + s] != null && sheetInfo['M' + s] != null) {
                            //データ数チェック
                            inputError = await coCk.dataCountCheck(sheetName, sheetInfo['L' + s].v, sheetInfo['M' + s].v, 14, s);
                            errorHsh.set("itemError" + inputErrorCount, inputError);
                            if (inputError != "") {
                                errorHsh.set("errorResult", "1");
                            }
                        }
                    }
                    inputErrorCount += 1;
                    if (errorHsh.get('errorResult') == "1") {
                        returnHsh.set('errorResult', "1");
                    }
                }


            } else if (summary == 2) {
                //4.A-SATRの入力チェック
                if (sheetInfo['M' + (dataCount + 10)] != null) {
                    //データ数チェック
                    inputError = await coCk.dataCountCheck(sheetName, 'A-SATR', sheetInfo['M' + (dataCount + 10)].v, 6, dataCount + 10);
                    errorHsh.set("itemError4", inputError);
                    if (inputError != "") {
                        errorHsh.set("errorResult", "1");
                    }
                }
            }
        }
        returnHsh.set('summaryError', errorHsh);
        if (errorHsh.get('errorResult') == "1") {
            returnHsh.set('errorResult', "1");
        }

    } catch (e) {
        console.log(e.stack);
        console.log('シート内の正入力チェックで予期せぬエラー発生');
        throw e;
    }
    return returnHsh;
}

//取り込み処理
async function importProcess(path, date) {
    console.log('取り込み処理(開始)');
    var resultHsh = new HashMap();
    resultHsh.set("errorResult", "0");
    var databaseError = '';
    var deleteSql = '';
    try {
        let workbook = XLSX.readFile(path, { cellDates: true });      //ファイルの読み込み
        let sheet_name_list = workbook.SheetNames;                 //シート名宣言
        //選択年月が履歴情報か最新情報かの判定取得
        var dateDecision = coGt.getHistoryLatest(date);

        //シートマスタ情報取得
        var sheetRows = await coGt.getSheetMaster();
        console.log('データ取り込み');

        //最新情報テーブルへの削除・登録処理
        if (dateDecision == 1) {

            //最新情報テーブルの削除処理  
            deleteSql = "DELETE FROM latest_information";
            databaseError = await coIt.deletingProcess(deleteSql);
            resultHsh.set('popUpError', databaseError);
            if (databaseError != '') {
                resultHsh.set('errorResult', "1");
            } else {
                //履歴テーブルへの削除処理
                deleteSql = "DELETE FROM management_history WHERE year_month_date = \"" + date + "\"";
                databaseError = await coIt.deletingProcess(deleteSql);
                resultHsh.set('popUpError', databaseError);
                resultHsh.set("sheetCount",sheetRows.length);

                if (databaseError != '') {
                    resultHsh.set('errorResult', "1");
                } else {
                    //シートマスタのデータ数分繰り返し
                    for (var i = 0; i < sheetRows.length; i++) {
                        for (var b = 0; b < sheet_name_list.length; b++) {
                            if (sheetRows[i].sheet_name == sheet_name_list[b]) {
                                let sheetInfo = workbook.Sheets[sheet_name_list[i]];
                                let endCol = sheetInfo['!ref'].match(/\:[A-Z+]([0-9]+)/)[1];       //末端行数
                                var subtractLineCount = 0;
                                var dataCount = 0;
                                var summary = sheetRows[i].summary_item;
                                var sheetNo = sheetRows[i].sheet_No;
                                var registrationCount = 0;
                                if (summary == 1) {
                                    dataCount = endCol - 17;
                                } else if (summary == 2) {
                                    dataCount = endCol - 10;
                                } else if (summary == 0) {
                                    dataCount = endCol - 5;
                                }

                                //データ数分繰り返し
                                for (var x = 5; x < (dataCount + 5); x++) {
                                    //必須項目の入力のある行を取り込む
                                    if (
                                        (sheetInfo['F' + x] != null && sheetInfo['G' + x] != null && sheetInfo['H' + x] != null && sheetInfo['I' + x] != null && sheetInfo['J' + x] != null && sheetInfo['J' + x] != null)
                                        ||
                                        sheetInfo['H' + x].v == "A-STAR") {

                                        var startDate = null;                                   //稼働開始日
                                        var endDate = null;                                     //稼働終了日
                                        var engineerName = sheetInfo['J' + x].v;                //技術者名★
                                        var eu = null;                                          //エンド
                                        var generalContractor = null;                           //元請

                                        var propositionCompany = ""                             //案件元★
                                        if( sheetInfo['F' + x] != null){
                                            propositionCompany = sheetInfo['F' + x].v;
                                        } 

                                        var earnings = 0                                        //売上
                                        if(Math.round(sheetInfo['L' + x] != null)){
                                            earnings = Math.round(sheetInfo['L' + x].v);
                                        }
                                        
                                        var propositionManager = "";                          //案件元営業★
                                        if( sheetInfo['G' + x] != null){
                                            propositionManager = sheetInfo['G' + x].v
                                        }
                                        
                                        var affiliationCompany = sheetInfo['H' + x].v;          //所属会社★

                                        var affiliationManager = "";                            //所属会社営業★
                                        if(sheetInfo['I' + x] != null){
                                            affiliationManager = sheetInfo['I' + x].v
                                        }          
                                        var skill = null;                                       //スキル
                                        var purchasePrice = 0;                                  //仕入
                                        var legalWelfareExpense = 0;                            //法定福利
                                        var grossProfit = 0;                                    //粗利益
                                        var grossMargin = 0;                                    //粗利率
                                        var propositionNumber = 0;                              //件数
                                        var realEarnings = 0;                                   //実売上
                                        var realGrossProfit = 0;                                //実粗利


                                        //稼働開始日
                                        if (sheetInfo['B' + x] != null) {
                                            // startDate = sheetInfo['B' + x].v.toFormat("YYYY/MM/DD");
                                            startDate = sheetInfo['B' + x].v
                                        }
                                        //稼働終了日
                                        if (sheetInfo['C' + x] != null) {
                                            // endDate = sheetInfo['C' + x].v.toFormat("YYYY/MM/DD");
                                            endDate = sheetInfo['C' + x].v
                                        }
                                        //エンド
                                        if (sheetInfo['D' + x] != null) {
                                            eu = sheetInfo['D' + x].v;
                                        }
                                        //元請
                                        if (sheetInfo['E' + x] != null) {
                                            generalContractor = sheetInfo['E' + x].v;
                                        }
                                        //スキル
                                        if (sheetInfo['K' + x] != null) {
                                            skill = sheetInfo['K' + x].v;
                                        }
                                        //仕入
                                        if (sheetInfo['M' + x] != null) {
                                            purchasePrice = Math.round(sheetInfo['M' + x].v);
                                        }
                                        //法定福利
                                        if (sheetInfo['N' + x] != null) {
                                            legalWelfareExpense = Math.round(sheetInfo['N' + x].v);
                                        }
                                        //粗利益
                                        if (sheetInfo['O' + x] != null) {
                                            grossProfit = Math.round(sheetInfo['O' + x].v);
                                        }
                                        //粗利率
                                        if (sheetInfo['P' + x] != null) {
                                            if (sheetInfo['L' + x].v != 0) {
                                                grossMargin = Math.round(sheetInfo['P' + x].v * 1000) / 10;
                                            }
                                        }
                                        //件数
                                        if (sheetInfo['Q' + x] != null) {
                                            propositionNumber = Math.round(sheetInfo['Q' + x].v * 10) / 10;
                                        }
                                        //実売上
                                        if (sheetInfo['R' + x] != null) {
                                            realEarnings = Math.round(sheetInfo['R' + x].v);
                                        }
                                        //実粗利
                                        if (sheetInfo['S' + x] != null) {
                                            realGrossProfit = Math.round(sheetInfo['S' + x].v);
                                        }

                                        //最新への登録
                                        //年月                稼働開始日                稼働終了日              技術者                      メールの送信済判定    削除済判定   エンド             元請                              案件先                                  売上                    案件元営業                        所属会社                             所属会社営業                                         スキル                 仕入                        法定福利                                     粗利益                      粗利率                      件数                             実売上                     実粗利                              シートNo                                                                                                         
                                        let insData = { 'year_month_date': date, 'start_date': startDate, 'end_date': endDate, 'engineer_name': engineerName, 'email_history': 0, 'deleted': 0, 'EU': eu, 'general_contractor': generalContractor, 'proposition_company': propositionCompany, 'earnings': earnings, 'proposition_manager': propositionManager, 'affiliation_company': affiliationCompany, 'affiliation_manager': affiliationManager, 'skill': skill, 'purchase_price': purchasePrice, 'legal_welfare_expense': legalWelfareExpense, 'gross_profit': grossProfit, 'gross_margin': grossMargin, 'proposition_number': propositionNumber, 'real_earnings': realEarnings, 'real_gross_profit': realGrossProfit, 'sheet_No': sheetNo };
                                        var preSql = 'INSERT INTO latest_information SET ?';
                                        //登録処理呼び出し
                                        databaseError = await coIt.registeringProcess(preSql, insData);      //最新情報登録処理
                                        resultHsh.set('popUpError', databaseError);
                                        if (databaseError != '') {
                                            resultHsh.set('errorResult', "1");
                                        }
                                        //履歴への登録
                                        let insData2 = { 'year_month_date': date, 'start_date': startDate, 'end_date': endDate, 'engineer_name': engineerName, 'deleted': 0, 'EU': eu, 'general_contractor': generalContractor, 'proposition_company': propositionCompany, 'earnings': earnings, 'proposition_manager': propositionManager, 'affiliation_company': affiliationCompany, 'affiliation_manager': affiliationManager, 'skill': skill, 'purchase_price': purchasePrice, 'legal_welfare_expense': legalWelfareExpense, 'gross_profit': grossProfit, 'gross_margin': grossMargin, 'proposition_number': propositionNumber, 'real_earnings': realEarnings, 'real_gross_profit': realGrossProfit, 'sheet_No': sheetNo };
                                        var preSql = 'INSERT INTO management_history SET ?';
                                        databaseError = await coIt.registeringProcess(preSql, insData2);     //履歴情報登録処理
                                        registrationCount += 1;
                                        resultHsh.set('popUpError', databaseError);
                                        if (databaseError != '') {
                                            resultHsh.set('errorResult', "1");
                                            break;
                                        }
                                    }
                                }
                                break;  //登録処理が終わったら次のfor文データに移る
                            }
                            if (resultHsh.get('errorResult') == '1') {
                                break;
                            }
                        }
                        if (resultHsh.get('errorResult') == '1') {
                            break;
                        }
                        resultHsh.set('completed' + i, sheetRows[i].sheet_name + '-' + registrationCount + '件');     //取り込み完了文言追加
                    }
                }
            }
        } else if (dateDecision == 0) {

            //履歴テーブルへの削除処理
            deleteSql = "DELETE FROM management_history WHERE year_month_date = \"" + date + "\"";
            databaseError = await coIt.deletingProcess(deleteSql);
            resultHsh.set('popUpError', databaseError);
            if (databaseError != '') {
                resultHsh.set('errorResult', "1");
            } else {

                //シートマスタ数分繰り返し
                for (var i = 0; i < sheetRows.length; i++) {
                    for (var b = 0; b < sheet_name_list.length; b++) {
                        if (sheetRows[i].sheet_name == sheet_name_list[b]) {
                            let sheetInfo = workbook.Sheets[sheet_name_list[b]];
                            let endCol = sheetInfo['!ref'].match(/\:[A-Z+]([0-9]+)/)[1];       //末端行数
                            var subtractLineCount = 0;
                            var dataCount = 0;
                            var summary = sheetRows[i].summary_item;
                            var sheetNo = sheetRows[i].sheet_No;
                            var registrationCount = 0;
                            if (summary == 1) {
                                subtractLineCount = 17;
                            } else if (summary == 2) {
                                subtractLineCount = 10;
                            } else if (summary == 0) {
                                subtractLineCount = 5;
                            }
                            dataCount = endCol - subtractLineCount;
                            //データ数分繰り返し

                            for (var x = 5; x < (dataCount + 5); x++) {
                                //必須項目の入力のある行を取り込む
                                if (sheetInfo['F' + x] != null && sheetInfo['G' + x] != null && sheetInfo['H' + x] != null && sheetInfo['I' + x] != null && sheetInfo['J' + x] != null && sheetInfo['J' + x] != null) {
                                    var startDate = null;                                   //稼働開始日
                                    var endDate = null;                                     //稼働終了日
                                    var engineerName = sheetInfo['J' + x].v;                //技術者名★
                                    var eu = null;                                          //エンド
                                    var generalContractor = null;                           //元請
                                    var propositionCompany = sheetInfo['F' + x].v;          //案件先★
                                    var earnings = Math.round(sheetInfo['L' + x].v);        //売上
                                    var propositionManager = sheetInfo['G' + x].v;          //案件元営業★
                                    var affiliationCompany = sheetInfo['H' + x].v;          //所属会社★
                                    var affiliationManager = sheetInfo['I' + x].v;          //所属会社営業★
                                    var skill = null;                                       //スキル
                                    var purchasePrice = 0;                                  //仕入
                                    var legalWelfareExpense = 0;                            //法定福利
                                    var grossProfit = 0;                                    //粗利益
                                    var grossMargin = 0;                                    //粗利率
                                    var propositionNumber = 0;                              //件数
                                    var realEarnings = 0;                                   //実売上
                                    var realGrossProfit = 0;                                //実粗利

                                    //稼働開始日
                                    if (sheetInfo['B' + x] != null) {
                                        startDate = sheetInfo['B' + x].v.toFormat("YYYY/MM/DD");
                                    }
                                    //稼働終了日
                                    if (sheetInfo['C' + x] != null) {
                                        endDate = sheetInfo['C' + x].v.toFormat("YYYY/MM/DD");
                                    }
                                    //エンド
                                    if (sheetInfo['D' + x] != null) {
                                        eu = sheetInfo['D' + x].v;
                                    }
                                    //元請
                                    if (sheetInfo['E' + x] != null) {
                                        generalContractor = sheetInfo['E' + x].v;
                                    }
                                    //スキル
                                    if (sheetInfo['K' + x] != null) {
                                        skill = sheetInfo['K' + x].v;
                                    }
                                    //仕入
                                    if (sheetInfo['M' + x] != null) {
                                        purchasePrice = Math.round(sheetInfo['M' + x].v);
                                    }
                                    //法定福利
                                    if (sheetInfo['N' + x] != null) {
                                        legalWelfareExpense = Math.round(sheetInfo['N' + x].v);
                                    }
                                    //粗利益
                                    if (sheetInfo['O' + x] != null) {
                                        grossProfit = Math.round(sheetInfo['O' + x].v);
                                    }
                                    //粗利率
                                    if (sheetInfo['P' + x] != null && sheetInfo['L' + x] != null) {
                                        if (sheetInfo['L' + x].v != 0) {
                                            grossMargin = Math.round(sheetInfo['P' + x].v * 1000) / 10;
                                        }
                                    }
                                    //件数
                                    if (sheetInfo['Q' + x] != null) {
                                        propositionNumber = Math.round(sheetInfo['Q' + x].v * 10) / 10;
                                    }
                                    //実売上
                                    if (sheetInfo['R' + x] != null) {
                                        realEarnings = Math.round(sheetInfo['R' + x].v);
                                    }
                                    //実粗利
                                    if (sheetInfo['S' + x] != null) {
                                        realGrossProfit = Math.round(sheetInfo['S' + x].v);
                                    }
                                    //履歴への登録
                                    let insData2 = { 'year_month_date': date, 'start_date': startDate, 'end_date': endDate, 'engineer_name': engineerName, 'deleted': 0, 'EU': eu, 'general_contractor': generalContractor, 'proposition_company': propositionCompany, 'earnings': earnings, 'proposition_manager': propositionManager, 'affiliation_company': affiliationCompany, 'affiliation_manager': affiliationManager, 'skill': skill, 'purchase_price': purchasePrice, 'legal_welfare_expense': legalWelfareExpense, 'gross_profit': grossProfit, 'gross_margin': grossMargin, 'proposition_number': propositionNumber, 'real_earnings': realEarnings, 'real_gross_profit': realGrossProfit, 'sheet_No': sheetNo };
                                    var preSql = 'INSERT INTO management_history SET ?';
                                    databaseError = await coIt.registeringProcess(preSql, insData2);     //履歴情報登録処理
                                    registrationCount += 1;
                                    resultHsh.set('popUpError', databaseError);
                                    if (databaseError != '') {
                                        resultHsh.set('errorResult', "1");
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                        if (resultHsh.get('errorResult') == '1') {
                            break;
                        }
                    }
                    if (resultHsh.get('errorResult') == '1') {
                        break;
                    }
                    resultHsh.set('completed' + i, sheetRows[i].sheet_name + '-' + registrationCount + '件');
                }
            }
        }

        if (resultHsh.get('errorResult') != '1') {

            //集計情報の取り込み
            //年月集計テーブルへの削除処理
            deleteSql = "DELETE FROM Summary_table WHERE year_month_date = \"" + date + "\"";
            databaseError = await coIt.deletingProcess(deleteSql);                //年月情報削除処理
            resultHsh.set('popUpError', databaseError);
            if (databaseError != '') {
                resultHsh.set('errorResult', "1");
            } else {

                //シート数分繰り返し
                for (var i = 0; i < sheetRows.length; i++) {
                    for (var b = 0; b < sheet_name_list.length; b++) {
                        if (sheetRows[i].sheet_name == sheet_name_list[b]) {
                            let sheetInfo = workbook.Sheets[sheet_name_list[b]];
                            let endCol = sheetInfo['!ref'].match(/\:[A-Z+]([0-9]+)/)[1];       //末端行数
                            var subtractLineCount = 0;
                            var dataCount = 0;
                            var summary = await coGt.getSummary(sheet_name_list[b]);

                            //集計情報のあるシート
                            if (summary != 0) {
                                var sheetNo = await coGt.getSheetNo(sheet_name_list[b]);
                                if (summary == 1) {
                                    subtractLineCount = 17;
                                } else if (summary == 2) {
                                    subtractLineCount = 10;
                                } else if (summary == 0) {
                                    subtractLineCount = 5;
                                }
                                dataCount = endCol - subtractLineCount;   //データの数
                                var summaryCount = (dataCount + 7);              //集計情報が書き始まる行数
                                var averageSales = 0;                            //平均売上
                                var averageProfit = 0;                            //平均利益
                                var operatingCount = 0.0;                          //稼働件数
                                var aStar = null;                         //A-SATR
                                var monthEndCount = null;                         //未落ち件数
                                var ordercountNew = 0;                            //新規受注件数
                                var ordercountNewTarget = 0;                            //新規受注件数目標
                                var ordercountGAP = 0;                            //受注件数GAP
                                let summaryData = null;


                                //平均売上
                                if (sheetInfo['M' + summaryCount] != null) {
                                    if (sheetInfo['M' + (summaryCount + 2)] != null) {          //稼働件数に0以外の自然数が入力されている時
                                        if (sheetInfo['M' + (summaryCount + 2)].v != 0) {
                                            averageSales = Math.round(sheetInfo['M' + summaryCount].v);
                                        }
                                    }
                                }
                                summaryCount += 1;
                                //平均利益
                                if (sheetInfo['M' + summaryCount] != null) {
                                    if (sheetInfo['M' + (summaryCount + 1)] != null) {          //稼働件数がnullまたは0の時→平均利益を0で登録する
                                        if (sheetInfo['M' + (summaryCount + 1)].v != 0) {
                                            averageProfit = Math.round(sheetInfo['M' + summaryCount].v);
                                        }
                                    }
                                }
                                summaryCount += 1;
                                //稼働件数
                                if (sheetInfo['M' + summaryCount] != null) {
                                    operatingCount = Math.round(sheetInfo['M' + summaryCount].v * 10) / 10;
                                }
                                summaryCount += 1;

                                if (summary == 2) {
                                    //A-STAR
                                    if (sheetInfo['M' + summaryCount] != null) {
                                        aStar = sheetInfo['M' + summaryCount].v;
                                    }
                                    //登録項目指定
                                    summaryData = { 'year_month_date': date, 'average_sales': averageSales, 'average_profit': averageProfit, 'operating_count': operatingCount, 'a_star': aStar, 'sheet_No': sheetNo };
                                } else if (summary == 1) {
                                    //未落ち件数
                                    if (sheetInfo['M' + summaryCount] != null) {
                                        monthEndCount = sheetInfo['M' + summaryCount].v;
                                    }
                                    summaryCount += 1;
                                    //新規受注件数
                                    if (sheetInfo['M' + summaryCount] != null) {
                                        ordercountNew = Math.round(sheetInfo['M' + summaryCount].v * 10) / 10;
                                    }
                                    summaryCount += 1;
                                    //新規受注件数目標
                                    if (sheetInfo['M' + summaryCount] != null) {
                                        ordercountNewTarget = Math.round(sheetInfo['M' + summaryCount].v * 10) / 10;
                                    }
                                    summaryCount += 1;
                                    //受注件数GAP
                                    if (sheetInfo['M' + summaryCount] != null) {
                                        ordercountGAP = Math.round(sheetInfo['M' + summaryCount].v * 10) / 10;
                                    }
                                    //登録項目指定
                                    summaryData = { 'year_month_date': date, 'average_sales': averageSales, 'average_profit': averageProfit, 'operating_count': operatingCount, 'month_end_count': monthEndCount, 'ordercount_new': ordercountNew, 'ordercount_new_target': ordercountNewTarget, 'ordercount_GAP': ordercountGAP, 'sheet_No': sheetNo };
                                }

                                //年月集計テーブルへの登録
                                var preSql = 'INSERT INTO summary_table SET ?';
                                databaseError = await coIt.registeringProcess(preSql, summaryData);      //年月情報登録処理
                                resultHsh.set('popUpError', databaseError);
                                if (databaseError != '') {
                                    resultHsh.set('errorResult', "1");
                                    break;
                                }
                            }
                            break;
                        }

                        if (resultHsh.get('errorResult') == '1') {
                            break;
                        }
                    }
                    if (resultHsh.get('errorResult') == '1') {
                        break;
                    }
                }
            }
        }

        if (resultHsh.get('errorResult') != '1') {

            //営業別受注件数の更新
            for (var b = 0; b < sheet_name_list.length; b++) {
                if (sheet_name_list[b] == '稼働一覧') {
                    let sheetInfo = workbook.Sheets[sheet_name_list[b]];
                    let endCol = sheetInfo['!ref'].match(/\:[A-Z+]([0-9]+)/)[1];       //末端行数
                    var summaryCount = endCol - 4;

                    //営業別受注件数の数分繰り返し(営業別受注件数の一番上の行数から末端行まで)
                    for (var s = summaryCount; s <= endCol; s++) {
                        if (sheetInfo['L' + s] != null && sheetInfo['M' + s] != null) {
                            var nameSp = String(sheetInfo['L' + s].v).split('月');  //月区切り
                            var nameRe = nameSp[1].replace('受注件数', '');  //受注件数を空文字へ

                            var updataSql = '';
                            var sheetNameRe = nameRe + '一覧';
                            var sheetNoNo = 0;

                            if (nameRe == '大倉') {
                                sheetNoNo = 1;
                                updataSql = 'UPDATE summary_table SET ordercount_sales = \'' + sheetInfo['M' + s].v + '\' WHERE year_month_date = \'' + date + '\' AND sheet_No = ' + sheetNoNo;
                            } else {
                                //名前から更新するシートNoを取得する
                                sheetNoNo = await coGt.getSheetNo(sheetNameRe);
                                updataSql = 'UPDATE summary_table SET ordercount_sales = \'' + sheetInfo['M' + s].v + '\' WHERE year_month_date = \'' + date + '\' AND sheet_No = ' + sheetNoNo;
                            }
                            //年月情報更新処理
                            databaseError = await coIt.updatingProcess(updataSql);
                            resultHsh.set('popUpError', databaseError);
                            if (databaseError != '') {
                                resultHsh.set('errorResult', "1");
                                break;
                            }
                        }
                    }
                    if (resultHsh.get('errorResult') == '1') {
                        break;
                    }
                }
            }
        }
    } catch (e) {
        console.log(e.stack);
        console.log('取り込み処理で予期せぬエラー発生');
        throw e;
    }
    return resultHsh;
}

module.exports = router;