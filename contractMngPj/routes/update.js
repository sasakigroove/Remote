//画面を表示する
var express = require('express');
var router = express.Router();
var HashMap = require('hashmap');

//チェックjsファイルを使用できるように読み込む（インスタンス化）
var C = require('./check.js');
var search = require('./search_execute.js');
var sheet_pulldown = require('./sheet_pulldown.js');

// HashMapのインスタンス化
var hmset = new HashMap();
var screenParamHsh = new HashMap();


const mysql2 = require('mysql2/promise');

//Mysqlのデータベース設定情報
var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'root',
    database: 'contractPj'
}


//初期表示要のメソッド
router.get('/', (req, res, next) => {
    var data = { msg: '', ROWS: "", hmset: '', NAME: '', checklinecount: '', pushButton: 0 };
    res.render('Search_List_view', data);
}
);




//取得してきた値（稼働終了技術者名,案件元,案件元営業,所属会社,所属会社営業,年月)
//更新OKボタン押された時
router.post('/updatestart', (req, res, next) => {
    updButtonOkPush(req,res);
});



async function updButtonOkPush(req,res) {

    // 画面情報取得
    var screenParam = await screenParamKeep(req);

    // 画面入力情報チェック
    var checkResult = await checkParam(screenParam);
    var checkResultNo = checkResult.get("checkResultNo");

    var msg = new HashMap();
    var resultDB = "";
    var updateResult;
    if(checkResultNo == 4){
        console.log("更新_04");
        updateResult = await update(screenParam);

        if(updateResult != ""){
            msg = updateResult;
        }

        screenParam.set("errorCount",0);

    }else{
        // エラー有りなら更新せずにそのまま
        screenParam.set("errorCount",checkResult.get("errorCount"));
        var errorMsg = checkResult.get("errorMsg");
        for(var i=0;i < errorMsg.size;i++){
            msg.set(i,errorMsg.get("msg"+i));
        }
    }

            // 検索処理
            resultDB = await search.searchcheck(
                screenParam.get("keyword"),             //ジャンルとキーワードで検索のキーワード
                screenParam.get("radioCond"),        //稼働期間で絞り込みのラジオボタン
                screenParam.get("genre_pulldown"),      //ジャンルとキーワードで検索のプルダウン
                screenParam.get("sheet_pulldown"),      //シートで検索のプルダウン
                screenParam.get("sort_button"),         //検索結果th部分のソートボタン　どのボタンが押されたか判定する用
                screenParam.get("start_before_year"),   //★稼働開始日で絞り込み年（前）
                screenParam.get("start_before_month"),  //★稼働開始日で絞り込み月（前）
                screenParam.get("start_after_year"),    //★稼働開始日で絞り込み年（後）
                screenParam.get("start_after_month"),   //★稼働終了日で絞り込み月（後）
                screenParam.get("end_before_year"),     //★稼働終了日で絞り込み年（前）
                screenParam.get("end_before_month"),    //★稼働終了日で絞り込み月（前）
                screenParam.get("end_after_year"),      //★稼働終了日で絞り込み年（後）
                screenParam.get("end_after_month"),     //★稼働終了日で絞り込み月（後）
                screenParam.get("sort1"),               //--↓ 各ソートボタンに画面でASCとDESCの保持をしておく用 ↓--
                screenParam.get("sort2"),
                screenParam.get("sort3") ,
                screenParam.get("sort4"), 
                screenParam.get("sort5"),
                screenParam.get("sort6") ,
                screenParam.get("sort7") ,
                screenParam.get("sort8") ,
                screenParam.get("sort9"),
                screenParam.get("sort10") ,
                screenParam.get("sort11") ,
                screenParam.get("sort12") ,      
                screenParam.get("sort13"),
                screenParam.get("sort14") ,
                screenParam.get("sort15") ,
                screenParam.get("sort16"),
                screenParam.get("sort17") ,
                screenParam.get("sort18") ,
                screenParam.get("sort19") ,             //--↑ 各ソートボタンに画面でASCとDESCの保持をしておく用 ↑--
                screenParam.get("errorcount"),            //★
                screenParam.get("excel_before_year"),   //★
                screenParam.get("excel_before_month"),  //★
                screenParam.get("excel_after_year"),    //★
                screenParam.get("excel_after_month"),   //★
                0          //★
            );

    var pushButton = 1;

    // 画面返却値の設定
    var data = await resultParamSet(msg,resultDB,screenParam,pushButton);

    // 検索画面へ遷移
    res.render('Search_List_view', data);
}

async function screenParamKeep(req){

        var checkcount = 0;
        var checklinecount = "";
    
        
        var recordCheckList = new HashMap();
    
        //よしこさんの検索結果画面の値保持-------------------------------------------------------------------------------------------------------------------------------------------------
    
        //検索結果数
        screenParamHsh.set("count",req.body.size);
        
        // ▼キーワード検索
        screenParamHsh.set("keyword",req.body.keyword);

        // 技術者名
        screenParamHsh.set("engineer_name",req.body.engineer_name);
    
        // ▼期間で絞り込み(0～8)
        screenParamHsh.set("radioCond",req.body.radioCond);
    
        screenParamHsh.set("start_before_pulldown1",req.body.start_before_pulldown1);
        screenParamHsh.set("start_before_pulldown2",req.body.start_before_pulldown2);
        screenParamHsh.set("start_after_pulldown1",req.body.start_after_pulldown1);
        screenParamHsh.set("start_after_pulldown2",req.body.start_after_pulldown2);
        screenParamHsh.set("end_before_pulldown1",req.body.end_before_pulldown1);
        screenParamHsh.set("end_before_pulldown2",req.body.end_before_pulldown2);
        screenParamHsh.set("end_after_pulldown1",req.body.end_after_pulldown1);
        screenParamHsh.set("end_after_pulldown2",req.body.end_after_pulldown2);
        screenParamHsh.set("genre_pulldown",req.body.genre_pulldown);
        screenParamHsh.set("sheet_pulldown",req.body.sheet_pulldown);
        screenParamHsh.set("sort_button",req.body.sort_button);
        
        screenParamHsh.set("excel_before_year",req.body.excel_before_year);
        screenParamHsh.set("excel_before_month",req.body.excel_before_month);
        screenParamHsh.set("excel_after_year",req.body.excel_after_year);
        screenParamHsh.set("excel_after_month",req.body.excel_after_month);

        
        // ▼ソートボタンが前回昇順/降順のどちらだったかの判別用
    

        screenParamHsh.set("sort1",req.body.sort1);
        screenParamHsh.set("sort2",req.body.sort2);
        screenParamHsh.set("sort3",req.body.sort3);
        screenParamHsh.set("sort4",req.body.sort4);
        screenParamHsh.set("sort5",req.body.sort5);
        screenParamHsh.set("sort6",req.body.sort6);
        screenParamHsh.set("sort7",req.body.sort7);
        screenParamHsh.set("sort8",req.body.sort8);
        screenParamHsh.set("sort9",req.body.sort9);
        screenParamHsh.set("sort10",req.body.sort10);
        screenParamHsh.set("sort11",req.body.sort11);
        screenParamHsh.set("sort12",req.body.sort12);
        screenParamHsh.set("sort13",req.body.sort13);
        screenParamHsh.set("sort14",req.body.sort14);
        screenParamHsh.set("sort15",req.body.sort15);
        screenParamHsh.set("sort16",req.body.sort16);
        screenParamHsh.set("sort17",req.body.sort17);
        screenParamHsh.set("sort18",req.body.sort18);
        screenParamHsh.set("sort19",req.body.sort19);

        
        for (var i = 0; i < req.body.size; i++) {

            var lineParamHsh = new HashMap();

            //チェックBOXの値
            var record = req.body["record" + i];
            if(record == null){
                record = 0;
            }

            //チェックBOXの判定
            if (record == 1) {

                console.log("チェック有りとなみなした");

                //稼働開始日
                lineParamHsh.set("startdate" + checkcount,req.body["start_date" + i]);
                lineParamHsh.set("startYear" + checkcount,req.body["start_date" + i].substr(0,4));
                lineParamHsh.set("startMonth" + checkcount,req.body["start_date" + i].substr(5,2));
                lineParamHsh.set("startDay" + checkcount,req.body["start_date" + i].substr(8,2));
    
                
                //プルダウンを保持するための稼働終了日
                lineParamHsh.set("endYear" + checkcount,req.body["endYear" + i]);
                lineParamHsh.set("endMonth" + checkcount,req.body["endMonth" + i]);
                lineParamHsh.set("endDay" + checkcount,req.body["endDay" + i]);
                //稼働終了年月日
                lineParamHsh.set("enddate" + checkcount,req.body["endYear" + i]+"/"+req.body["endMonth" + i]+"/"+req.body["endDay" + i])
    
                //技術者名
                lineParamHsh.set("engineer_name" + checkcount,req.body["engineer_name" + i]);

                //所属会社
                lineParamHsh.set("affiliation_company" + checkcount,req.body["affiliation_company" + i]);

                //所属会社営業
                lineParamHsh.set("affiliation_manager" + checkcount,req.body["affiliation_manager" + i]);

                //案件元営業
                lineParamHsh.set("proposition_manager" + checkcount,req.body["proposition_manager" + i]);

                //案件元
                lineParamHsh.set("proposition_company" + checkcount,req.body["proposition_company" + i]);

                //取り込み年月
                lineParamHsh.set("year_month_date" + checkcount,req.body["year_month_date" + i]);

                //行番号
                lineParamHsh.set("lineNo" + checkcount, i);

                //1行ごとの情報を画面情報に追加
                screenParamHsh.set("lineParam"+checkcount,lineParamHsh);

                //チェックされた行のみカウントされる
                checkcount++;
    
                if (checklinecount != "") {
                    checklinecount += ",";
                }
                checklinecount += i;
            } else {
                // チェック無しの行は何もしない
            }
    
            
            
            recordCheckList.set("record" + i,record);
            
        }

        screenParamHsh.set("recordCheckList",recordCheckList);
        screenParamHsh.set("checklinecount",checklinecount);
        screenParamHsh.set("checkcount",checkcount);

        return screenParamHsh;
}

// 画面入力情報チェック処理呼び出し
async function checkParam(screenParam){
    var checkResult = C.check(screenParam);

    return checkResult;
}

// 更新処理
async function update(screenParam) {
    console.log("更新始まる");

    var updateResult = "";

    // チェック行数を取得
    var checkcount = screenParam.get("checkcount");

    try {

        // チェックされた行分更新処理を行う
        for (var i = 0; i < checkcount; i++) {

            // 1行分の情報取得
            var lineParam = screenParam.get("lineParam"+i)

            // 稼働終了日の更新日付を取得
            var end_date = lineParam.get('enddate' + i);

            // 更新対象絞り込み情報取得
            var engineer_name = lineParam.get('engineer_name' + i);
            var proposition_company = lineParam.get('proposition_company' + i);
            var proposition_manager = lineParam.get('proposition_manager' + i);
            var affiliation_company = lineParam.get('affiliation_company' + i);
            var affiliation_manager = lineParam.get('affiliation_manager' + i);
            var year_month_date = lineParam.get('year_month_date' + i);


            //DB接続
            mycon = await mysql2.createConnection(mysql_setting);



            //更新 最新状態テーブル上の稼働終了日とメール送信済み判定を0に更新
            const [updateupdate_L_i] = await mycon.query('update Latest_information set end_date=\'' + end_date + '\',email_history=\'' + '0' + '\' WHERE engineer_name=\'' + engineer_name + '\' AND proposition_company=\'' + proposition_company + '\'AND proposition_manager=\'' + proposition_manager + '\'AND affiliation_company=\'' + affiliation_company + '\'AND affiliation_manager=\'' + affiliation_manager + '\'');
            console.log('更新OK')
            //更新 履歴テーブル上の稼働終了日更新
            const [updateupdate_M_h] = await mycon.query('update Management_history set end_date=\'' + end_date + '\' WHERE engineer_name=\'' + engineer_name + '\' AND proposition_company=\'' + proposition_company + '\'AND proposition_manager=\'' + proposition_manager + '\'AND affiliation_company=\'' + affiliation_company + '\'AND affiliation_manager=\'' + affiliation_manager + '\'AND year_month_date=\'' + year_month_date + '\'');
        }

        // DB接続解除
        mycon.end();

    } catch (err) {
        updateResult = "更新処理で異常終了しました。";
        console.log("SQLエラー内容→" + err);

    }


    return updateResult;
}

// 画面返却情報設定
async function resultParamSet(checkResultMsg,resultDB,screenParam,pushButton){

    var resultDBWork = "";
    var resultDBSizeWork = 0;
    var summary_judgement = "";
    var sum_data = "";
    var summary_hsh = "";
    var madika = "";
    var now_mm = "";
    var eigyoubetu_hantei = "";
    var eigyoubetu_name = "";
    var eigyoubetu_deta = "";
    var summary_check = "";
    var eigyoubetu_check = "";

    if(resultDB == ""){

    }else{
        resultDBWork = resultDB.get("rowsKey");
        resultDBSizeWork = resultDB.get("size");
        summary_judgement = resultDB.get("summary_judgement");
        sum_data = resultDB.get('sum');
        summary_hsh = resultDB.get('summary_hsh');
        madika = resultDB.get('madika');
        now_mm = resultDB.get('now_mm_hsh');
        eigyoubetu_hantei = resultDB.get('eigyoubetu_hantei');
        eigyoubetu_name = resultDB.get('eigyoubetu_name');
        eigyoubetu_deta = resultDB.get('eigyoubetu_deta');
        summary_check = resultDB.get('summary_check');
        eigyoubetu_check = resultDB.get('eigyoubetu_check');
    }
    
    var sheet_data = await sheet_pulldown.sheet_pulldown_search();
    var sheetDataWork = sheet_data.get("rowsKey");

    var resultData = {
        errorcount:screenParam.get("errorCount"),
        warningcount:screenParam.get("warningcount"),
        pushButton:pushButton,
        sheet_data:sheetDataWork,
        msg:checkResultMsg,
        ROWS:resultDBWork,
        size:resultDBSizeWork,
        summary_judgement:summary_judgement,
        sum_data:sum_data,
        summary_data:summary_hsh,
        madika: madika,
        now_mm:now_mm,
        eigyoubetu_hantei:eigyoubetu_hantei,
        eigyoubetu_name:eigyoubetu_name,
        eigyoubetu_deta:eigyoubetu_deta,
        summary_check:summary_check,
        eigyoubetu_check:eigyoubetu_check,
        deletemsg:'', // ★いらなくない？
        updatemsg:'', // ★いらなくない？
        startYear:screenParam.get("startYear"),
        startMonth:screenParam.get("startMonth"),
        startDay:screenParam.get("startDay"),
        start_after_year:screenParam.get("start_after_year"),
        start_after_month:screenParam.get("start_after_month"),
        end_before_year:screenParam.get("end_before_year"),
        end_before_month:screenParam.get("end_before_month"),
        end_after_year:screenParam.get("end_after_year"),
        end_after_month:screenParam.get("end_after_month"),
        excel_before_year:screenParam.get("excel_before_year"),
        excel_before_month:screenParam.get("excel_before_month"),
        excel_after_year:screenParam.get("excel_after_year"),
        excel_after_month:screenParam.get("excel_after_month"),
        engineer_name:screenParam.get("engineer_name"), // ★
        checklinecount:screenParam.get("checklinecount"),
        enddate:screenParam.get("end_date"),
        endYear:screenParam.get("endYear"),
        endMonth:screenParam.get("endMonth"),
        endDay:screenParam.get("endDay"),
        hmpurudata:screenParam.get("hmpurudata"),
        keyword:screenParam.get("keyword"),
        radioCond:screenParam.get("radioCond"),
        genre_pulldown:screenParam.get("genre_pulldown"),
        sheet_pulldown:screenParam.get("sheet_pulldown"),
        sort_button:screenParam.get("sort_button"),
        start_before_pulldown1:screenParam.get("start_before_pulldown1"),
        start_before_pulldown2:screenParam.get("start_before_pulldown2"),
        start_after_pulldown1:screenParam.get("start_after_pulldown1"),
        start_after_pulldown2:screenParam.get("start_after_pulldown2"),
        end_before_pulldown1:screenParam.get("end_before_pulldown1"),
        end_before_pulldown2:screenParam.get("end_before_pulldown2"),
        end_after_pulldown1:screenParam.get("end_after_pulldown1"),
        end_after_pulldown2:screenParam.get("end_after_pulldown2"),
        sort1:screenParam.get("sort1"),
        sort2:screenParam.get("sort2"),
        sort3:screenParam.get("sort3"),
        sort4:screenParam.get("sort4"),
        sort5:screenParam.get("sort5"),
        sort6:screenParam.get("sort6"),
        sort7:screenParam.get("sort7"),
        sort8:screenParam.get("sort8"),
        sort9:screenParam.get("sort9"),
        sort10:screenParam.get("sort10"),
        sort11:screenParam.get("sort11"),
        sort12:screenParam.get("sort12"),
        sort13:screenParam.get("sort13"),
        sort14:screenParam.get("sort14"),
        sort15:screenParam.get("sort15"),
        sort16:screenParam.get("sort16"),
        sort17:screenParam.get("sort17"),
        sort18:screenParam.get("sort18"),
        sort19:screenParam.get("sort19"),
        recordCheckList:screenParam.get("recordCheckList"),
    }

    return resultData;
}

module.exports = router;