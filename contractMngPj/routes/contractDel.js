//画面を表示する
var express = require('express');
var router = express.Router();
var HashMap = require('hashmap');
require('date-utils');

//チェックjsファイルを使用できるように読み込む（インスタンス化）
var search = require('./search_execute.js');
var sheet_pulldown = require('./sheet_pulldown.js');

// HashMapのインスタンス化
var screenParamHsh = new HashMap();

const mysql2 = require('mysql2/promise');

//Mysqlのデータベース設定情報
var mysql_setting = {
  host: 'localhost',
  user: 'root',
  password: 'root',
  database: 'contractPj'
}

let mycon;


//取得してきた値（稼働終了技術者名,案件元,案件元営業,所属会社,所属会社営業,年月)ーーーーーー
router.post('/contractDelCheck', (req, res, next) => {

  deleteButtonPush(req,res)

})

router.post('/contractDel', (req, res, next) => {

  deleteButtonOkPush(req,res)

})



async function deleteButtonOkPush(req,res){

  console.log("契約情報削除開始");

  var resultDB;
  var msg=new HashMap();
  msg.set(0,"更新完了しました");

  var screenParam = await screenParamKeep(req);
  var resultDel = await contractDelJob(screenParam);

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
    0         //★
);

  // 画面返却情報設定
  var data = await resultParamSet(msg,resultDB,screenParam,0);

  res.render('Search_List_view', data);
}

async function deleteButtonPush(req,res){

          var resultDB;
          var msg = new HashMap();
          msg.set(0,"");

          // 画面情報保持
          var screenParam = await screenParamKeep(req);

          var checkResult = await checkParam(screenParam);


          
          
          
          if(checkResult.get("errorCount") > 0){
            var errorMsg = checkResult.get("errorMsg");
            for(var i=0;i < checkResult.get("errorCount");i++){
                msg.set(i,errorMsg.get(i));
            }
            screenParam.set("errorCount",checkResult.get("errorCount"));
          }else if(checkResult.get("warningCount") > 0){
            var warningMsgList = checkResult.get("warningMsgList");
            for(var i=0;i < checkResult.get("warningCount") ;i++){
                msg.set(i,warningMsgList.get(i));
            }
            screenParam.set("warningcount",checkResult.get("warningCount"));
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
          
          var okNg;
          
          if(checkResult.get("errorCount") > 0){
            okNg = 0;
            console.log("msg:"+msg.get(0));
          }else{
            okNg = 2;
          }

          
          console.log("warningCount:"+checkResult.get("warningCount"));
          console.log("msg:"+msg.get(0));

          // 画面返却情報設定
          var data = await resultParamSet(msg,resultDB,screenParam,okNg);

          res.render('Search_List_view', data);

}

//jobAllでチェック処理を書くーーーーーーーーーーーーーーーーーーー
async function checkParam(screenParam) {

  //削除ボタン押下当日の日付取得
  var lineParam;
  

  var dt = new Date();
  var today1 = dt.toFormat("YYYY/MM/DD");
  var checkResult = new HashMap();
  var warningMsgList = new HashMap();
  var errorMsgList = new HashMap();

  var errCnt = 0;
  var warningCnt = 0;
  var msg = "";

  var checkcount = screenParam.get("checkcount");

  //チェックボックスにチェックが1つ以上あったかチェック
  if (checkcount != 0) {
  } else {
    //チェックボックスに何もチェックしていなかったらポップアップを表示
    msg = '必ず1つ以上チェックしてください';
  }
          //削除当日の日付を取得
          var dt = new Date();
          var formatted = dt.toFormat("YYYY/MM");
  

  //ファイル取り込み年月＜現在の情報については契約終了に更新できないようにチェック追加
  if(msg == ""){
      // ファイル取り込み年月のチェック
    for(var i = 0; i < checkcount;i++){
      
      lineParam = screenParam.get("lineParam"+i);

      var lineNo = lineParam.get("lineNo"+i)
      var nengetu = lineParam.get('year_month_date'+i);
    
      if(nengetu > formatted){
        //取り込み年月＞現在日時だった→警告ポップアップの表示
        warningMsgList.set(i,(i+1)+'行目に取り込み年月が当月以降のものがあります。本当に削除しますか？');
        warningCnt++;
        // break;
    
      }else if(nengetu < formatted){
        //取り込み年月＜現在日時だった→削除できないポップアップ
        errorMsgList.set(i,lineNo+"行目の取り込み年月が先月以前のものが含まれているため削除できません");
        errCnt++;
    
      }else if(nengetu == formatted){
      }

      if (msg == 0) {
          //選択された技術者が契約終了日以内であるかチェック
          var endDate = lineParam.get('enddate'+i);
        if (endDate　>= today1) {
          //稼働終了日が削除当日を過ぎていない場合は確認ポップアップを表示
          warningMsgList.set(i,lineNo+'行目の契約終了日以内のものが選択されていますが削除してよろしいですか？');
          warningCnt++;
        } else {
        }
      }

    }
  
  }else{
  }

  checkResult.set("warningCount",warningCnt);
  checkResult.set("warningMsgList",warningMsgList);
  checkResult.set("errorCount",errCnt);
  checkResult.set("errorMsg",errorMsgList);
  return checkResult;
}






//jobAllDeleteでチェック処理を書く
async function contractDelJob(screenParam) {

  var hmSelectResult = new HashMap();
  var msg = '';

  var lineParam = new HashMap();
  

  try {

        //削除当日の日付を取得
        var dt = new Date();
        var formatted = dt.toFormat("YYYY/MM");

    for (var i = 0; i < screenParam.get("checkcount"); i++) {

      lineParam = screenParam.get("lineParam"+i);

      //hmsetに入っている分だけ(チェックされた分だけ)削除更新処理実行
      //絞り込み条件　削除済み判定、技術者名、案件元、案件営業、所属会社、所属会社営業

      //DB接続
      mycon = await mysql2.createConnection(mysql_setting);

      //最新テーブルを更新(削除されたものを削除済みに更新)
      const [rowsUpdateLatest, fieldsLatest] = await mycon.query('UPDATE latest_information'
        + ' SET deleted = \'' + 1 + '\''
        + ' WHERE deleted = 0 AND engineer_name = \'' + lineParam.get("engineer_name" + i) + '\''
        + ' AND proposition_company = \'' + lineParam.get("proposition_company" + i) + '\''
        + ' AND proposition_manager = \'' + lineParam.get("proposition_manager" + i) + '\''
        + ' AND affiliation_company = \'' + lineParam.get("affiliation_company" + i) + '\''
        + ' AND affiliation_manager = \'' + lineParam.get("affiliation_manager" + i) + '\'');

      //履歴テーブルを更新(削除されたものを削除済みに更新)
      const [rowsUpdateHistory, fieldsHistory] = await mycon.query('UPDATE management_history'
        + ' SET deleted = \'' + 1 + '\''
        + ' WHERE year_month_date >=\''+lineParam.get('year_month_date'+i)+'\''
        + ' AND deleted = 0 AND engineer_name = \'' + lineParam.get("engineer_name" + i) + '\''
        + ' AND proposition_company = \'' + lineParam.get("proposition_company" + i) + '\''
        + ' AND proposition_manager = \'' + lineParam.get("proposition_manager" + i) + '\''
        + ' AND affiliation_company = \'' + lineParam.get("affiliation_company" + i) + '\''
        + ' AND affiliation_manager = \'' + lineParam.get("affiliation_manager" + i) + '\'');
    }



    //集計テーブルに新しい情報を更新するための検索(年月・シートNo・売上合計・稼働件数・あらりを取得)
    const [rowsSelect, fieldsSelect] = await mycon.query('SELECT year_month_date,sheet_No,sum(earnings)as uriage,COUNT(*) AS kadoukensu,SUM(gross_profit)as arari'
      + ' FROM management_history'
      + ' WHERE deleted = \'' + 0 + '\''
      // + ' AND year_month_date =\''+formatted+'\''
      + ' GROUP BY year_month_date,sheet_No'
      + ' ORDER BY sheet_No');

    for (var i = 0; i < rowsSelect.length; i++) {

      //年月の取得
      var year_month_date = rowsSelect[i].year_month_date;
      //シートナンバーの取得
      var sheet_No = rowsSelect[i].sheet_No;
      //売上合計を計算するための売上を取得
      var uriageSum = rowsSelect[i].uriage;
      //稼働件数を取得
      var kadoukensu = rowsSelect[i].kadoukensu;
      //粗利益の取得
      var arari = rowsSelect[i].arari;

      //売上平均の計算(売上合計÷稼働件数)
      var average_sales = uriageSum / kadoukensu;

      //平均利益の計算(粗利益÷稼働件数)
      var average_profit = arari / kadoukensu;

      console.log('年月、シートNo、売上、稼働件数、粗利益'+year_month_date,sheet_No,uriageSum,kadoukensu,arari);
      console.log('売上平均、平均利益'+average_sales,average_profit);

      //update
      //年月集計テーブル :Summary_tableを未削除の値のみで更新
      const [rowsUpdateLatest, fieldsLatest] = await mycon.query(' UPDATE Summary_table'
        + ' SET average_sales = \'' + average_sales + '\',average_profit = \'' + average_profit + '\',operating_count = \'' + kadoukensu + '\''
        + ' WHERE year_month_date = \'' + year_month_date + '\' AND sheet_No = \'' + sheet_No + '\'');
    }

    //DB接続解除
    mycon.end();

  } catch (error) {
      console.log('エラー発生→' + error);
      hmSelectResult.set('MSG', 'エラー発生');
  } finally {
  }

  return hmSelectResult;

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

  console.log("サイズ:"+req.body.size);

  for (var i = 0; i < req.body.size; i++) {

      var lineParamHsh = new HashMap();

      //チェックBOXの値
      var record = req.body["record" + i];
      if(record == null){
          record = 0;
      }

      //チェックBOXの判定
      if (record == 1) {

          console.log((i+1)+"行目をチェック有りとなみなした");

          

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
          lineParamHsh.set("enddate" + checkcount,req.body["endYear" + i]+"年"+req.body["endMonth" + i]+"月"+req.body["endDay" + i]+"日")

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
          lineParamHsh.set("lineNo" + checkcount, i+1);

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

      //検索結果分保持

      
  }

  screenParamHsh.set("recordCheckList",recordCheckList);
  screenParamHsh.set("checklinecount",checklinecount);
  screenParamHsh.set("checkcount",checkcount);

  return screenParamHsh;
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