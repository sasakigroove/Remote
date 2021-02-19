//画面を表示する
var express = require('express');
var router = express.Router();
var HashMap = require('hashmap');


//チェックjsファイルを使用できるように読み込む（インスタンス化）
var update= require('./update.js');
// HashMapのインスタンス化
var hmset = new HashMap();
var hmget = new HashMap();
var hmpurudate = new HashMap();

const mysql2 = require('mysql2/promise');

//Mysqlのデータベース設定情報

var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'root',
    database: 'contractPj'

}

//DB接続情報を保持する為の変数

let mycom = null;


//初期表示要のメソッド
router.get('/', (req, res, next) => {
    
    var data = { msg: '', ROWS: '' ,NAME: '',checklinecount:'',pushButton:'',keyword:'',radio:'',genre_pulldown:'',sheet_pulldown:'',sort_button:'',
    start_before_pulldown1:'',start_before_pulldown2:'',start_after_pulldown1:'',start_after_pulldown2:'',
    end_before_pulldown1:'',end_before_pulldown2:'',
    end_after_pulldown1:'',end_after_pulldown2:'',
    sort1:'',sort2:'',sort3:'',sort4:'',sort5:'',sort6:'',sort7:'',sort8:'',sort9:'',sort10:'',sort11:'',sort12:'',sort13:'',errorcount:'',
    sort14:'',sort15:'',sort16:''};
   
   
    res.render('search', data);
    console.log("updateここまで動いたよ");
}
);
//取得してきた値（稼働終了日,エンド,元請,案件先,年月)

router.post('/searchbutton', (req, res, next) => {
    var year = req.body.year;
    console.log(year);
    
    var month = req.body.month;
    console.log(month);
    
    var day = req.body.day;
    console.log(day);
    //案件元営業
    var proposition_manager = req.body.proposition_manager;
    console.log(proposition_manager);
    //技術者名
    var engineer_name  = req.body.name;
    console.log(engineer_name);
    //所属会社
    var affiliation_company = req.body.affiliation_company;
    console.log(affiliation_company);
    //所属会社営業
    var affiliation_manager = req.body.affiliation_manager;
    console.log(affiliation_manager);
   //案件元
    var proposition_company = req.body.proposition_company;
    console.log(proposition_company);
    //年月
    var year_month_date = req.body.general_contractor;
    console.log(year_month_date);
    var checklinecount ="";
    var hmpurudate ="";
    var pushButton ="";
    var errorcount = "";
    mainJob(year,month,day,proposition_manager,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,checklinecount,hmpurudate,pushButton,errorcount,res);
});
async function mainJob(year,month,day,proposition_manager,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,checklinecount,hmpurudate,pushButton,errorcount,res){

    var hmwork = await jobAll(year,month,day,proposition_manager,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,checklinecount,hmpurudate,pushButton,errorcount,res);
    
    var size = hmwork.get('size');
    var NAME = hmwork.get('name');
    var work = hmwork.get('ROWS');
   
    
    var data= {msg: hmwork.get('MSG'), ROWS: work, size: hmwork.get('size'), NAME: hmwork.get('name'),checklinecount:checklinecount,hmpurudate:hmpurudate,pushButton:pushButton,
    keyword:'28',radio:'29',
    genre_pulldown:'1',sheet_pulldown:'2',sort_button:'3',
    start_before_pulldown1:'4',start_before_pulldown2:'5',start_after_pulldown1:'6',start_after_pulldown2:'7',
    end_before_pulldown1:'8',end_before_pulldown2:'9',
    end_after_pulldown1:'10',end_after_pulldown2:'11',
    sort1:'12',sort2:'13',sort3:'14',sort4:'15',sort5:'16',sort6:'17',sort7:'18',sort8:'19',sort9:'20',sort10:'21',sort11:'22',sort12:'23',sort13:'24',
    sort14:'25',sort15:'26',sort16:'27',errorcount:errorcount
};
   console.log("画面にわたす直前",data);
    res.render('search', data);

    
     } 



     async function jobAll(year,month,day,proposition_manager,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,res){
    if(engineer_name){
        rset = " WHERE engineer_name LIKE\"" + engineer_name + "%\"";
        console.log(rset);
        hmset.set("name",engineer_name);
        hmget = await namekensaku(rset,hmset);
    }else{
        
    }
    return hmget;
};







async function namekensaku(rset,hmset) {


    console.log('IDsyainame検索きた？')

    console.log(rset)

    //DB接続
    mycon = await mysql2.createConnection(mysql_setting);

    console.log('接続OK')

    const [rows, fields] = await mycon.query('select start_date,end_date,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,proposition_manager from testtable' + rset);

    console.log('検索かけた')


    console.log(rows)
    for(var s =0;s< rows.length; s++) {
   
   console.log(s);
    }

    if (rows.length == 0) {
        console.log('検索NG')
        //var msg = '検索結果がありませんでした';
        hmset.set('MSG','検索結果がありませんでした')
    } else {
        console.log('検索OK')
        //var msg = '検索結果がでました';
        
        hmset.set('MSG','検索結果が出ました')
        hmset.set('ROWS', rows)
        hmset.set('size', s);

       
    }
     // DB接続解除
     mycon.end();


     return hmset;
    
};
    
exports.namecheck = async function (rset) {


    

    console.log(rset)

    //DB接続
    mycon = await mysql2.createConnection(mysql_setting);

    console.log('二回目接続OK')

    const [rows, fields] = await mycon.query('select start_date,engineer_name,affiliation_company,affiliation_manager,proposition_company,year_month_date,proposition_manager from testtable' + rset);

    console.log('検索かけた')


    console.log(rows)
    for(var s =0;s< rows.length; s++) {
   
   console.log(s);
    }

    if (rows.length == 0) {
        console.log('検索NG')
        //var msg = '検索結果がありませんでした';
        hmset.set('MSG','検索結果がありませんでした')
    } else {
        console.log('検索OK')
        //var msg = '検索結果がでました';
        
        hmset.set('MSG','検索結果が出ました')
        hmset.set('ROWS', rows)
        hmset.set('size', s);

       
    }
    console.log("二回目",hmset);
     // DB接続解除
     mycon.end();


     return hmset;
    
};
    
    
    module.exports = router;