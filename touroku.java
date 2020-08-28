package com.example.demo;

import java.io.UnsupportedEncodingException;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

import org.hibernate.annotations.Check;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class touroku {

	@Autowired
	private JdbcTemplate jdbcTemplate;

	@RequestMapping("/tourokugamen")
	public ModelAndView hello(@ModelAttribute tourokuInfo form) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

		System.out.println(form.getSname());
		System.out.println(form.getTu1());
		System.out.println(form.getTu2());
		System.out.println(form.getTu3());
		System.out.println(form.getyear());
		System.out.println(form.getmonth());
		System.out.println(form.getday());
		System.out.println(form.getseizou());
		System.out.println(form.getsyozaichi());

		String[] Tuusyouhairetu = {form.getTu1(),form.getTu2(),form.getTu3()};
		Hashtable hash = new Hashtable();
		Hashtable Snamehash = new Hashtable();
		Hashtable Tuusyouhash = new Hashtable();
		Hashtable hashT = new Hashtable();
		int result = 0;
		

		String date2 = form.getyear()+"/"+form.getmonth()+"/"+form.getday();

		try {
			if(form.getSname().length() > 0 && form.getTu1().length() > 0 && (form.getyear().length() > 0 && Integer.parseInt(form.getyear()) != 0) && (form.getmonth().length() > 0 && Integer.parseInt(form.getmonth()) != 0) && (form.getday().length() > 0 && Integer.parseInt(form.getday()) != 0)) {
				kyoutuu chk = new kyoutuu();
				String msgSname = chk.SnameCheck(form.getSname());
				if(msgSname.length() == 0) {
					Map selectMapsname = jdbcTemplate.queryForMap(" SELECT COUNT(*) as cnt FROM koujyoutable WHERE seishikiname = ?",form.getSname());
					System.out.println(selectMapsname);
					int count = Integer.parseInt((selectMapsname).get("cnt").toString());
					if(count == 0) {

						DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
						// 日付解析を厳密に行う設定にする(これがないと1/32が2/1で通ってしまう)
						df.setLenient(false);
						try {
							df.parse(date2);

							for(int n = 0; n <= 2; n++) {
								if(Tuusyouhairetu[n] != null && Tuusyouhairetu[n].length() > 0) {
									String msgTuusyou = chk.TuusyouCheck(Tuusyouhairetu[n]);
									Tuusyouhash.put("TUUSYOU"+n,Tuusyouhairetu[n]);
									if(msgTuusyou.length() == 0) {
									}else {
										hash.put("MSG",msgTuusyou);
										break;
									}
								}else {
									Tuusyouhash.put("TUUSYOU"+n,"");
								}
							}
							if(hash.size()> 0) {
							}else {
								int Id = 0;
								Map selectMapidcheck =jdbcTemplate.queryForMap(" SELECT max(id) as mx FROM koujyoutable");
								System.out.println(selectMapidcheck);
								if((selectMapidcheck).get("mx") == null) {
									System.out.println("IDがNULL");
									Id = 1;
								}else {
									System.out.println("IDのMAXは"+(selectMapidcheck).get("mx"));
									Id = Integer.parseInt((selectMapidcheck).get("mx").toString())+1;
									if(Id <= 99999) {
									}else {
										hash.put("MSG","工場IDが登録上限に達したためこれ以上登録できません");
									}
								}

								String ID = String.valueOf(Id);
								if(ID.length() == 1) {
									ID = "0000"+ID;
									System.out.println(ID);
								}else if(ID.length() == 2) {
									ID = "000" + ID;
									System.out.println(ID);
								}else if(ID.length() == 3) {
									ID = "00" + ID;
									System.out.println(ID);
								}else if(ID.length() == 4) {
									ID = "0" + ID;
									System.out.println(ID);
								}

								result  = jdbcTemplate.update("INSERT INTO koujyoutable(id,seishikiname,date,seizouno,syozaichino)VALUE(?,?,?,?,?)",ID,form.getSname(),date2,form.getseizou(),form.getsyozaichi());

								int i = 0;
								int TNo = 1;
								while(i < Tuusyouhash.size()) {
									result  = jdbcTemplate.update("INSERT INTO tuusyoutable(id,tuusyouno,tuusyou)VALUE(?,?,?)",ID,TNo,Tuusyouhash.get("TUUSYOU"+i).toString());
									i++;
									TNo++;
								}
								hash.put("MSG","登録が完了しました。工場IDは「"+ID+"」です。");
							}
						} catch (ParseException e) {
							// 日付妥当性NG時の処理を記述
							hash.put("MSG","創立年月日が正しく入力できていません");
						}

					}else {
						hash.put("MSG","入力された正式名称は既に存在するため登録できません");
					}
				}else {
					hash.put("MSG",msgSname);
				}
			}else {
				hash.put("MSG", "必須項目が入力されていません");
			}
		}catch(Exception e) {
			//↓エラー内容（赤字）を表示させる
			e.printStackTrace();
			//↓エラー時の処理
			hash.put("MSG", "エラーが発生しました");
			System.out.println("エラーが発生しました");
		}finally {
			ModelAndView mv = new ModelAndView();
			mv.addObject("MSG", hash.get("MSG").toString());

			List selectMap = null;
			List selectMapsyo = null;
			selectMap = (List) jdbcTemplate.queryForList("SELECT*FROM seizoumaster");
			selectMapsyo = (List) jdbcTemplate.queryForList("SELECT*FROM syozaichimaster");

			mv.setViewName("touroku");
			mv.addObject("emplist",selectMap);
			mv.addObject("emplistsyo",selectMapsyo);
			return mv;
		}
	}
	
	@RequestMapping("/modoru")
	public String topScreen() {
		return "kanri";
	}
}
