function find_0176_en(xlDoc, sel) {
	var rng, bkSel, line, result='';

	line = ExpandParameter("$y");

	//xl.DisplayAlerts = False;
	//xlDoc = "C:\\1002\\doc\\クレンジングDB設計書\\D365_0176_仕入先_cleansing_VendVendorV2Entity.xlsx";
	//xlDoc = "C:\\1002\\doc\\クレンジングDB設計書\\D365_0629_取引先_cleansing_earth_md_corporation.xlsx";
	bk = xl.WorkBooks.Open(xlDoc, 0, 1);
	sRngNm = "DB4属性名_30文字制限考慮後確定値";
	//rng = bk.WorkSheets('#176_仕入先(リハ2向け)').Range(sRngNm);
	rng = bk.names("DB4属性名_30文字制限考慮後確定値").RefersToRange;

	//TraceOut("----->rng=" + sel + "<");
	c = rng(1,1);
    c0 = c;
	for (i=0; i< 10; i++) {
		TraceOut("c0:"+c0.Address + ", c.Offset(1):"+c.Offset(1,0).Address);
		if (i == 0) {
			c = rng.Find(sel, c);
		} else {
			c = rng.FindNext(c);
		}

		if (c == null) {
			break;
		} 

		if (c.Address == c0.Address) {
			TraceOut("break out == ...");
			break;
		}
		if (i==0) c0 = c;
			s = c.Value + "," + c.Offset(0, 1).Value;

		result += (s + "\n");
	}

	//TraceOut("findCol_176: <<< " + rng.Value);

	//for (var i=1; i<16; i++) {
	//	var en=new Enumerator(rng);
	//	if (!rng(1,1).Offset(0, i).Value)
	//		break;
	//	sel = bkSel;
	//	//	en.moveFirst();
	//	while(!en.atEnd()) {
	//		var p = en.item();
	//		if (!p.Value)
	//			break;
	//		if (p.Offset(0, i).Value)
	//			sel = sel.replace(new RegExp(p.Value, "mg"), p.Offset(0, i).Value)
	//		en.moveNext();
	//	}
	//	result += sel;
	//}

	TraceOut("line:" + line);
	if (result) {
		Jump(line, 1);
		Down();
		InsText(bk.Name + " ------\n");
		InsText(result);
	}
	return result;
}

function find_0176_jp(xlDoc, sel) {
	var rng, bkSel, line, result='';

	line = ExpandParameter("$y");

	//xl.DisplayAlerts = False;
	//xlDoc = "C:\\1002\\doc\\クレンジングDB設計書\\D365_0176_仕入先_cleansing_VendVendorV2Entity.xlsx";
	//xlDoc = "C:\\1002\\doc\\クレンジングDB設計書\\D365_0629_取引先_cleansing_earth_md_corporation.xlsx";
	bk = xl.WorkBooks.Open(xlDoc, 0, 1);
	sRngNm = "属性名_日本語名";
	rng = bk.names("属性名_日本語名").RefersToRange;

	//TraceOut("----->rng=" + sel + "<");
	c = rng(1,1);
    c0 = c;
	for (i=0; i< 10; i++) {
		TraceOut("c0:"+c0.Address + ", c.Offset(1):"+c.Offset(1,0).Address);
		if (i == 0) {
			c = rng.Find(sel, c);
		} else {
			c = rng.FindNext(c);
		}

		if (c == null) {
			break;
		} 

		if (c.Address == c0.Address) {
			TraceOut("break out == ...");
			break;
		}
		if (i==0) c0 = c;
			s = c.Value + "," + c.Offset(0, -1).Value;

		result += (s + "\n");
	}

	if (result) {
		Jump(line, 1);
		InsText(bk.Name + " ------\n");
		InsText(result);
	}
	return result;
}

function main() {
	xl = GetObject("","Excel.Application");

	xlDocs = [
		"C:\\1002\\doc\\クレンジングDB設計書\\D365_0629_取引先_cleansing_earth_md_corporation.xlsx",
		"C:\\1002\\doc\\クレンジングDB設計書\\D365_0176_仕入先_cleansing_VendVendorV2Entity.xlsx",
		"C:\\1002\\doc\\クレンジングDB設計書\\D365_0026_事業所_cleansing_account.xlsx",
		"C:\\1002\\doc\\クレンジングDB設計書\\D365_0129_事業所_cleansing_CustCustomerV3.xlsx"
		];

	sel = GetSelectedString();
	TraceOut("sel:" + sel);
	if (!sel)
		sel = GetLineStr(0).replace(/\r?\n/, '');

	for (f in xlDocs) {
		TraceOut(">>" + xlDocs[f]);
		if (!find_0176_en(xlDocs[f], sel)) {
			find_0176_jp(xlDocs[f], sel);
		}
		break;
	}
}

//main();

//var sbuf, sel;
//
//	buf = GetCookie("window", 'search_doc');
//if (sbuf) {
//	sel = GetSelectedString();
//	TraceOut("sel:" + sel);
//	if (!sel)
//		sel = GetLineStr(0).replace(/\r?\n/, '');
//
//	SetCookie("window", 'search_doc', sel);
//}
//	//buf = GetCookie("window", 'search_doc');
//TraceOut("search_doc:"+ buf);

function UseDb(sel) {
	sel = sel.split('').join("%");
	
	var conn = new ActiveXObject("ADODB.Connection");
	var rs = new ActiveXObject("ADODB.Recordset");

	var dbPath = "C:\\1002\\tool\\DBmemo1028.accdb";
	var connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;

	conn.Open(connStr);
	rs.Open(
		"SELECT nm, nmjp, doc FROM cols where [nm] like '%target%' or [nmjp] like '%target%'".replace(/target/g, sel), conn);

	var result = "";
	while (!rs.EOF) {
	    s = (rs.Fields(0).Value+","+rs.Fields(1).Value+","+rs.Fields('doc').Value);
		result += (s + "\n");
	    rs.MoveNext();
	}

	if (result) {
		GoLineEnd(0x08);
		Char(13);
		InsText(result);
	}

	rs.Close();
	conn.Close();
}

var sel;
sel = GetSelectedString();
if (!sel)
	sel = GetLineStr(0).replace(/\r?\n/, '');
UseDb(sel);


//// activateExplorer.js
//var shell = new ActiveXObject("Shell.Application");
//var wsh = new ActiveXObject("WScript.Shell");
//
//// loop through all open windows
//var windows = new Enumerator(shell.Windows());
//for (; !windows.atEnd(); windows.moveNext()) {
//    var win = windows.item();
//    try {
//        if (win && win.FullName && win.FullName.toLowerCase().indexOf("explorer.exe") >= 0) {
//            // optional: check folder path
//            // if (win.LocationURL && win.LocationURL.toLowerCase().indexOf("downloads") < 0) continue;
//
//            // bring to front
//            win.Visible = true;
//            win.Focus();
//
//        	var url = win.LocationURL;
//        	var title = url.substring(url.lastIndexOf("/")+1);
//        	TraceOut(win.Document.Folder.Items().Count + "<<");
//
//        	wsh.AppActivate(title);
//            //WScript.Echo("Activated: " + win.LocationURL);
//            break;
//        }
//    } catch (e) {}
//}
