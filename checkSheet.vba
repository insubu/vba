// var ttt=[1, 3];
// TraceOut(Object.prototype.toString.call(ttt)=="[object Array]");
// TraceOut(true);
// TraceOut(false); 
// if (ttt instanceof String) {
// 	TraceOut("yes, ttt instanceof Array")
// }
var key_defvalues = {
	"cntBracket" : 120
};

var shell = new ActiveXObject("Shell.Application");
var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var sakura = "C:\\1002\\apps\\sakura\\sakura.exe";
var xls = GetObject("","Excel.Application");

function openTDC_Url_Bug_1653689(cmd) {
	var usrSelCmd = GetUserSelectedFuncName(cmd);
	TraceOut("...usrCmd:" + usrSelCmd);
	switch(usrSelCmd) {
	case "url-Repos": urlBug_1653689 = 'https://dev.azure.com/earthPJ/Migration/_git/fb-tr-docs-sql-tr?path=/SQL-B/%E3%83%9E%E3%82%B9%E3%82%BF%E3%83%BC';
			break;
	case "url-TDC_Bug_1653689": urlBug_1653689 = 'https://fujifilm0-my.sharepoint.com/personal/fxyaf1e_000_fujifilm_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Ffxyaf1e%5F000%5Ffujifilm%5Fcom%2FDocuments%2F%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E5%85%B1%E6%9C%89%2F09%5F%E6%88%90%E6%9E%9C%E7%89%A9%2F01%5F%E3%83%9E%E3%82%B9%E3%82%BF%E9%A0%98%E5%9F%9F%2F03%5F%E9%9A%9C%E5%AE%B3%E8%AA%BF%E6%9F%BB%E3%83%BB%E9%9A%9C%E5%AE%B3%E5%AF%BE%E5%BF%9C%2FBUG%5F1653689&sortField=Modified&isAscending=false&csf=1&web=1&e=qd7iMR&FolderCTID=0x0120004CDDA9B56334B24B9687859985A79951&OR=Teams%2DHL&CT=1727763409109&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNDA4MTcwMDQyMSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D';
            break;
	case "url-ADO_Bug_1653689": urlBug_1653689 = 'https://dev.azure.com/earthPJ/Migration/_workitems/edit/1653689';
			break;
	case MENU_S05A/*xls-Evi_Bug_1653689*/:
			urlBug_1653689 = 'C:\\1002\\src\\BUG_0331\\BUG_1653689_単体テスト?追加対応.xlsx';
			break;
	case MENU_S05B/*teams*/:
			urlBug_1653689 = "https://teams.microsoft.com/l/message/19:meeting_MjY1OGZiMTgtZjY2My00NmU4LTlhYTctM2Q4ZTNlMjYxMGY1@thread.v2/1743659741167?context=%7B%22contextType%22%3A%22chat%22%7D";
			break;
	case "dir-Bug_1653689": urlBug_1653689 = 'C:\\1002\\src\\BUG_0331\\web\\edit\\0403_TOIDが必要';
			break;
	case "view-2-12-Run.sql" :
			urlBug_1653689 = "C:\\1002\\src\\BUG_0331\\UT\\0026-1_2-12_RUN_UPDATE.sql";
			break;
	case "view-2-12-upd1.sql" :
			urlBug_1653689 = "C:\\1002\\src\\BUG_0331\\web\\edit\\0403_TOIDが必要\\0026-1_2-12_UPDATE1.sql"
			break;
	case "view-2-12-upd2.sql" :
			urlBug_1653689 = "C:\\1002\\src\\BUG_0331\\web\\edit\\0403_TOIDが必要\\0026-1_2-12_UPDATE2.sql"
			break;
	case "[EE]view-2-12-upd3.sql" :
			urlBug_1653689 = "C:\\1002\\src\\BUG_0331\\web\\edit\\0403_TOIDが必要\\0026-1_2-12_UPDATE3.sql"
			break;
	default:
			return false;
	}
	if ((/^http/i).test(urlBug_1653689))
		shell.ShellExecute(urlBug_1653689);
	else {
		if ((/view-/i).test(usrSelCmd))
			wsh.Run(sakura + ' "' + urlBug_1653689 + '"');
		else if ((/xls-/i).test(usrSelCmd)) {
			xls.Workbooks.Open(urlBug_1653689);
			xls.WindowState = 1;
		}
		else
			shell.Explore(urlBug_1653689);
	}
	return true;
}

var MENU_S01 = "url-Repos"                ;
var MENU_S02 = "[S]Bug_1653689-0026-1"    ;
var MENU_S03 =   "url-TDC_Bug_1653689"    ;
var MENU_S04 =   "url-ADO_Bug_1653689"    ;
var MENU_S05 =   "dir-Bug_1653689"        ;
var MENU_S050=   "[S]doc-Bug_1653689"     ;
var MENU_S05A=   "xls-Evi_Bug_1653689"    ;
var MENU_S05B=   "[E]teams-DELETEが必要" ;
var MENU_S06 =   "[S]view-Source-sql"     ;
var MENU_S07 =     "view-2-12-Run.sql"    ;
var MENU_S08 =     "view-2-12-upd1.sql"   ;
var MENU_S09 =     "view-2-12-upd2.sql"   ;
var MENU_S10 =   "[EE]view-2-12-upd3.sql" ;

var menuId = [];
var menuArr = 	[
	 MENU_S01//"url-Repos"  
	,MENU_S02//"[S]Bug_1653689-0026-1" 
	,MENU_S03//   "url-TDC_Bug_1653689"
	,MENU_S04//   "url-ADO_Bug_1653689" 
	,MENU_S05//   "dir-Bug_1653689" 
    ,MENU_S050
	,MENU_S05A//  "xls-Evi_Bug_1653689"    ;
    ,MENU_S05B
	,MENU_S06//   "[S]view-Source-sql" 
	,MENU_S07//     "view-2-12-Run.sql" 
	,MENU_S08//     "view-2-12-upd1.sql" 
	,MENU_S09//     "view-2-12-upd2.sql" 
	,MENU_S10//   "[EE]view-2-12-upd3.sql" 
	];

for (var i in menuArr) {
	var m = menuArr[i];
	if ((/\[S\]/).test(m)) continue;
	//TraceOut("arr..." + menuArr[i]);
	menuId.push(menuArr[i]);
}

function GetUserSelectedFuncName(cmdNo) {
	return menuId[cmdNo-1];
}

var cmd = CreateMenu( 0, menuArr.join(",") );
TraceOut("userSelCmd:" + GetUserSelectedFuncName(cmd));

	var done = openTDC_Url_Bug_1653689(cmd);
// if (done == false) {
// 	switch(cmd) {
//  	case 1:
//  		openTDC_Url_Bug_1653689();
//  		break;
// 	default:
// 		TraceOut("cmd:" + cmd)//tst1();
// 		break;
// 	}
// }


function dbAccess() {
	var conn = new ActiveXObject("ADODB.Connection");
	conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\1002\\tool\\DBmemo1028.accdb');

	var sql = "select * from nameRule";
	var rs = conn.Execute(sql);

	// 	itr(rs, function(x){
	// 		TraceOut("..col:" + x.Fields("name").Value);
	// 		//Complement.AddList(x.Name);
	// 	});
	//for (var it = new Enumerator(rs.Fields); !it.atEnd(); it.moveNext()) {
	//	TraceOut("col:" + it.item().Name);
	//}
	while (!rs.EOF) {
	 	//TraceOut("caption:"+rs.Fields("name").Value);
		TraceOut("caption:"+rs["name"]);
		rs.MoveNext();
	}

	rs.Close();
	conn.Close();
}

function itr(sset, fun) {
	for (var it = new Enumerator(sset); !it.atEnd(); it.moveNext()) {
		fun(it.item());
	}
}
//ShowSettingsDialog();

// buf = ExpandParameter("$E");
// TraceOut("E"+(new Date()).getSeconds()+":"+buf, 1);
// buf = ExpandParameter("$e");
// TraceOut("e"+(new Date()).getSeconds()+":"+buf, 1);

function ShowSettingsDialog() {
//   var tname = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName() + ".hta");
//   var xmlfile = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName() + ".xml");
	var tname = ("C:\\1002\\tool\\sakura\\tst.hta");
	var mshta = wsh.ExpandEnvironmentStrings("%SystemRoot%\\System32\\mshta.exe");

	var exec = wsh.Exec( "\"" + mshta + "\" \"" + tname + "\"" );

	// 	var exec = wsh.Exec( "C:\\1002\\tool\\sakura\\myf.exe" );
// 	var input = exec.StdOut.ReadAll();

	TraceOut(input);
}

function run(sh, cmd) {
  sh.Run(cmd, 1, true);
}

// r = /\d+/g;
// s = "this 2 is a 3, last is 4";
// // while (m = r.exec(s)) {
// // 	TraceOut("$t"+(new Date()).getSeconds()+":"+ (parseInt(m[0])+20) + m.index, 1);
// // }
// 	var sr = s.replace(r, function(mt) {
// 		return "<"+(parseInt(mt)+30) + ">";
// 	})
// 	TraceOut("$t"+(new Date()).getSeconds()+":"+sr, 1);

// var wsh = new ActiveXObject("WScript.Shell");
// 
// var xl = GetObject("","Excel.Application");
// var cnt = xl.CustomListCount;
// 
// for (var i=1; i<= xl.CustomListCount; i++) {
// var result='';
// var pw=". C:\\1002\\tool\\sakura\\getCustList.ps1 " + i;
// var cmd="powershell -NoProfile -WindowStyle Hidden -Command \"" + pw + "\"";
//  	var exe = wsh.Exec(cmd);
// 	while (!exe.StdOut.AtEndOfStream) {
// 	  result += exe.StdOut.ReadLine()+"@@";
// 	  //break;
// 	}
// 	TraceOut("caption:"+result);
// }

// xl = GetObject("","Excel.Application");
// var w;
// // 	TraceOut(xl.GetCustomListContents(1).length);
// 	var en=xl.GetCustomListContents(1);
// TraceOut(en[1]);

// 	for (var i=1; i<=xl.CustomListCount; i++) {
// 		w = xl.Application.GetCustomListContents(i);
// 		if (w)
// 			TraceOut("$t"+(new Date()).getSeconds()+":"+(typeof w), 1);
// 	}

//   var xl, rng, sel, bkSel, result='';
//   
// 	sel = GetLineStr(0);
// 	if (IsTextSelected) {
// 		sel=GetSelectedString();
// 	}
// 	bkSel = sel;
// 
//     xl = GetObject("","Excel.Application");
//   	rng = xl.WorkBooks("memo.xlsm").WorkSheets("work").Range("sakura連携");
// for (var i=1; i<6; i++) {
// 	var en=new Enumerator(rng);
// 	if (!rng(1,1).Offset(0, i).Value)
// 		break;
// TraceOut("-->"+rng(1,1).Offset(0, i).Value);
// 	sel = bkSel;
// //	en.moveFirst();
// 	while(!en.atEnd()) {
// 		var p = en.item();
// 		if (!p.Value)
// 			break;
// 		TraceOut("p:"+p.value);
// 		if (p.Offset(0, i).Value)
// 			sel = sel.replace(new RegExp(p.Value, "mg"), p.Offset(0, i).Value)
// 		en.moveNext();
// 	}
// 	result += sel;
// }
// TraceOut("--**"+result);
  

// NextWindow();
// WindowList();

//exportDictToXMLFile('keyVal.xml', key_defvalues);

function exportDictToXMLFile(filepath, dict) {
	function appendXmlChild(doc, root, name, val){
	    var el = doc.createElement("property");
	    var cdata = doc.createCDATASection(val);
	    el.appendChild(cdata);
	    el.setAttribute("name", key);
	    el.setAttribute("type", typeof val);
	    root.appendChild(el);
	}

  var doc = new ActiveXObject("MSXML2.DOMDocument");
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var ts = fso.OpenTextFile(filepath, 2, true, -1);
  var root = doc.createElement("properties");
    for (var key in dict) {
		appendXmlChild(doc, root, key, dict[key])
	}
  doc.appendChild(root);
  ts.Write(doc.xml);
  ts.Close();
}
