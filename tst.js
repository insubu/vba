(function(){
	var cmd = Plugin.GetCommandNo();
	switch(cmd) {
// 	case 1:
// 		insert();
// 		break;
	default:
		tst1();
		break;
	}
})();

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

var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

	function tst1() {
		var conn = new ActiveXObject("ADODB.Connection");
		conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\1002\\tool\\DBmemo1028.accdb');

		var w = Complement.GetCurrentWord();
		//TraceOut("GetCurrentWord:" + w);

		var rs = new ActiveXObject("ADODB.Recordset");
		var sql = "select * from complement where [type] like '%@%'";
		//var rs = conn.Execute(sql);
		rs.Open(sql.replace(/@/g, w.split('').join('%')), conn);

// 			itr(rs, function(x){
// 				TraceOut("..col:" + x.テーブル名);
// 				//Complement.AddList(x.Fields("テーブル名").Value);
// 			});
		//for (var it = new Enumerator(rs.Fields); !it.atEnd(); it.moveNext()) {
		//	TraceOut("col:" + it.item().Name);
		//}
 		while (!rs.EOF) {
 			//TraceOut("cont:"+rs.Fields("content").Value);
 			Complement.AddList(rs.Fields("content").Value);
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

	//var exec = wsh.Exec( "\"" + mshta + "\" \"" + tname + "\"" );
	var exec = wsh.Exec( "C:\\1002\\tool\\sakura\\myf.exe" );
	var input = exec.StdOut.ReadAll();

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
