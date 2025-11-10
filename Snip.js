
var WSC_PATH = "C:\\1002\\tool\\sakura\\com.wsf";
var com=GetObject("Script:" + WSC_PATH);
com.add(String, Array);

//■function
function EditorStart() {
	TraceOut("EditorStart: <<< handler...");
}

function DocumentOpen() {
	TraceOut("DocumentOpen : <<< handler...");
}

function insertNewLine() {
 	GoLineEnd(0x08);
	Char(13);
}
function pasteBlock() {
	if (IsTextSelected) {
		var sel=GetSelectedString();
		var lineTo = parseInt(GetSelectLineTo());
		Jump(LayoutToLogicLineNum(lineTo));
		InsText(sel);
	} else {
		var sel=GetLineStr(0);
		GoLineEnd(0x08);
		InsText("\r\n"+sel);
	}
}
function goLinePos1() {
	Jump(ExpandParameter("$y"));
}
function blockCmt() {
	var fn=ExpandParameter("$f");
	fn=fn.replace(/[^\.]+\./, '').toLowerCase();

	var cmt = "#";
	if (fn in ftype) {
		cmt = ftype[fn];
	}

	var line = GetLineStr(GetSelectLineFrom);
	//TraceOut("$t:line:"+line, 1);
	var isCmted = (new RegExp("^"+cmt)).test(line);

	if (IsTextSelected == 0) {
		//TraceOut("$t:line cmt.", 1);
		goLinePos1();
		if (isCmted){
			Editor.Replace("^"+cmt+"[ \\t]?", "", 4);
        }
		else
			InsText(cmt+" ");
	}
	else {
		if (GetSelectLineFrom == GetSelectLineTo) {
			goLinePos1();
			if (isCmted)
				Editor.Replace("^"+cmt+"[ \\t]?", "", 4);
			else
				InsText(cmt+" ");
		} else {
			var s = GetSelectLineFrom;
			var e = GetSelectLineTo;
			var eC = GetSelectColmTo;
			
			CancelMode();
			Jump(parseInt(s));
			BeginSelect();
			Jump(parseInt(e)+(eC>1?1:0));
			if (isCmted)
				ReplaceAll("^"+cmt+"[ \\t]?", "", 132);
			else
				ReplaceAll('^', cmt+" ", 132);	// すべて置換
			SearchClearMark();
		}
	}
	ReDraw(0);	// 再描画
}

function bracketSel() {
	CancelMode();
	SearchPrev('{|\\[|\\(|"', 4);	// 前を検索
	SearchClearMark();
	Right();
	BeginSelect();
	BracketPair();
	var s = GetSelectedString();
	if (s) {
		Copy();
	}
	SearchClearMark();
	ReDraw(0);	// 再描画
}

function blockSel() {
	CancelMode();
	var y=ExpandParameter("$y");
	SearchPrev('^\\s*$', 4);
	if (y==ExpandParameter("$y")) 
		GoFileTop()
	else
		Down();
	
	y=ExpandParameter("$y");
	BeginSelect( );
	SearchNext('^\\s*$', 4);
	if (y==ExpandParameter("$y")) 
		GoFileEnd()

	SearchClearMark();
}

function insert() {
	MoveHistSet();

	if (IsTextSelected() == 0) {
		WordLeft_Sel(0);  //@@↓1
		WordLeft_Sel(0);  //↓@@1
		var s = GetSelectedString();
		if (/\s/.test(s)) {
			WordRight_Sel(0);
		}
	}
	var word = GetSelectedString();
	if (word == null || word.trim() == '') return;
	//TraceOut("snip>>>" + word);
	
	if (!(word in snip)) {
		loadKeyVal(word);
		if (!(word in snip)) {
			return;
		}
	}
	
	var snipBody = snip[word];	
	if (/@@\d/.test(word)) {
		//TraceOut("$t"+(new Date()).getSeconds()+":"+word, 1);
		var line = GetLineStr(0).replace(/@@\d|\r\n|\r|\n/g, '');
		var tk = line.split(',');
		for (var i=0; i<tk.length; i++) {
			var reg = new RegExp('@'+i+'@', 'g')
			snipBody = snipBody.replace(reg,
				tk[i]?
					tk[i].replace(/・/g, "\n　・"):
					"");
		}
		var now = new Date();
		var mi = now.getMinutes();
		var hh = now.getHours();
		if (mi < 30) {
			mi = 0;
		} else {
			if (mi >= 57) {
				mi = 0;
				hh += 1;
			} else {
				mi = 30;
			}
		}
		var time = String(hh).padStart() + ':'
			+ String(mi).padStart();
		var log="C:\\1002\\tool\\sakura\\wktime.log";
		if (/@time@/.test(snipBody))
			com.appendFile(fso, log, "wk:"+ (new Date())+":"+time);
		snipBody = snipBody.replace(/@time@/g, time);
		snipBody = snipBody + "\r\n"; 
		Down();
		goLinePos1();
	}
	
	InsText(snipBody);
	MoveHistPrev();
	
	if (/\$\$/.test(snipBody)) {
		SearchNext("$$", 0);
		DeleteBack(0);
	}

}

function loadKeyVal(word) {
	buf = GetCookie("window", word.replace(/@/g, '_'));
// 	TraceOut("getcookie:"+(new Date()).getSeconds()+":"+(buf?buf:"<not found>"), 1);
	if (buf)
		snip[word] = buf;
	else {
		buf = com.importKeySettingFromXMLFile(path,word);
		if (buf) {
			snip[word] = buf;
			SetCookie("window", word.replace(/@/g, '_'), snip[word]);
		}
	}
	return buf;
}

function explorerSrc本番() {
	var wins = shellApp.Windows();
	env("MST_EXPL") = "";
	for (var i=0; i<wins.Count; i++) {
		var n = wins.Item(i);
		if (/マスター/.test(n.LocationName)) {
			if (!env("MST_EXPL")) {
				env("MST_EXPL") = n.LocationName;
			}
			TraceOut("    >>"+n.LocationName);
		}
		if (/差分_/.test(n.LocationName)) {
			env("MST_EXPL") = n.LocationName;
			//TraceOut("    >>"+env("MST_EXPL"));
			break;
		}
	}
	if (!env("MST_EXPL")) {
		shellApp.Explore("C:\\1002\\本番\\マスター");
		env("MST_EXPL") = 'マスター';
	} else {                 
		wsh.AppActivate(env("MST_EXPL"));
		Sleep(100);
		wsh.SendKeys("% ");
		wsh.SendKeys("r");
	}
}

function NumIns() {
	r = /\d+/gm;
	s = GetSelectedString;
	if (s) {
		var sr = s.replace(r, function(mt) {
			return (parseInt(mt)+1);
		});
		if (IsTextSelected == 2)
			InsBoxText(sr);
		else
			InsText(sr);
	} else {
	    s = GetLineStr(0);
		var sr = s.replace(r, function(mt) {
			return (parseInt(mt)+1);
		});
		line = ExpandParameter("$y");
		Jump(parseInt(line)+1);
		InsText(sr);
		Up()
	}
}

function pasteBlockWithXlsRepl() {
	var xl, rng, sel, bkSel, line, result='';

	line = ExpandParameter("$y");
	sel = GetLineStr(0);
	if (IsTextSelected) {
		sel=GetSelectedString();
 		line= GetSelectLineTo();
	}
	bkSel = sel;

	xl = GetObject("","Excel.Application");
	rng = xl.WorkBooks("memo.xlsm").WorkSheets("work").Range("sakura連携");
	for (var i=1; i<16; i++) {
		var en=new Enumerator(rng);
		if (!rng(1,1).Offset(0, i).Value)
			break;
		sel = bkSel;
		//	en.moveFirst();
		while(!en.atEnd()) {
			var p = en.item();
			if (!p.Value)
				break;
			if (p.Offset(0, i).Value)
				sel = sel.replace(new RegExp(p.Value, "mg"), p.Offset(0, i).Value)
			en.moveNext();
		}
		result += sel;
	}
	
	Jump(line);
	InsText(result);
}

//■main
var path="C:\\1002\\apps\\sakura\\plugins\\snip\\tst.xml";
var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var stream = new ActiveXObject("ADODB.Stream");
var shellApp = new ActiveXObject("Shell.Application");
var env=wsh.Environment("User");

var buf = "";
var snip = {
	"\\t": 'TraceOut("$t"+(new Date()).getSeconds()+":"+$$buf, 1);'
};
var ftype = {
	"txt":"--",
	"cs":"//",
	"js":"//",
	"mac":"//",
	"sql":"--",
	"vbs":"'"
};

(function(){
	var cmd = Plugin.GetCommandNo();
	switch(cmd) {
	case 1:
		insert();
		break;
	case 2:
		blockSel();
		break;
	case 3:
		blockCmt();
		break;
	case 4:
		bracketSel();
		break; 
	case 5:
		explorerSrc本番();
		break;
	case 6:
		insertNewLine();
		break;
	case 7:
		pasteBlock();
		break;
	case 8:
		pasteBlockWithXlsRepl();
		break;		
    //case 9: //used in tst.js?

	case 10: //選択数字＋１
		NumIns();
		break;		
	}
})();
