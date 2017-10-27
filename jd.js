var xlsx = require("node-xlsx");
var fs = require('fs');
var opts = {};
for (var key of process.argv.splice(2)) {
	var keys = key.split('=');
	opts[keys[0]] = keys[1];
}
opts.head = opts.head || 0;

function indexArr (str, arr, begin) {
	var mi = 100000;
	for (var i = 0; i < arr.length; i++) {
		var ix = str.indexOf(arr[i], begin);
		if (ix != -1 && ix < mi) {
			mi = ix;
		}
	}
	if (mi != 100000) {
		return mi;
	} else {
		return -1
	}
}

function matchsSplit (str, index, splits) {
	for (var key in splits) {
		for (var i = 0; i < key.length; i++) {
			if (str.charAt(index + i) != key.charAt(i)) {
				break;
			}
		}
		if (i == key.length) {
			return key;
		}
	}
}

function cmp (line, isHead, a, b) {
	var cta = {};
	var ctb = {};

	var ctas = {};
	var ctbs = {};
	var errs = [];

	if (!a || !b) {errs.push("缺少列");return errs}

	if(typeof(b) == "number") {
		errs.push("译文格式错误， 不应为数字");
		return errs;
	}


	// 检测
	for (var i = 0; i < encodes.length; i++) {
		var cax = a.match(encodes[i].src);
		var cbx = b.match(encodes[i].dest);
		var ca = 0;
		var cb = 0;
		
		if (cax) {
			ca = cax.length;
			a = a.replace(encodes[i].src, '');
		}
		
		if (cbx) {
			cb = cbx.length;
			for (var j = 0; j < ca; j++) {
				b = b.replace(encodes[i].destSingle, '');
			}
		}
		if (ca > cb) {
			console.log(`-------错误：【${encodes[i].srcRaw} => ${encodes[i].destRaw}】转换错误:  ${ca}  :  ${cb}`)
			errs.push(`【${encodes[i].srcRaw} => ${encodes[i].destRaw}】转换错误:  ${ca}  :  ${cb}`);
		}
	}

	var b2 = b.replace(/<sup>/g, '').replace(/<\/sup>/g, '').replace(/<sub>/g, '').replace(/<\/sub>/g, '')


	for (var i = 0; i < b2.length; i++) {
		if (!EUC[b2.charAt(i)]) {
			if (halfCode.indexOf(b2.charAt(i)) == -1) {
				console.log("-------错误：【"+b2.charAt(i)+"】不在EUC范围内(第"+(i+1)+")")
				errs.push("【"+b2.charAt(i)+"】不在EUC范围内(第"+(i+1)+")");
			} else {
				console.log("-------错误：【"+b2.charAt(i)+"】半角符号(第"+(i+1)+")")
				errs.push("【"+b2.charAt(i)+"】半角符号(第"+(i+1)+")");
			}
		}
	}

	if (a.indexOf("<img") != -1 || a.indexOf('<table') != -1 || a.indexOf('<maths') != -1) {
		return errs;
	}

	if (opts.quote && !isHead) {
		var end = b.charAt(b.length-1);
		if (!quotes[end]) {
			if(!['及','或','或者','且','其中','和'].some((word) => {
				return a.endsWith(word);
			} && ["【実施例】","［実施例］","（実施例）"].indexOf(b) !== -1)) {
				errs.push("结尾必须是标点");
				console.log("-------"+line+"错误：结尾必须是标点" + end + "(第"+(i+1)+")");
			}
		}
	}

	a = a.replace(/“/g,"\"")
		.replace(/”/g,"\"")
		.replace(/’/g,"\'")
		.replace(/‘/g,"\'")
		.replace(/–/g, "-")
		.replace(/—/g, "-")
		.replace(/-+/g, "-")
		.replace(/∶/g, "")
		.replace(/应该/g, "");

	for (var i = 0; i < codeb.length; i++) {
		a = a.replace(new RegExp(codeb[i],'g'), codea[i]);
	}

	b = b.replace(/“/g,"\"")
		.replace(/”/g,"\"")
		.replace(/’/g,"\'")
		.replace(/‘/g,"\'")
		.replace(/应该/g, "");

	for (var key in splits) {
		cta[key] = 0;
		ctb[key] = 0;
	}

	for (var i = ng.length - 1; i >= 0; i--) {
		if (b.indexOf(ng[i]) !== -1) {
			console.log("-------错误：【"+ng[i]+"】NG用语")
			errs.push("【"+ng[i]+"】NG用语");
		}
	};

	/*
	for (var key in oneway) {
		ctas[key] = 0;
		ctbs[key] = 0;
	}
	for (var key in oneway) {
		var ra = new RegExp(key, 'g');
		var rb = new RegExp(oneway[key], 'g');
		var ca = a.match(ra);
		var cb = b.match(rb);
		if (ca) {
			ca = ca.length;
			a = a.replace(ra, '');
		} else {ca = 0}
		if (cb) {
			cb = cb.length;
			for (var i = 0; i < ca; i++) {
				b = b.replace(new RegExp(oneway[key]), '');
			}
		} else {cb = 0}
		if (ca > cb) {
			console.log("-------错误：【"+key+"】转换错误:" + ca + ":" + cb)
			errs.push("【"+key+"】转换错误:" + ca + ":" + cb);
		}

	}
	*/
	//console.log(b);

	for (var i = 0; i < a.length; i++) {
		var m = matchsSplit(a, i, splits);
		if (m) {
			//coma.push(m);
			cta[m]++;
		}
	}
	for (var i = 0; i < b.length; i++) {
		var m = matchsSplit(b, i, splitRev);
		//comb.push(m);
		if (m) {
			ctb[splitRev[m]]++;
		}
	}

	var match = true;
	for (var key in splits) {
		if (cta[key] != ctb[key]) {
			match = false;
			console.log("-------错误：【"+key+"】数量不匹配:" + cta[key] + ":" + ctb[key])
			errs.push("【"+key+"】数量不匹配:" + cta[key] + ":" + ctb[key]);
		}
	}
	if(b.indexOf('\n') !== -1) {
		console.log("-------错误：存在换行")
		errs.push("存在换行");
	}
	// if (match) {
	// 	var startIndex = 0;
	// 	for (var i = 0; i < coma.length; i++) {
	// 		var newstartIndex = indexArr(b, splits[coma[i]], startIndex);
	// 		if (newstartIndex == -1) {
	// 			console.log("-------错误：顺序混乱:" + startIndex + ":" + coma[i]);
	// 			errs.push("顺序混乱:" + startIndex + ":" + coma[i]);
	// 			return errs;
	// 		}
	// 		startIndex = newstartIndex;
	// 	}
	// }
	return errs;

}

function xls2txt (nm) {

	var list = xlsx.parse("./excel/" + nm);
	var data = list[0].data;
	var res = [];

	for (var i = 0; i < data.length; i++) {
		//test
		if (i < opts.head) {continue}
		if (data[i][0] != undefined) {
			var isHead = false;
			//console.log(data[i][7])
			if (data[i][7] == "0") {
				isHead = true;
			}
			//console.log(isHead)
			var d = cmp(i, isHead, data[i][opts.l1], data[i][opts.l2]);
			if (data[i].length != opts.l3) {
				d.push("列数量错误")
			}
			if (d.length) {
				data[i][opts.l3] = ("问题：\n" + d.join(';\n'));
			} else {
				data[i][opts.l3] = (" ");
			}
			res.push(data[i].join('\t'));
		}
	}

	var buffer = xlsx.build([{name: "sheet1", data: data}]);
		
	fs.writeFile('./检测/' + nm.replace('.txt', '.xlsx'), buffer, function (err) {
		if (err) throw err;
		console.log('详见:/检测/' + nm);
	});

}

var EUC = {};

var splits = {
	"\"": ["\""],
	"'": ["'"],
	"所述": ["前記"],
	"该": ["該"],
	"<sup>": ['<sup>'],
	"</sup>": ['</sup>'],
	"<sub>": ['<sub>'],
	"</sub>": ['</sub>'],
	"上述": ["上記"],
}

var codea = ["·", "0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",",","、",";",":","%","(",")","+","-",".","<","=",">","≤","≥","[","]","{","}","/","|","~","_","《","》","≈"]
var codeb = ["・", "０","１","２","３","４","５","６","７","８","９","ａ","ｂ","ｃ","ｄ","ｅ","ｆ","ｇ","ｈ","ｉ","ｊ","ｋ","ｌ","ｍ","ｎ","ｏ","ｐ","ｑ","ｒ","ｓ","ｔ","ｕ","ｖ","ｗ","ｘ","ｙ","ｚ","Ａ","Ｂ","Ｃ","Ｄ","Ｅ","Ｆ","Ｇ","Ｈ","Ｉ","Ｊ","Ｋ","Ｌ","Ｍ","Ｎ","Ｏ","Ｐ","Ｑ","Ｒ","Ｓ","Ｔ","Ｕ","Ｖ","Ｗ","Ｘ","Ｙ","Ｚ","，","、","；","：","％","（","）","＋","－","．","＜","＝","＞","≦","≧","［","］","｛","｝","／","｜","～","＿","《","》","≒"]
for (var i = 0; i < codea.length; i++) {
	splits[codea[i]] = [codea[i], codeb[i]];
}
var halfCode = ["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",",","、",";",":","%","'","\"","(",")","+","-",".","<","=",">","≤","≥","[","]","{","}","/","|","~","_","《","》","≈"];
var splitRev = {};
for (var key in splits) {
	for (var i = 0; i < splits[key].length; i++) {
		splitRev[splits[key][i]] = key;
	}
}

var ng = [];
var quotes = {};
var encodes = [];

fs.readdir("./excel", function(err, files){
    if (err) {
        console.log(err);
        return;
    }
    ng = fs.readFileSync('./LIB/ng.data') + '';
    ng = ng.split('\n');

    var quotedata = fs.readFileSync('./LIB/quotes.data') + '';
    quotedata = quotedata.split('\n');
    for (var key in quotedata) {
    	quotes[quotedata[key]] = true;
    }

    var encodeData = fs.readFileSync('./LIB/encode.data') + '';
    encodeData = encodeData.split('\n');
    for (var key in encodeData) {
    	var mean = encodeData[key].split('#')[0].trim();
    	if (mean) {
	    	var pair = mean.split('<<<');
	    	var src = pair[1].trim();
	    	if (src.charAt(0) == '"' && src.charAt(src.length-1) == '"') {
	    		src = src.substr(1, src.length-2);
	    	}
	    	var dest = pair[0].trim();
	    	if (dest.charAt(0) == '"' && dest.charAt(dest.length-1) == '"') {
	    		dest = dest.substr(1, dest.length-2);
	    	}
	    	encodes.push({
	    		srcRaw: src,
	    		destRaw: dest,
	    		src: new RegExp(src, 'g'),
	    		dest: new RegExp(dest, 'g'),
	    		destSingle: new RegExp(dest, '')
	    	})
	    }
    }
    
    fs.readFile('./LIB/EUC.csv', function (err, data) {
    	var EUCs = (data+"").replace(/\n,/g,'');
    	for (var i = 0; i < EUCs.length; i++) {
    		EUC[EUCs.charAt(i)] = true;
    	}
    	for (var i = 0; i < files.length; i++) {
	    	if (files[i].match(/\.xlsx$/)) {
	    		console.log(files[i])
		    	xls2txt(files[i]);
		    }
	    }
	});
});