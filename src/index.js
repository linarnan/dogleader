const MatchColNames = [
	"訂單編號 (單)", //A
	"訂單狀態 (單)", //B
	"收件者姓名 (單)", //C
	"商品名稱 (品)", //D
	"商品選項名稱 (品)", //E
	"數量", //F
	"寄送方式 (單)", //G
	"商品選項貨號" //H
];
const additionalCol = [
	{ i: "I", name: "營業額", f: '' },
	{ i: "J", name: "成本", f: '' },
	{ i: "K", name: "利潤", f: 'I$-J$' },
	{ i: "L", name: "稅", f: '(I$-J$) * 0.05 + I$ * 0.06 * 0.2' },
	{ i: "M", name: "稅後利潤", f: 'K$ - L$' }, //利潤 - 稅
	{ i: "N", name: "百分比", f: 'M$ / I$' }
];

function debug() {
	console.log.apply(console, arguments);
}

function onFileChanged(e) {
	var files = e.target.files,
		f = files[0];
	var reader = new FileReader();
	reader.onload = function (e) {
		var data = new Uint8Array(e.target.result);
		loadXLSXBuffer(data);
	};
	reader.readAsArrayBuffer(f);
}

function loadXLSXBuffer(data) {
	var xls = XLSX.read(data, {
		type: "array"
	});
	console.log(xls);
	var aoa = parseData(xls);
	var wb = XLSX.utils.book_new();
	var newWs = XLSX.utils.aoa_to_sheet(aoa);

	debug("newWs: ", newWs)
	// if (newWs) {
	// 	newWs["L2"].f = `K${i}*0.05 + I${i}*0.06*0.2`;
	// }
	debug(wb);
	var newSheetName = "result";
	wb.SheetNames.push(newSheetName);
	wb.Sheets[newSheetName] = newWs;
	postProcessing(newWs, aoa, additionalCol);

	debug("wb", wb);
	writeXLSX_workbook(wb)
}

function writeXLSX_workbook(wb) {
	var date = dateFormat(new Date());
	XLSX.writeFile(wb, date + ".xlsx");
}


function parseData(xls) {
	//console.log(xls);
	var sheet1 = xls.Workbook.Sheets[0];
	var sheetName = sheet1.name;
	var rows = xls.Sheets[sheetName];
	var range = xls.Sheets[sheetName]["!ref"];
	var matches = range.match(/[A-Z]{1,2}\d*:([A-Z]{1,2})(\d*)/);
	var cells = xls.Strings;
	var numRows = matches[2];
	//console.log(sheet1, range, matches);
	var colNum = colName2Num(matches[1]);
	console.log("Max colNum", colNum);
	var selectCols = [];
	var colNameMap = {};
	for (var i = 1; i <= colNum; i++) {
		if (rows[cellIdx(i, 1)]) {
			var colName = rows[cellIdx(i, 1)].v;
			colNameMap[colName] = i;
		}
		//console.log(colName);
	}

	console.log("Column Name Map", colNameMap);

	for (var i = 0; i < MatchColNames.length; i++) {
		var mIdx = colNameMap[MatchColNames[i]];
		//console.log("mIdex", mIdx, MatchColNames, colName);
		if (mIdx) {
			selectCols.push(mIdx);
		}
	}

	debug(selectCols);
	var result = [MatchColNames];
	var skip = false;
	for (var i = 1; i <= numRows; i++) {
		let r = [];
		skip = false;
		for (var j = 0; j < selectCols.length; j++) {
			let idx = cellIdx(selectCols[j], i);
			console.log("Get column value: ", idx, rows[idx]);
			if (rows[idx]) {
				if (j === 1 && rows[idx].v !== "待出貨") {
					skip = true;
					break;
				}
				r.push(rows[idx].v);
			} else {
				r.push("");
			}
		}
		if (!skip) {
			result.push(r);
		}
	}

	//add additionalCol Columns
	for (var j = 0; j < additionalCol.length; j++) {
		Array.prototype.push.call(MatchColNames, additionalCol[j].name);
	}

	var shouldColumnNum = result[0].length;
	//debug("shouldColumnNum", shouldColumnNum);
	//padding all row to exactly column number
	for (var i = 1; i < result.length; i++) {
		//debug("result[i]", result[i]);
		if (result[i].length < shouldColumnNum) {
			var l = shouldColumnNum - result[i].length;
			for (var c = 0; c < l; c++) {
				result[i].push('');
			}
		}
	}
	debug("result", result);
	return result;
}


function postProcessing(newWs, result, additionalCol) {
	for (var i = 1; i < result.length; i++) {
		var ii = i + 1;

		for (var j = 0; j < additionalCol.length; j++) {
			var colIdx = additionalCol[j].i;
			newWs[`${colIdx}${ii}`].f = additionalCol[j].f.replace(/\$/g, ii);
			newWs[`${colIdx}${ii}`].t = 'n';
		}
	}
}

function dateFormat(_t) {
	var t = _t ? _t : new Date();

	var year = t.getFullYear();
	var month = t.getMonth() + 1;
	var date = t.getDate();
	return `${year}-${month < 10 ? "0" : ""}${month}-${
		date < 10 ? "0" : ""
		}${date}`;
}

const ColIdx = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

function cellIdx(colNum, rowNum) {
	var idx = colNum2Name(colNum) + rowNum;
	//console.log(idx);
	return idx;
}

function colName(num) {
	return ColIdx.charAt(colNum - 1);
}

function colName2Num(col) {
	if (col.length === 2) {
		return (ColIdx.indexOf(col[0]) + 1) * 26 + (ColIdx.indexOf(col[1]) + 1);
	} else {
		return ColIdx.indexOf(col[0]) + 1;
	}
}

function colNum2Name(colNum) {
	if (colNum <= 26) {
		return ColIdx.charAt(colNum - 1);
	} else {
		return (
			ColIdx.charAt(Number.parseInt(colNum / 26) - 1) +
			ColIdx.charAt(Number.parseInt(colNum % 26) - 1)
		);
	}
}

function bootstrap() {
	document
		.getElementById("fileUpload")
		.addEventListener("change", onFileChanged, false);
}
