exports.convert = function(source,destination,callback){
	const {spawn}=require('child_process');
	const pLib = require('path');
	const fs = require('fs')
	const extension = pLib.extname(source).substr(1);
	var appType;
	/* TEMPLATE
	if("".indexOf(extension)!=-1){
		
	}
	*/
	if("xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xlt;xla;xlm;xlw;odc;ods;prn;csv;dsn;mdb;mde;accdb;accde;dbc;iqy;dqy;rqy;oqy;cub;atom;atomsvc;dbf;xll;xlb;slk;dif;xlk;bak".split(";").indexOf(extension)!=-1){
		appType = "XL"
	}
	if("pptx;ppt;pptm;ppsx;pps;ppsm;potx;pot;potm;odp;thmx;docx;docm;doc;ppam;ppa".split(";").indexOf(extension)!=-1){
		appType = "PP"
	}
	if("docx;docm;dotx;dotm;doc;odt;docx;docm;doc;dotx;dotm;dotx;dotm;rtf;odt;doc;wpd;doc".split(";").indexOf(extension)!=-1){
		appType = "WD"
	}
	if(fs.existsSync(source)){
		const topdf = spawn(pLib.resolve(__dirname,'mso-pdf.exe'),[appType,source,destination])
		topdf.stdout.on('data',function(data){
			console.log(data.toString())
		})
		topdf.on('close',callback)
	} else {
		callback({error:(new Error("Source file does not exist.")),file:source})
	}
}

