# MSO-PDF

This is likely extremely poorly written, but as is life

This utility converts word, excel and powerpoint documents to pdf using the Excel/Word/PowerPoint interop libraries.

## Usage

```js
	const toPdf = require('mso-pdf')
	var source 		= "C:\\Users\\sancarn\\Documents\\TestToPDF.xlsx"
	var destination	= "C:\\Users\\sancarn\\Documents\\TestToPDF.pdf"
	toPdf.convert(source,destination,function(errors){
		if(errors) console.log(errors)
		console.log(destination + " converted")
	})
	
```
