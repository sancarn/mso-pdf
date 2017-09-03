ControlGet, clipboard, List, , Combobox2, Open ahk_class #32770

/*
EXCEL::
	All Files (*.*)
	All Excel Files (*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xlt;*.htm;*.html;*.mht;*.mhtml;*.xml;*.xla;*.xlm;*.xlw;*.odc;*.ods)
	Excel Files (*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xla;*.xlt;*.xlm;*.xlw)
	All Web Pages (*.htm;*.html;*.mht;*.mhtml)
	XML Files (*.xml)
	Text Files (*.prn;*.txt;*.csv)
	All Data Sources (*.odc;*.udl;*.dsn;*.mdb;*.mde;*.accdb;*.accde;*.dbc;*.iqy;*.dqy;*.rqy;*.oqy;*.cub;*.atom;*.atomsvc)
	Access Databases (*.mdb;*.mde;*.accdb;*.accde)
	Query Files (*.iqy;*.dqy;*.oqy;*.rqy)
	dBase Files (*.dbf)
	Microsoft Excel 4.0 Macros (*.xlm;*.xla)
	Microsoft Excel 4.0 Workbooks (*.xlw)
	Worksheets (*.xlsx;*.xlsm;*.xlsb;*.xls)
	Workspaces (*.xlw)
	Templates (*.xltx;*.xltm;*.xlt)
	Add-ins (*.xlam;*.xla;*.xll)
	Toolbars (*.xlb)
	SYLK Files (*.slk)
	Data Interchange Format (*.dif)
	Backup Files (*.xlk;*.bak)
	OpenDocument Spreadsheet (*.ods)

POWERPOINT:
	All Files (*.*)
	All PowerPoint Presentations (*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm;*.potx;*.pot;*.potm;*.odp)
	Presentations and Shows (*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm)
	PowerPoint XML Presentations (*.xml)
	PowerPoint Templates (*.potx;*.pot;*.potm)
	Office Themes (*.thmx)
	All Outlines (*.txt;*.rtf;*.docx;*.docm;*.doc;*.wpd)
	PowerPoint Add-ins (*.ppam;*.ppa)
	OpenDocument Presentations (*.odp)

WORD:
	All Files (*.*)
	All Word Documents (*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.mht;*.mhtml;*.xml;*.odt;*.pdf)
	Word Documents (*.docx)
	Word Macro-Enabled Documents (*.docm)
	XML Files (*.xml)
	Word 97-2003 Documents (*.doc)
	All Web Pages (*.htm;*.html;*.mht;*.mhtml)
	All Word Templates (*.dotx;*.dotm;*.dot)
	Word Templates (*.dotx)
	Word Macro-Enabled Templates (*.dotm)
	Word 97-2003 Templates (*.dot)
	Rich Text Format (*.rtf)
	Text Files (*.txt)
	OpenDocument Text (*.odt)
	PDF Files (*.pdf)
	Recover Text from Any File (*.*)
	WordPerfect 5.x (*.doc)
	WordPerfect 6.x (*.wpd;*.doc)
	
	
The above is then transformed with regexr using
	/.* \(\*.(.*)\)/g --> List("$1;\n")
	/\w*\*;|\*.|\s+/g --> Replace("")

into
xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xlt;htm;html;mht;mhtml;xml;xla;xlm;xlw;odc;ods;xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xla;xlt;xlm;xlw;htm;html;mht;mhtml;xml;prn;txt;csv;odc;udl;dsn;mdb;mde;accdb;accde;dbc;iqy;dqy;rqy;oqy;cub;atom;atomsvc;mdb;mde;accdb;accde;iqy;dqy;oqy;rqy;dbf;xlm;xla;xlw;xlsx;xlsm;xlsb;xls;xlw;xltx;xltm;xlt;xlam;xla;xll;xlb;slk;dif;xlk;bak;ods;pptx;ppt;pptm;ppsx;pps;ppsm;potx;pot;potm;odp;pptx;ppt;pptm;ppsx;pps;ppsm;xml;potx;pot;potm;thmx;txt;rtf;docx;docm;doc;wpd;ppam;ppa;odp;docx;docm;dotx;dotm;doc;dot;htm;html;rtf;mht;mhtml;xml;odt;pdf;docx;docm;xml;doc;htm;html;mht;mhtml;dotx;dotm;dot;dotx;dotm;dot;rtf;txt;odt;pdf;doc;wpd;doc;

Refined lists:

	Excel:
		xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xlt;htm;html;mht;mhtml;xml;xla;xlm;xlw;odc;ods;xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xla;xlt;xlm;xlw;htm;html;mht;mhtml;xml;prn;txt;csv;odc;udl;dsn;mdb;mde;accdb;accde;dbc;iqy;dqy;rqy;oqy;cub;atom;atomsvc;mdb;mde;accdb;accde;iqy;dqy;oqy;rqy;dbf;xlm;xla;xlw;xlsx;xlsm;xlsb;xls;xlw;xltx;xltm;xlt;xlam;xla;xll;xlb;slk;dif;xlk;bak;ods;
	
	PowerPoint:
		pptx;ppt;pptm;ppsx;pps;ppsm;potx;pot;potm;odp;pptx;ppt;pptm;ppsx;pps;ppsm;xml;potx;pot;potm;thmx;txt;rtf;docx;docm;doc;wpd;ppam;ppa;odp;
	
	Word:
		docx;docm;dotx;dotm;doc;dot;htm;html;rtf;mht;mhtml;xml;odt;pdf;docx;docm;xml;doc;htm;html;mht;mhtml;dotx;dotm;dot;dotx;dotm;dot;rtf;txt;odt;pdf;doc;wpd;doc;
		
		
NO DUPLICATES
	Excel:
		xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xlt;htm;html;mht;mhtml;xml;xla;xlm;xlw;odc;ods;prn;txt;csv;udl;dsn;mdb;mde;accdb;accde;dbc;iqy;dqy;rqy;oqy;cub;atom;atomsvc;dbf;xll;xlb;slk;dif;xlk;bak
	PowerPoint:
		pptx;ppt;pptm;ppsx;pps;ppsm;potx;pot;potm;odp;xml;thmx;txt;rtf;docx;docm;doc;wpd;ppam;ppa;
	Word:
		docx;docm;dotx;dotm;doc;dot;htm;html;rtf;mht;mhtml;xml;odt;pdf;txt;wpd;
		
		
	Assumptions:
		You can open html, htm in excel, powerpoint and word however we will presumably open them in iframe
		mhtml, mht format (MIME HTML) may also be possible via iframe? But I haven't seen it used in general...
		txt format will likely just be read and previewed explicitely as text in an input control.
		rtf format can be passed to word but since it is quite a popular format we should probably render it directly as html.
		xml format, as above
		wpd to be opened in word not powerpoint
		prn is a space delimited text format and should be opened with excel
		udl has no preview in excel
		pdf will be opened with pdf.js so no need for them here
		
REFINED LISTS:
	Excel:
		xlsx;xlsm;xlsb;xlam;xltx;xltm;xls;xlt;xla;xlm;xlw;odc;ods;prn;csv;dsn;mdb;mde;accdb;accde;dbc;iqy;dqy;rqy;oqy;cub;atom;atomsvc;dbf;xll;xlb;slk;dif;xlk;bak
	PowerPoint:
		pptx;ppt;pptm;ppsx;pps;ppsm;potx;pot;potm;odp;thmx;docx;docm;doc;ppam;ppa;
	Word:
		docx;docm;dotx;dotm;doc;odt;docx;docm;doc;dotx;dotm;dotx;dotm;rtf;odt;doc;wpd;doc;