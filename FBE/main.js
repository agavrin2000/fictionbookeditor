var IDOK = 1;
var IDCANCEL = 2;
var IDABORT = 3;
var IDRETRY = 4;
var IDIGNORE = 5;
var IDYES = 6;
var IDNO = 7;

// NodeTypes constants
var ELEMENT_NODE = 1;
var ATTRIBUTE_NODE = 2;
var TEXT_NODE = 3;
var CDATA_SECTION_NODE = 4;
var ENTITY_REFERENCE_NODE = 5;
var ENTITY_NODE = 6;
var PROCESSING_INSTRUCTION_NODE = 7;
var COMMENT_NODE = 8;
var DOCUMENT_NODE = 9;
var DOCUMENT_TYPE_NODE = 10;
var DOCUMENT_FRAGMENT_NODE = 11;
var NOTATION_NODE = 12;

// Image prefix
var imgPrefix = "fbw-internal:";

window.onerror = errorHandler; // document.lvl=0;
//var ImagesInfo = new Array();

// Internalization functions
var curInterfaceLang; // Current interface language
APITranslation = [{
	eng: "Error at line {dvalue1}:\n {dvalue2} ",
	rus: "Ошибка в строке {dvalue1}:\n {dvalue2} ",
	ukr: "Error at line {dvalue1}:\n {dvalue2} "
},
{
	eng: '"{dvalue1}" loading error:\n\n{dvalue2}\nLine: {dvalue3} , Col: {dvalue4}',
	rus: 'Ошибка при загрузке файла:"{dvalue1}":\n\n{dvalue2}\nСтрока: {dvalue3} , Поз: {dvalue4}',
	ukr: '"{dvalue1}" loading error:\n\n{dvalue2}\nLine: {dvalue3} , Col: {dvalue4}'
},
{
	eng: 'insert image',
	rus: 'вставка рисунка',
	ukr: 'insert image'
},
{
	eng: 'make epigraph',
	rus: 'создать эпиграф',
	ukr: 'make epigraph'
},
{
	eng: 'make annotation',
	rus: 'создать аннотацию',
	ukr: 'make annotation'
},
{
	eng: 'insert inline image',
	rus: 'вставка рисунка в текст',
	ukr: 'insert inline image'
},
{
	eng: 'make text-author',
	rus: 'создать слова автора',
	ukr: 'make text-author'
},
{
	eng: 'Error loading file:"{dvalue1}":\nCan\'t prepare document for Body mode.',
	rus: 'Ошибка при загрузке файла:"{dvalue1}":\nНевозможно подготовить BODY',
	ukr: 'Error loading file:"{dvalue1}":\nCan\'t prepare document for Body mode.'
},
{
	eng: 'Notes',
	rus: 'Примечания',
	ukr: 'Notes'
},
{
	eng: 'Cite',
	rus: 'Цитата',
	ukr: 'Cite'
},
{
	eng: 'Stanza',
	rus: 'Строфа',
	ukr: 'Stanza'
},
{
	eng: 'Title',
	rus: 'Заголовок',
	ukr: 'Title'
},
{
	eng: 'Annotation',
	rus: 'Аннотация',
	ukr: 'Annotation'
},
{
	eng: 'Epigraph',
	rus: 'Эпиграф',
	ukr: 'Epigraph'
},
{
	eng: 'Author Text',
	rus: 'Текст автора',
	ukr: 'Author Text'
},
{
	eng: 'Subtitle',
	rus: 'Подзаголовок',
	ukr: 'Subtitle'
}
];

//======================================
// Public API
//======================================

//--------------------------------------
// Gets a binary object for the memory protocol to display book images in the browser
function FillCoverList() {
	return FillLists();
}

function OnBinaryIdChange() {
	FillLists();
}

// Don't forget to call FillCoverList() when you have finished adding objects!
///<summary>Add new binary (picture). 
///Don't forget to call FillCoverList() when you have finished adding objects!</summary>
///<param name='fullpath'>Filepath</param>
///<param name='id'>Proposed picture ID</param>
///<param name='type'>Type of binary (image)</param>
///<param name='data'>Binary data</param>
///<returns>Real ID of picture</returns>
function apiAddBinary(fullpath, id, type, data) {
	// check if this ID already exists and generate new ID, keeping file extension (good for image saving)
	var idx = 0;
	var patt = /(?:\.([^.]+))?$/;
	var ext = patt.exec(id)[1];
	curid = id.replace(/\.[^/.]+$/i, "");
	var idx = $("#binobj DIV").filter(function () { return $("#id", this).val().search(curid) != -1; }).length;
	if (idx > 0)
		curid = id + "_" + (idx - 1);
	curid += "." + ext;
	var div = document.createElement("DIV");

	div.innerHTML = '<button id="del" ' +
		'onmouseover="HighlightBorder(this, true);" ' +
		'onmouseout="HighlightBorder(this, false);" ' +
		'onclick="Remove(this.parentNode);" unselectable="on">&#x72;</button>';
	if (type.search("image") != -1) {
		imghref = "fbw-internal:#" + curid;
		div.innerHTML += '<button id="show"' +
			'onmouseover="ShowImageOfBin(this.parentNode,\'prev\');"' +
			'onmouseout="HidePrevImage();"' +
			'onclick="ShowImageOfBin(this.parentNode,\'full\');"' +
			'unselectable="on">&#x4e;</button>';
		div.innerHTML += '<button id="save" ' +
			'onclick="SaveImage(this.parentNode);" ' +
			'unselectable="on">&#xcd;</button>';
	}

	div.innerHTML += '<label unselectable="on">ID:</label><input type="text" maxlength="256" id="id" style="width:20em;"/><label unselectable="on">Type:</label><input type="text" style="width:8em;" maxlength="256" id="type" value=""/>';
	div.innerHTML += '<label unselectable="on">Size:</label><input type="text" disabled id="size" style="width:5em;" value="' + window.external.GetBinarySize(data) + '"/>';

	if (type.search("image") != -1) {
		var Dims;

		if (fullpath != "")
			Dims = window.external.GetImageDimsByPath(fullpath);
		else
			Dims = window.external.GetImageDimsByData(data);

		if (Dims != "") {
			var imgWidth = Dims.substring(0, Dims.search("x"));
			var imgHeight = Dims.substring(Dims.search("x") + 1, Dims.length);
			div.innerHTML += '<label unselectable="on">w*h:</label><input type="text" disabled id="dims" style="width:6em;" value="' + imgWidth + 'x' + imgHeight + '">';
		}
	}

	$(div).find("#id").val(curid).attr("oldId", curid).change(function () {
		OnBinaryIdChange();
	});
	$(div).find("#type").val(type);
	$(div).data("base64data", data);

	$("#binobj").append(div);
	return curid;
}

///<summary>Get image data as base64 string </summary>
///<param name='id'>search value of value-attribute</param>
///<returns> if found: base64 string or undefined  </returns>
function GetImageData(id) {
	return $("#binobj DIV").filter(function () { return $("#id", this).val() == id }).data("base64data");
}

///<summary>Show\Hide border around repeatable element</summary>
///<param name='override'>true=show, false=hide</param>
///<returns> if found: base64 string or undefined  </returns>
function HighlightBorder(obj, override) {
	$(obj).parent().css("border", override ? "2px solid #FF0000" : "");
}

///<summary>Show popup image</summary>
///<param name='source'>source value of image</param>
function ShowPrevImage(source) {
	// get width\heihgt values
	var strWH = $("#binobj DIV").filter(function () { return ("fbw-internal:#" + $("#id", this).val()) == source }).find("#dims").val();

	if (strWH === undefined) return; //not found
	var imgWidth = Number(strWH.slice(0, strWH.indexOf("x")));
	var imgHeight = Number(strWH.slice(strWH.indexOf("x") + 1));

	var btnHeight = event.srcElement.offsetHeight;

	// Calculate window&image dimensions
	coordX = event.clientX;
	coordY = event.clientY;

	var scrollX = 0;
	var scrollY = 0;
	if (document.documentElement && (document.documentElement.scrollTop || document.documentElement.scrollLeft)) {
		scrollX += document.documentElement.scrollLeft;
		scrollY += document.documentElement.scrollTop;
	}

	var winWidth = 0;;
	var winHeight = 0;
	if (document.documentElement && document.documentElement.clientWidth && document.documentElement.clientHeight) {
		winWidth = document.documentElement.clientWidth;
		winHeight = document.documentElement.clientHeight;
	}

	var place = "top";
	var baseWidth = coordX;
	if (baseWidth > (winWidth - coordX))
		baseWidth = winWidth - coordX;

	baseWidth = baseWidth * 2;

	var baseHeight = coordY;
	if (baseHeight < (winHeight - coordY)) {
		baseHeight = winHeight - coordY;
		place = "bottom";
	}

	baseHeight -= btnHeight;

	var ratio;
	if (baseWidth < 200) {
		ratio = 200 / baseWidth;
		baseWidth = 200;
		baseHeight *= ratio;
	}

	var scaleTo = "width";
	ratio = baseWidth / imgWidth;
	if (imgWidth < imgHeight) {
		scaleTo = "height";
		ratio = baseHeight / imgHeight;
	}

	if (imgWidth > baseWidth || imgHeight > baseHeight) {
		switch (scaleTo) {
			case "width":
				imgWidth = baseWidth;
				imgHeight = ratio * imgHeight;
				if (imgHeight > baseHeight) {
					ratio = baseHeight / imgHeight;
					imgHeight = baseHeight;
					imgWidth = ratio * imgWidth;
				}
				break;
			case "height":
				imgHeight = baseHeight;
				imgWidth = ratio * imgWidth;
				if (imgWidth > baseWidth) {
					ratio = baseWidth / imgWidth;
					imgWidth = baseWidth;
					imgHeight = ratio * imgHeight;
				}
				break;
		}
	}
	$("#prevImg").attr({
		"src": source,
		"width": imgWidth,
		"height": imgHeight
	}).css("cursor", "default");
	$("#prevImgPanel").css({
		"left": (coordX + scrollX - Math.round(imgWidth / 2)) + "px",
		"top": (place == "top") ? (coordY + scrollY - Math.round(imgHeight) - btnHeight) + "px"
			: (coordY + scrollY + btnHeight) + "px"
	});
	setTimeout('$("#prevImgPanel").css("visibility", "visible")', 500);
}

///<summary>Hide popup image</summary>
function HidePrevImage() {
	$("#prevImg").attr({
		"src": "",
		"width": 0,
		"height": 0
	});
	$("#prevImgPanel").css("visibility", "hidden");
}

///<summary>Show image in full mode</summary>
///<param name='source'>source value of image</param>
function ShowFullImage(source) {
	HidePrevImage();

	// get width\heihgt values
	var strWH = $("#binobj DIV").filter(function () { return ("fbw-internal:#" + $("#id", this).val()) == source }).find("#dims").val();

	if (strWH)
		return; //not found
	var imgWidth = Number(strWH.slice(0, strWH.indexOf("x")));
	var imgHeight = Number(strWH.slice(strWH.indexOf("x") + 1));

	// Calculate window&image dimensions
	var scrollX = 0;
	var scrollY = 0;
	if (document.documentElement && (document.documentElement.scrollTop || document.documentElement.scrollLeft)) {
		scrollX += document.documentElement.scrollLeft;
		scrollY += document.documentElement.scrollTop;
	}

	var winWidth = 0;;
	var winHeight = 0;
	if (document.documentElement && document.documentElement.clientWidth && document.documentElement.clientHeight) {
		winWidth = document.documentElement.clientWidth;
		winHeight = document.documentElement.clientHeight;
	}

	$("#fullImg").attr({
		"src": source,
		"width": imgWidth,
		"height": imgHeight
	}).css({
		"top": (imgHeight < winHeight) ? ((winHeight - imgHeight) / 2) + "px" : "0px",
		"cursor": "default"
	});
	$("#fullImgPanel").css({
		"left": "0px",
		"width": (winWidth) + "px",
		"height": Math.max(winHeight, imgHeight) + "px",
		"top": (scrollY) + "px",
		"visibility": "visible"
	});
}

///<summary>Hide full image view</summary>
function HideFullImage() {
	$("#fullImg").attr({
		"src": "",
		"width": 0,
		"height": 0
	});
	$("#fullImgPanel").css("visibility", "hidden");
}

///<summary>Start saving image process</summary>
///<param name='binDiv'>DIV-element with image</param>
function SaveImage(binDiv) {
	if ($(binDiv).find("INPUT#id").val()) {
		window.external.SaveBinary(src, $(binDiv).data("base64data"), 1);
	}
}

///<summary>Load xsl</summary>
///<param name='path'>Full Path to xsl-file</param>
///<param name='lang'>Current app interface language</param>
///<returns>xslt-object with loaded xsl, or false in error</returns>
function LoadXSL(path, lang) {
	var xslt = new ActiveXObject("Msxml2.XSLTemplate.6.0");
	var xsl = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
	xsl.setProperty("ResolveExternals", true);
	xsl.async = false;
	var proc;

	xsl.load(path);
	var doc = xsl.documentElement;
	var imp = doc.firstChild;

	var ats = imp.attributes;
	var href = ats.getNamedItem("href");
	switch (lang) {
		case "russian":
			href.nodeValue = "rus.xsl";
			break;
		case "ukrainian":
			href.nodeValue = "ukr.xsl";
			break;
		case "english":
		default:
			href.nodeValue = "eng.xsl";
			break;
	}


	if (xsl.parseError.errorCode) {
		errCantLoad(xsl, path);
		return false;
	}

	xslt.stylesheet = xsl;
	return xslt;
}

///<summary>Event handler for Click</summary>
function ClickOnDesc() {
	var srcName = event.srcElement.nodeName;
	if (srcName == "FIELDSET" || srcName == "LABEL" || srcName == "DIV" || srcName == "LEGEND") {
		document.body.focus();
	}
}

///<summary>Show cover</summary>
///<param name='prntEl'>Parent DIV element</param>
///<param name='fullImg'>true=Full view, false=Preview mode</param>
function ShowCoverImage(prntEl, fullImg) {
	var cover = $(prntEl).find("SELECT").val();
	if (fullImg)
		ShowFullImage("fbw-internal:" + cover);
	ShowPrevImage("fbw-internal:" + cover);
}

///<summary>Transform xml</summary>
///<param name='xslt'>XSLT object with loaded xsl</param>
///<param name='dom'>xml-document</param>
///<returns>true - success, false - bad structure maint.html</returns>
function TransformXML(xslt, dom) {
	var body = document.getElementById("fbw_body");
	if (!body) {
		return false;
	}

	var desc = document.getElementById("fbw_desc");
	if (!desc) {
		return false;
	}

	proc = xslt.createProcessor();
	proc.input = dom;

	var par = desc.parentNode;
	proc.setStartMode("description");
	proc.transform();
	par = desc.parentNode;
	desc.removeNode(true);
	desc.innerHTML = proc.output;
	// desc.onclick = ClickOnDesc;
	par.appendChild(desc);
	PutBinaries(dom);
	SetupDescription(desc);

	proc.setStartMode("body");
	proc.transform();
	par = body.parentNode;
	body.removeNode(true);
	body.innerHTML = proc.output;
	par.appendChild(body);

	//window.external.InflateParagraphs(body);
	//	document.fbwFilename=name;
	document.urlprefix = "fbw-internal:";

	// Show needed description elements according to settings
	$("#fbw_desc SPAN[id]").each(function () {
		ShowElement($(this).attr("id"), window.external.GetExtendedStyle($(this).attr("id")));
	});

	return true;
}

///<summary>Load and transform xml</summary>
///<param name='dom'>xml-document</param>
///<param name='lang'>Current app interface language</param>
///<returns>html-document</returns>
function LoadFromDOM(dom, lang) {

	dom.async = false;
	dom.preserveWhiteSpace = true;

	// fb2.xsl should be in app directory! Why??
	var xpath = window.external.GetStylePath() + "\\fb2.xsl";

	// transform to html
	var ret = TransformXML(LoadXSL(xpath, lang), dom);

	return ret;
}

///<summary>Create xml-document from string</summary>
///<param name='text'>source text</param>
///<returns>xml-document or error</returns>
// TO-DO Transfer in C++
function XmlFromText(text) {
	var xml = new ActiveXObject("Msxml2.DOMDocument.6.0");
	xml.async = false;
	xml.preserveWhiteSpace = true;
	xml.loadXML(text);
	if (xml.parseError.errorCode) {
		return xml.parseError;
	}
	return xml;
}

// TO-DO ???
///<summary>Replace nbsp on repChar in all text node</summary>
///<param name='elem'>starting xml-node</param>
///<param name='repChar'>replacing char</param>
function recursiveChangeNbsp(elem, repChar) {
	var el = elem;
	while (el) {
		if (el.nodeType == TEXT_NODE)
			el.nodeValue = el.nodeValue.replace(/\u00A0/g, repChar);
		if (el.nodeType == ELEMENT_NODE && el.firstChild)
			recursiveChangeNbsp(el.firstChild, repChar);
		el = el.nextSibling;
	}
}

///<summary>Create xml-document from file</summary>
///<param name='path'>Path to xml file</param>
///<param name='lang'>Current app interface language</param>
///<returns>true or error message</returns>
function apiLoadFB2_new(path, strXML, lang) {
	//	var css_filename = $("#css").attr("href");
	//	$("#css").attr("href","");
	try {
		curInterfaceLang = lang;
		var xml = new ActiveXObject("Msxml2.DOMDocument.6.0");
		xml.async = false;
		xml.preserveWhiteSpace = true;

		xml.loadXML(strXML);
		if (xml.parseError.errorCode) {
			return getTrString('"{dvalue1}" loading error:\n\n{dvalue2}\nLine: {dvalue3} , Col: {dvalue4}',
				[path, xml.parseError.reason, xml.parseError.line, xml.parseError.linepos]);
		}

		xml.setProperty("SelectionLanguage", "XPath");
		xml.setProperty("SelectionNamespaces", "xmlns:fb='" + fbNS + "' xmlns:xlink='" + xlNS + "'");

		// Replace nbsp on special char
		var nbspChar = window.external.GetNBSP();
		if (nbspChar) {
			if (nbspChar != "\u00A0") {
				// replace in annotation
				var sel = xml.selectSingleNode("/fb:FictionBook/fb:description/fb:title-info/fb:annotation");
				if (sel)
					recursiveChangeNbsp(sel, nbspChar);
				// replace in history
				sel = xml.selectSingleNode("/fb:FictionBook/fb:description/fb:document-info/fb:history");
				if (sel)
					recursiveChangeNbsp(sel, nbspChar);
				// replace in body
				sel = xml.selectSingleNode("/fb:FictionBook/fb:body");
				while (sel) {
					if (sel.nodeName == "body")
						recursiveChangeNbsp(sel, nbspChar);
					sel = sel.nextSibling;
				}
			}
		}

		if (!LoadFromDOM(xml, lang))
			return getTrString('Error loading file:"{dvalue1}":\nCan\'t prepare document for Body mode.', [path]);

		// generate GUID if it is empty
		var $id = $("#fbw_desc #diID");
		if ($id.Length)
			if (path.indexOf("blank.fb2") != -1) {
				$id.val(window.external.GetUUID());
			}
	}
	catch (e) {
		alert('Ошибка ' + e.name + ":" + e.message + "\n" + e.stack);
	}
	// apiShowDesc(false);

	// restore CSS href
	//$("#css").attr("href", css_filename);

	return true;
}

///<summary>Switch description/body views</summary>
///<param name='state'>true for description view</param>
function apiShowDesc(state) {
	if (state) {
		$("#fbw_body").css("display", "none");
		$("#fbw_desc").css("display", "block");
		//document.ScrollSave = document.body.scrollTop;
		document.body.scrollTop = 0;
	}
	else {
		$("#fbw_body").css("display", "block");
		$("#fbw_desc").css("display", "none");
		//document.body.scrollTop = document.ScrollSave;
	}
}

//// == DESC == ///////////////////////////////////////////////////////////////////

///<summary>Remove DIV but keep children</summary>
///<param name='className'>DIV class</param>
function apiCleanUp(className) {
	var cnt = $("DIV." + className).contents();
	$("DIV." + className).replaceWith(cnt);
	/*
		var divs = document.getElementsByTagName("DIV");
		for(var i=0; i < divs.length; i++)
		{
			if (divs[i].className == className)
				RemoveOuterTags(divs[i]);
		}
	*/
}

///<summary>Build Cover dropdown list using all images names</summary>
function FillLists() {
	$("#tiCover, #stiCover").find("#href").each(function () {
		var cover = $(this).val();
		if (cover === null)
			cover = "";

		$(this).empty();

		$(this).append($("<option value=''></option>"));

		var $list = $(this);
		$("#binobj DIV #id").each(function () {
			var ext = $(this).val();
			if (ext.search(/\.jpg|\.png|\.jpeg/i) != -1) {
				$list.append("<option value='#" + ext + "'>" + ext + "</option>");
			}
		});
		$(this).val(cover);
	});
}

///<summary>Generates new ID and put it into the field</summary>
function NewDocumentID(msg) {
	if (AskYesNo(msg))
		$("#fbw_desc diID").val(window.external.GetUUID());
}

///<summary>Puts lines (borderTop)to space several elements of same type
// Also prevents deletion of last element</summary>
///<param name='obj'>Container with cloned elements</param>
function PutSpacers(obj) {
	if (obj.id == "binobj") return;
	var $jq = $(obj).find("DIV");

	$jq.each(function (index) {
		if (obj.nodeName != "DIV" && index == 0) {
			$(this).css({ "marginTop": "0", "paddingTop": "0", "borderTop": "none" });
			$(this).find("#del").prop("disabled", $jq.length <= 1);
		}
		else {
			$(this).css({ "marginTop": "0.2em", "paddingTop": "0.2em", "borderTop": "solid #D0D0BF 1px" });
			$(this).children("#del").prop("disabled", false);
		}
	});
}

// Initializes fieldsets.
function InitFieldsets() {
	$("fieldset").each(function () {
		PutSpacers(this);
	});
}

///<summary>Set languages in drop-down box</summary>
function SetSelLangComboBox(id) {
	$langList = $(id);
	var sel = $langList.val();
	$langList.find("option").filter(function () { return this.text == "-org-"; }).remove();
	$langList.find("option[value='" + sel + "']").prop("selected", true)
}

///<summary>Set selected languages</summary>
function SelectLanguages() {
	// title-info language
	SetSelLangComboBox("#tiLang");
	SetSelLangComboBox("#tiSrcLang");
	SetSelLangComboBox("#stiLang");
	SetSelLangComboBox("#stiSrcLang");
}

///<summary>Initializes default fields in description section if needed</summary>
///<param name='desc'>Description container element</param>
function SetupDescription(desc) {
	// set version of document if it's empty
	if ($(desc).find("#diVersion").val() == "")
		$(desc).find("#diVersion").val("1.0");

	// set current dates if they are empty
	var date = new Date();

	var year = date.getFullYear();

	var day = date.getDate();
	if (day < 10) day = "0" + day;

	var mon = date.getMonth() + 1;
	if (mon < 10) mon = "0" + mon;

	if ($(desc).find("#diDate").val() == "")
		$(desc).find("#diDate").val(year + "-" + mon + "-" + day);

	if ($(desc).find("#diDateVal").val() == "")
		$(desc).find("#diDateVal").val(year + "-" + mon + "-" + day);

	// set ID for document if it's empty
	if ($(desc).find("#diID").val() == "")
		$(desc).find("#diID").val(window.external.GetUUID());

	//puts lines (borderTop) to space several elements of same type in description
	$("fieldset").each(function () {
		PutSpacers(this);
	});

	//set Editor used if it's empty
	var $prgs = $(desc).find("#diProgs");
	if ($prgs.val() == "") {
		$prgs.val("FictionBook Editor " + window.external.GetProgramVersion());
	}

	SelectLanguages();
}

///<summary>Shows Genres menu and gets the genre (button event handler)</summary>
function GetGenre(d) {
	var v = window.external.GenrePopup(d, window.event.screenX, window.event.screenY);
	if (v) {
		$(d).parent().find("#genre").val(v);
	}
}

///<summary>Clone element without DIVs children</summary>
///<param name='obj'>Element to clone</param>
function dClone(obj) {
	/*	$qn = $(obj).clone();
		return $qn.children("DIV").remove();*/
	var qn = obj.cloneNode(true);
	var cn = qn.firstChild;

	while (cn) {
		var nn = cn.nextSibling;
		if (cn.nodeName == "DIV")
			cn.removeNode(true);
		cn = nn;
	}
	return qn;
}

///<summary>Delete image</summary>
///<param name='obj'>DIV with image</param>
function Remove(obj) {
	if (obj.base64data != null) {
		var pic_id = $(obj).find("input#id").val();
		if (pic_id === undefined) pic_id = "";
	}

	var pn = obj.parentNode;
	obj.removeNode(true);
	PutSpacers(pn);

	//update image????? src=pic_id??
	if (pic_id) {
		pic_id = "fbw-internal:#" + pic_id;
		$("IMG[src='" + pic_id + "']").removeAttr("src").attr("src", pic_id);
		// update Cover lists
		FillLists();
	}
}

///<summary>Add element of same type</summary>
///<param name='obj'>Element to clone</param>
function Clone(obj) {
	$(obj).after(dClone(obj));
	PutSpacers(obj.parentNode);

	if ($(obj).attr("id") == 'href')
		FillLists(); // this is the Cover field, refill dropdown
	/* HTML DOM variant
	obj.insertAdjacentElement("afterEnd",dClone(obj)); 
    PutSpacers(obj.parentNode);

    if(obj.id=='href') FillLists(); // this is the Cover field
 */
}

///<summary>Clone child element of same type in description view</summary>
///<param name='obj'>Container with cloned elements</param>
function ChildClone(obj) {
	$(obj).append(dClone(obj));
	PutSpacers(obj);
}

//// == DESC == ///////////////////////////////////////////////////////////////////

/// WORKING WITH XMLDocument

var fbNS = "http://www.gribuser.ru/xml/fictionbook/2.0";
var xlNS = "http://www.w3.org/1999/xlink";

///<summary>Create indentation</summary>
///<param name='node'>Node</param>
///<param name='len'>indent size</param>
function Indent(node, len) {
	var s = "\r\n";
	while (len--)
		s = s + " ";
	node.appendChild(node.ownerDocument.createTextNode(s));
}

///<summary>Create attribute and set value</summary>
///<param name='node'>Node</param>
///<param name='name'>attribute name</param>
///<param name='val'>attribute value</param>
function SetAttr(node, name, val) {
	var an = node.ownerDocument.createNode(ATTRIBUTE_NODE, name, fbNS);
	an.appendChild(node.ownerDocument.createTextNode(val));
	node.setAttributeNode(an);
}

///<summary>Create xml-element with indented text</summary>
///<param name='node'>xml-node</param>
///<param name='name'>Node name</param>
///<param name='val'>text</param>
///<param name='force'>if force=true, allow empty text</param>
///<param name='indent'>Indentation size</param>
///<returns>true if node added and text is nonempty</returns>
function MakeText(node, name, val, force, indent) {
	var ret = true;
	if (val.length == 0) {
		if (!force)
			return false
		else ret = false;
	}
	var tn = node.ownerDocument.createNode(ELEMENT_NODE, name, fbNS);
	tn.appendChild(node.ownerDocument.createTextNode(val));
	Indent(node, indent);
	node.appendChild(tn);
	return ret;
}

///<summary>Transfer author info to xml</summary>
///<param name='node'>Author xml-node</param>
///<param name='name'>Author xml-node name</param>
///<param name='dv'>DIV with author info</param>
///<param name='force'>if force=true, allow empty text</param>
///<param name='indent'>Indentation size</param>
///<returns>true if xml-node created</returns>
function MakeAuthor(node, name, dv, force, indent) {
	var au = node.ownerDocument.createNode(ELEMENT_NODE, name, fbNS);
	var added = false;

	var $dv = $(dv);
	if ($dv.find("#first").val() || $dv.find("#middle").val() ||
		$dv.find("#last").val() || $dv.find("#nick").val()) {
		added = MakeText(au, "first-name", $dv.find("#first").val(), true, indent + 1) || added;
		added = MakeText(au, "middle-name", $dv.find("#middle").val(), false, indent + 1) || added;
		added = MakeText(au, "last-name", $dv.find("#last").val(), true, indent + 1) || added;
		if ($dv.find("#id").length > 0)
			added = MakeText(au, "id", $dv.find("#id").val(), false, indent + 1) || added;
	}
	added = MakeText(au, "nickname", $dv.find("#nick").val(), false, indent + 1) || added;
	added = MakeText(au, "home-page", $dv.find("#home").val(), false, indent + 1) || added;
	added = MakeText(au, "email", $dv.find("#email").val(), false, indent + 1) || added;

	if (added || force) {
		Indent(node, indent);
		node.appendChild(au);
		Indent(au, indent);
	}
	return added;
}

///<summary>Create xml-node with dates</summary>
///<param name='node'>xml-node</param>
///<param name='d'>Date</param>
///<param name='v'>Date display value</param>
///<param name='indent'>Indentation size</param>
///<returns>true if xml-node created</returns>
function MakeDate(node, d, v, indent) {
	var dt = node.ownerDocument.createNode(ELEMENT_NODE, "date", fbNS);
	var added = false;
	if (v.length > 0) {
		added = true;
		SetAttr(dt, "value", v);
	}
	if (d.length > 0) added = true;

	dt.appendChild(node.ownerDocument.createTextNode(d));
	Indent(node, indent);
	node.appendChild(dt);
	return added;
}

///<summary>Transfer sequence info to xml</summary>
///<param name='xn'>xml-node</param>
///<param name='hn'>DIV with sequence info</param>
///<param name='indent'>Indentation size</param>
///<returns>true if xml-node created</returns>
function MakeSeq2(xn, hn, indent) {
	var added = false;
	var newxn = xn.ownerDocument.createNode(ELEMENT_NODE, "sequence", fbNS);

	var name = $(hn).find("#name").first().val();
	var num = $(hn).find("#number").first().val();
	SetAttr(newxn, "name", name);
	if (num)
		SetAttr(newxn, "number", num);

	$(hn).children("DIV").each(function () {
		added = MakeSeq2(newxn, this, indent + 1) || added;
	});

	if (added || name.length > 0 || num.length > 0) {
		Indent(xn, indent);
		xn.appendChild(newxn);
		if (newxn.hasChildNodes())
			Indent(newxn, indent);
		added = true;
	}
	return added;
}

function MakeSeq(xn, hn, indent) {
	var added = false;
	$(hn).children("DIV").each(function () {
		added = MakeSeq2(xn, this, indent + 1) || added;
	});

	return added;
}

function IsEmpty(ii) {
	if (!ii || !ii.hasChildNodes()) return true;

	for (var v = ii.firstChild; v; v = v.nextSibling)
		if (v.nodeType == ELEMENT_NODE && v.baseName != "empty-line")
			return false;

	return true;
}

///<summary>Create and fill title info</summary>
///<param name='doc'>xml-document</param>
///<param name='desc'>HTML description element</param>
///<param name='ann'>Text of annotation</param>
///<param name='indent'>Indentation size</param>
function MakeTitleInfo(doc, desc, ann, indent) {
	var ti = doc.createNode(ELEMENT_NODE, "title-info", fbNS);
	Indent(desc, indent);
	desc.appendChild(ti);

	// genres
	$("#tiGenre DIV").each(function () {
		var ge = doc.createNode(ELEMENT_NODE, "genre", fbNS);

		var match = $(this).find("#match").val();
		if (match.length > 0 && match != "100")
			SetAttr(ge, "match", match);

		ge.appendChild(doc.createTextNode($(this).find("#genre").val()));
		Indent(ti, indent + 1);
		ti.appendChild(ge);
	});

	/*
	 var list=document.all.tiGenre.getElementsByTagName("DIV");
	 for(var i=0; i<list.length; i++)
	 {
	  var ge=doc.createNode(ELEMENT_NODE,"genre",fbNS);
	  if (list.item(i).all.match)
	  {
		var match = list.item(i).all.match.value;
		if(match.length>0 && match!="100") SetAttr(ge,"match",match);
	  }
	
	  ge.appendChild(doc.createTextNode(list.item(i).all.genre.value));
	  Indent(ti,indent+1);
	  ti.appendChild(ge);
	 }
	*/
	// authors
	$("#tiAuthor DIV").each(function () {
		MakeAuthor(ti, "author", this, false, indent + 1);
	});

	/*if(!added && $list.length>0) 
	   MakeAuthor(ti,"author",$list.first(),true,indent+1);*/
	// book title
	MakeText(ti, "book-title", $("#tiTitle").val(), true, indent + 1);
	/*
	 var added=false;
	 list=document.all.tiAuthor.getElementsByTagName("DIV");
	 for(var i=0; i<list.length; i++) added=MakeAuthor(ti,"author",list.item(i),false,indent+1) || added;
	
	 if(!added && list.length>0) MakeAuthor(ti,"author",list.item(0),true,indent+1);
	
	 MakeText(ti,"book-title",document.all.tiTitle.value,true,indent+1);
	*/

	// annotation, will be filled by body.xsl
	if (!IsEmpty(ann)) {
		Indent(ti, indent + 1);
		ti.appendChild(ann);
	}

	MakeText(ti, "keywords", $("#tiKwd").val(), false, indent + 1);
	MakeDate(ti, $("#tiDate").val(), $("#tiDateVal").val(), indent + 1);


	// coverpage images
	var cp = doc.createNode(1, "coverpage", fbNS);
	$("#tiCover DIV #href").each(function () {
		if (this.value.length > 0) {
			var xn = doc.createNode(ELEMENT_NODE, "image", fbNS);
			var an = doc.createNode(ATTRIBUTE_NODE, "l:href", xlNS);
			an.appendChild(doc.createTextNode(this.value));
			xn.setAttributeNode(an);
			Indent(cp, indent + 2);
			cp.appendChild(xn);
		}
	});

	if (cp.hasChildNodes) {
		Indent(ti, indent + 1);
		ti.appendChild(cp);
	}

	MakeText(ti, "lang", $("#tiLang").val(), false, indent + 1);
	MakeText(ti, "src-lang", $("#tiSrcLang").val(), false, indent + 1);

	// translator
	$("#tiTrans DIV").each(function () {
		MakeAuthor(ti, "translator", this, false, indent + 1);
	});
	/* 
	 // translator
	 list=document.all.tiTrans.getElementsByTagName("DIV");
	
	 for(var i=0; i<list.length; i++) MakeAuthor(ti,"translator",list.item(i),false,indent+1);
	*/
	// sequence
	MakeSeq(ti, $("#tiSeq"), indent + 1);
	Indent(ti, indent);
}
/*
function IsSeqExist2(hn) {
  var name =  hn.all("name",0).value;
  var num  =  hn.all("number",0).value;
  if (name || num)
    return true;

  for (var cn = hn.firstChild; cn; cn = cn.nextSibling) {
    if (cn.nodeName == "DIV") {
        if(IsSeqExist2(cn))
            return true;
    }
  }
  return false;
}

function IsSeqExist(hn) {
  for (var cn = hn.firstChild; cn; cn=cn.nextSibling) {
    if (cn.nodeName == "DIV") {
        if(IsSeqExist2(cn))
            return true;
    }
  }
  return false;
}
*/
///<summary>Check if there are any info in Source title info block</summary>
///<returns>true if not empty</returns>
function IsSTBFieldTextExist() {
	// authors
	if ($("#stiAuthor DIV input").filter(function () { return this.value != ""; }).length > 0)
		return true;
	/*list = document.getElementById("tiAuthor").getElementsByTagName("DIV");
	for (var i = 0; i < list.length; ++i) {
		if (list.item(i).querySelector("#first").value || list.item(i).querySelector("#middle").value ||
			list.item(i).querySelector("#last").value || list.item(i).querySelector("nick").value ||
			list.item(i).querySelector("#home").value || list.item(i).querySelector("#email").value)
			return true;
	}*/

	// book title
	if ($("#stiTitle").val())
		return true;

	// genres
	if ($("#stiGenre DIV #genre").val())
		return true;

	// keywords
	if ($("#stiKwd").val())
		return true;

	// date
	if ($("#stiDateVal").val() || $("#stiDate").val())
		return true;

	// coverpage images
	if ($("#stiCover DIV").filter(function () { return $("#href", this).val() }).length > 0)
		return true;

	// lang
	if ($("#stiLang").val() || $("#stiSrcLang").val())
		return true;

	if ($("#stiTrans DIV input").filter(function () { return this.value != "" }).length > 0)
		return true;

	// translator
	/*list = document.stiTrans.getElementsByTagName("DIV");
	for (var i = 0; i < list.length; ++i) {
		if (list.item(i).first.value || list.item(i).middle.value || list.item(i).last.value ||
			list.item(i).nick.value || list.item(i).home.value || list.item(i).email.value)
			return true;
	}*/

	// sequence
	if ($("#stiSeq DIV input").filter(function () { return this.value != "" }).length > 0)
		return true;
	//if(IsSeqExist(document.all.stiSeq))
	//  return true;
	return false;
}

///<summary>Fill Source title info block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='desc'>description xml-element</param>
///<param name='ann'>Not used (annotation??)</param>
///<param name='indent'>Indent value</param>
function MakeSourceTitleInfo(doc, desc, ann, indent) {
	var sti = doc.createNode(1, "src-title-info", fbNS);
	if (IsSTBFieldTextExist()) {
		Indent(desc, indent);
		desc.appendChild(sti);

		// genres
		$("#stiGenre DIV").each(function () {
			var ge = doc.createNode(1, "genre", fbNS);
			var match = $("#match", this).val();
			if (match.length > 0 && match != "100")
				SetAttr(ge, "match", match);
			ge.appendChild(doc.createTextNode($("#genre", this).val()));
			Indent(sti, indent + 1);
			sti.appendChild(ge);
		});
		// authors
		var added = false;
		$("#stiAuthor DIV").each(function () {
			added = MakeAuthor(sti, "author", this, false, indent + 1) || added;
		});
		if (!added && $("#stiAuthor DIV").length > 0) {
			MakeAuthor(sti, "author", $("#stiAuthor DIV").first(), true, indent + 1);
		}

		MakeText(sti, "book-title", $("#stiTitle").val(), true, indent + 1);
		MakeText(sti, "keywords", $("#stiKwd").val(), false, indent + 1);
		MakeDate(sti, $("#stiDate").val(), $("#stiDateVal").val(), indent + 1);

		// coverpage images
		var cp = doc.createNode(ELEMENT_NODE, "coverpage", fbNS);
		$("#stiCover DIV #href").each(function () {
			var xn = doc.createNode(ELEMENT_NODE, "image", fbNS);
			var an = doc.createNode(ATTRIBUTE_NODE, "l:href", xlNS);
			an.appendChild(doc.createTextNode(this.value));
			xn.setAttributeNode(an);
			Indent(cp, indent + 2);
			cp.appendChild(xn);
		});

		if (cp.hasChildNodes) {
			Indent(sti, indent + 1);
			sti.appendChild(cp);
		}

		MakeText(sti, "lang", $("#stiLang").val(), false, indent + 1);
		MakeText(sti, "src-lang", $("#stiSrcLang").val(), false, indent + 1);

		// translator
		$("#stiTrans DIV").each(function () {
			MakeAuthor(sti, "translator", this, false, indent + 1);
		});
		/*		list = document.stiTrans.getElementsByTagName("DIV");
				for (var i = 0; i < list.length; ++i)
					MakeAuthor(sti, "translator", list.item(i), false, indent + 1);
		*/
		// sequence
		MakeSeq(sti, $("#stiSeq"), indent + 1);
		Indent(sti, indent);
	}
}

///<summary>Fill Book info block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='desc'>description xml-element</param>
///<param name='hist'>History</param>
///<param name='indent'>Indent value</param>
function MakeDocInfo(doc, desc, hist, indent) {
	var added = false;
	var di = doc.createNode(1, "document-info", fbNS);

	// authors
	var i;

	// authors
	var added = false;
	$("#diAuthor DIV").each(function () {
		added = MakeAuthor(di, "author", this, false, indent + 1) || added;
	});
	if (!added && $("#diAuthor DIV").length > 0) {
		MakeAuthor(di, "author", $("#diAuthor DIV").first(), true, indent + 1);
	}

	added = MakeText(di, "program-used", $("#diProgs").val(), false, indent + 1) || added;
	added = MakeDate(di, $("#diDate").val(), $("#diDateVal").val(), indent + 1) || added;

	// src-url
	$("#diURL INPUT").each(function () {
		added = MakeText(di, "src-url", this.value, false, indent + 1) || added;
	});

	added = MakeText(di, "src-ocr", $("#diOCR").val(), false, indent + 1) || added;
	added = MakeText(di, "id", $("#diID").val(), true, indent + 1) || added;
	added = MakeText(di, "version", $("#diVersion").val(), true, indent + 1) || added;

	// history
	if (!IsEmpty(hist)) {
		Indent(di, indent + 1);
		di.appendChild(hist);
		added = true;
	}

	// only append document info it is non-empty
	if (added) {
		Indent(desc, indent);
		desc.appendChild(di);
		Indent(di, indent);
	}
}

///<summary>Fill Publish info block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='desc'>description xml-element</param>
///<param name='indent'>Indent value</param>
function MakePubInfo(doc, desc, indent) {
	var added = false;
	var pi = doc.createNode(1, "publish-info", fbNS);

	added = MakeText(pi, "book-name", $("#piName").val(), false, indent + 1) || added;
	added = MakeText(pi, "publisher", $("#piPub").val(), false, indent + 1) || added;
	added = MakeText(pi, "city", $("#piCity").val(), false, indent + 1) || added;
	added = MakeText(pi, "year", $("#piYear").val(), false, indent + 1) || added;
	added = MakeText(pi, "isbn", $("#piISBN").val(), false, indent + 1) || added;

	// sequence
	added = MakeSeq(pi, $("#piSeq"), indent + 1) || added;

	// append non-empty publisher info only
	if (added) {
		Indent(desc, indent);
		desc.appendChild(pi);
		Indent(pi, indent);
	}
}

///<summary>Fill Custom info block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='desc'>description xml-element</param>
///<param name='indent'>Indent value</param>
function MakeCustInfo(doc, desc, indent) {
	$("#ci DIV").each(function () {
		var t = $(this).find("#type").val();
		var v = $(this).find("#val").val();
		if (t || v) {
			var ci = doc.createNode(1, "custom-info", fbNS);
			SetAttr(ci, "info-type", t);
			ci.appendChild(doc.createTextNode(v));
			Indent(desc, indent);
			desc.appendChild(ci);
		}
	});
}

///<summary>Fill Stylesheet block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='indent'>Indent value</param>
function MakeStylesheets(doc, indent) {
	var s = $("#stylesheetId").val();
	if (s) {
		var styles = doc.createNode(1, "stylesheet", fbNS);
		SetAttr(styles, "type", "text/css");
		Indent(doc.documentElement, 1);
		doc.documentElement.appendChild(styles);
		Indent(styles, 1);
		styles.appendChild(doc.createTextNode(s));
	}
}

///<summary>Fill Description block in xml</summary>
///<param name='doc'>xml doc</param>
///<param name='ann'>Annottion</param>
///<param name='hist'>History</param>
function GetDesc(doc, ann, hist) {
	MakeStylesheets(doc, 1);
	var desc = doc.createNode(1, "description", fbNS);
	Indent(doc.documentElement, 1);
	doc.documentElement.appendChild(desc);
	MakeTitleInfo(doc, desc, ann, 2);
	MakeSourceTitleInfo(doc, desc, ann, 2);
	MakeDocInfo(doc, desc, hist, 2);
	MakePubInfo(doc, desc, 2);
	MakeCustInfo(doc, desc, 2);
	Indent(desc, 1);
}

///<summary>Fill Binaries block in xml</summary>
///<param name='doc'>xml doc</param>
function GetBinaries(doc) {
	$("#binobj DIV").each(function () {
		var newb = doc.createNode(1, "binary", fbNS);
		newb.dataType = "bin.base64";
		newb.nodeTypedValue = $(this).data("base64data");

		SetAttr(newb, "id", $(this).find("#id").val());
		SetAttr(newb, "content-type", $(this).find("#type").val());

		newb.dataType = undefined;
		Indent(doc.documentElement, 1);
		doc.documentElement.appendChild(newb);
	});
	/* var bo=document.all.binobj.getElementsByTagName("DIV");
	 for(var i=0; i<bo.length; i++)
	 {
	  var newb=doc.createNode(1,"binary",fbNS);
	  newb.dataType="bin.base64";
	  newb.nodeTypedValue=bo[i].base64data;
	
	  SetAttr(newb,"id",bo[i].all.id.value);
	  SetAttr(newb,"content-type",bo[i].all.type.value);
	
	  newb.dataType=undefined;
	  Indent(doc.documentElement,1);
	  doc.documentElement.appendChild(newb);
	 }*/
}

///<summary>Add all binary data in HTML</summary>
///<param name='doc'>xml doc</param>
function PutBinaries(doc) {
	var nerr = 0;
	var bl = doc.selectNodes("/fb:FictionBook/fb:binary");

	for (var i = 0; i < bl.length; i++) {
		if (bl[i].tagName != "binary") continue;

		bl[i].dataType = "bin.base64";
		var id = bl[i].getAttribute("id");
		var dt;

		try {
			dt = bl[i].nodeTypedValue;
		}
		catch (e) {
			if (nerr++ < 3)
				MsgBox("Invalid base64 data for " + id);
			continue;
		}

		apiAddBinary("", id, bl[i].getAttribute("content-type"), dt);
	}

	if (nerr > 3)
		MsgBox(nerr - 3 + " more invalid images ignored");

	// update Cover lists
	FillLists();
}

//// == BODY == ///////////////////////////////////////////////////////////////////

function GoTo(elem) {
	if (!elem) return;
	var b = elem.getBoundingClientRect();
	if (b.bottom - b.top <= window.external.getViewHeight())
		window.scrollBy(0, (b.top + b.bottom - window.external.getViewHeight()) / 2);
	else
		window.scrollBy(0, b.top);
	//var r = document.getSelection()
	var r = document.createRange();
	if (!r || !("compareEndPoints" in r))
		return;
	r.moveToElementText(elem);
	r.collapse(true);
	if (r.parentElement !== elem && r.move("character", 1) == 1)
		r.move("character", -1);
	r.select();
}

// Skip elements: empty <P>, n1-n3
// returns next sibling
function SkipOver(np, n1, n2, n3) {
	while (np) {
		if (!(np.tagName == "P" && (!np.firstChild || np.firstChild.tagname == "br")) &&  //not an empty P
			(!n1 || (np.tagName != n1 && np.className != n1)) && // and not n1
			(!n2 || (np.tagName != n2 && np.className != n2)) && // and not n2
			(!n3 || (np.tagName != n3 && np.className != n3)))   // and not n3
			break;
		np = np.nextSibling;
	}
	return np;
}
//-----------------------------------------------

function StyleCheck(cp, st) {
	if (!cp || cp.tagName != "P")
		return false;

	var pp = cp.parentElement;
	if (!pp || pp.tagName != "DIV")
		return false;

	switch (st) {
		// ordinary style
		case "":
			if (pp.className != "section" && pp.className != "title" && pp.className != "epigraph" &&
				pp.className != "stanza" && pp.className != "cite" && pp.className != "annotation" &&
				pp.className != "history")
				return false;
			break;

		case "subtitle":
			if (pp.className != "section" && pp.className != "stanza" && pp.className != "cite" && pp.className != "annotation")
				return false;
			break;

		case "text-author":
			if (pp.className != "cite" && pp.className != "epigraph" && pp.className != "poem") return false;
			if ((cp.nextSibling && cp.nextSibling.className != "text-author")) return false;
			break;

		case "code":
			if ((cp.className == "text-author" || cp.className == "subtitle" && pp.className == "section") ||
				(pp.className == "stanza" && cp.tagName == "P") ||
				(cp.className == "" && pp.className == "section" && cp.tagName == "P") ||
				((cp.className == "td" || cp.className == "th") && pp.className == "tr"))
				return true;
			else
				return false;

			break;
	}

	return true;
}
//-----------------------------------------------
function SetStyle(cp, check, name) {
	if (!StyleCheck(cp, name))
		return;

	if (check)
		return true;

	if (name.length == 0)
		name = "normal";

	window.external.BeginUndoUnit(document, name + " style");
	window.external.SetStyleEx(document, cp, name);
	window.external.EndUndoUnit(document);
}

///<summary>Set style=normal</summary>
///<param name='cp'>Element</param>
///<param name='check'>if true don/t set style, just check</param>
function StyleNormal(cp, check) {
	return SetStyle(cp, check, "");
}

function SetStyleNormal() {
	if (!CanApplyStyle("normal"))
		return;

	// can be inside section, body, poem, stanza
	var rng = document.selection.createRange();
	if (!rng)
		return false;

	var begin_range = rng.duplicate();
	var end_range = rng.duplicate();
	begin_range.collapse(true);
	end_range.collapse(false);

	var begin_str = "";
	var begin_el = begin_range.parentElement();

	// add closed inline tags
	while (begin_el.tagName != "P") {
		switch (begin_el.tagName) {
			case "B":
			case "I":
			case "SUP":
			case "SUB":
			case "STRIKE":
			case "U":
				begin_str = "</" + begin_el.tagName + ">" + begin_str;
				break;
			default:
				break;
		}
		begin_el = begin_el.parentElement;
	}

	var end_str = "";
	var end_el = end_range.parentElement();

	// add opened inline tags
	while (end_el.tagName != "P") {
		switch (end_el.tagName) {
			case "B":
			case "I":
			case "SUP":
			case "SUB":
			case "STRIKE":
			case "U":
				end_str += "<" + end_el.tagName + ">";
				break;
			default:
				break;
		}
		end_el = end_el.parentElement;
	}

	var int_str = rng.htmlText;
	alert(rng.htmlText);
	//remove inline style tags
	int_str = int_str.replace(/<STRONG>/gi, "");
	int_str = int_str.replace(/<EM>/gi, "");
	int_str = int_str.replace(/<SUP>/gi, "");
	int_str = int_str.replace(/<SUB>/gi, "");
	int_str = int_str.replace(/<STRIKE>/gi, "");
	int_str = int_str.replace(/<U>/gi, "");
	int_str = int_str.replace(/<\/STRONG>/gi, "");
	int_str = int_str.replace(/<\/EM>/gi, "");
	int_str = int_str.replace(/<\/SUP>/gi, "");
	int_str = int_str.replace(/<\/SUB>/gi, "");
	int_str = int_str.replace(/<\/STRIKE>/gi, "");
	int_str = int_str.replace(/<\/U>/gi, "");
	alert(int_str);

	window.external.BeginUndoUnit(document, "reset style");
	rng.pasteHTML(begin_str + int_str + end_str);
	window.external.EndUndoUnit(document);
}

///<summary>Set style=text-author</summary>
///<param name='cp'>Element</param>
///<param name='check'>if true don/t set style, just check</param>
function StyleTextAuthor(cp, check) {
	return SetStyle(cp, check, "text-author");
}

///<summary>Set style=subtitle</summary>
///<param name='cp'>Element</param>
///<param name='check'>if true don't set style, just check</param>
function StyleSubtitle(cp, check) {
	return SetStyle(cp, check, "subtitle");
}

///<summary>Set style=code</summary>
///<param name='cp'>Element</param>
///<param name='check'>if true don't set style, just check</param>
///<param name='range'>Range object</param>
function StyleCode(check, cp, range) {
	if (check && cp && range) {
		// That is due to MSHTML bug
		if (cp.tagName == "P") {
			html = new String(range.htmlText);
			if (html.indexOf("<DIV") != -1) {
				return false;
			}
		}

		if (cp.tagName == "DIV" && cp.className == "section") {
			var end = range.duplicate();
			range.collapse(true);
			end.collapse(false);

			per = range.parentElement();
			ped = end.parentElement();

			while (per && per.tagName != "P")
				per = per.parentElement;
			while (ped && ped.tagName != "P")
				ped = ped.parentElement;

			if (per && ped && per.parentElement == ped.parentElement) {
				return true;
			}
		}
	}

	return SetStyle(cp, check, "code");
}

///<summary>Check if it's possible to apply style or add image</summary>
///<returns>true if can</returns>
function CanApplyStyle(st) {
	if (document.selection.type == "Control")
		return false;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	switch (st) {
		// subtitle style
		case "subtitle":
			// can be in section|annotation|cite (not realized: can be in stanza also but just one after title!)
			if (!$(rng.parentElement()).closest("div.section,.stanza,.annotation,.cite,.history")[0])
				return false;
			break;

		case "body_notes":
			if ($("div.body[fbname=notes]").length>0)
				return false;

			break;
		case "code":
			break;

		case "bold":
		case "italic":
		case "superscript":
		case "subscript":
		case "strikethrough":
		case "normal":
			break;

		case "link":
			// can be just in v|th|td|text-author|subtitle|no-class
			if (rng.parentElement() && rng.parentElement().nodeName != "P")
				return false;

			if (!$(rng.parentElement()).closest("p:not([class]),.v,.th,.td,.text-author,.subtitle")[0])
				return false;

			break;

		case "text-author":
			// can be in cite|epigraph|poem
			if (!$(rng.parentElement()).closest("div.epigraph,.cite,.poem")[0])
				return false;
			break;

		case "annotation":
			// can be inside one section only
			rng.collapse(true);
			var begin_node = $(rng.parentElement()).closest("div.section")[0];
			if (!begin_node)
				return false;

			// check if annotation already exists
			if ($(begin_node).children("div.annotation")[0])
				return false;

			break;

		case "epigraph":
			// can be inside section, body, poem
			rng.collapse(true);
			if (!$(rng.parentElement()).closest("div.section,.body,.poem")[0])
				return false;

			break;
		case "cite":
			// selection end can be inside section, annotation, epigraph
			rng.collapse(false);
			if (!$(rng.parentElement()).closest("div.section,.annotation,.epigraph")[0])
				return false;
			break;

		case "poem":
			// selection end can be inside section, cite, annotation, epigraph
			rng.collapse(false);
			if (!$(rng.parentElement()).closest("div.section,.cite,.annotation,.epigraph")[0])
				return false;
			break;

		case "stanza":
			// selection end can be inside poem
			rng.collapse(false);
			if (!$(rng.parentElement()).closest("div.poem")[0])
				return false;
			break;

		case "title":
			// can be inside section, body, poem, stanza
			rng.collapse(true);
			var begin_node = $(rng.parentElement()).closest("div.section,.body,.poem")[0];
			if (!begin_node)
				return false;

			// check if title already exists
			if ($(begin_node).children("div.title")[0])
				return false;
			break;

		case "section":
			// slect should be in one P-element
			rng.collapse(true);
			if (!$(rng.parentElement()).closest("div.section")[0])
				return false;

			break;

		case "image":
			if (rng.text != "")
				return false;
			if (!$(rng.parentElement()).closest("div.body,.section")[0] &&
				(!(rng.parentElement().nodeName == "P" && $(rng.parentElement()).closest("p:not([class]),.v,.th,.td,.text-author,.subtitle")[0]))
				)
				return false;
			break;

		case "table":
			if (!$(rng.parentElement()).closest("div.section,.cite,.annotation,.history")[0])
				return false;
			break;

		case "inlineimage":
			if (rng.parentElement().nodeName != "P")
				return false;

			if (!$(rng.parentElement()).closest("p:not([class]),.v,.th,.td,.text-author,.subtitle")[0])
				return false;
			break;

	}
	return true;
}

///<summary>Apply style or add image</summary>
///<returns>true if can</returns>
function applyStyle(st) {
	if (!canApplyStyle(st))
		return;
	/*	
		var sel = document.getSelection();
		if (!sel || sel.rangeCount == 0)
			return false;
		var rng = sel.getRangeAt(0);
		if (!rng)
			return false;
		rng = rng.cloneRange();
		
		switch (st) {
			// ordinary style
			case "normal":
				var start_p = $(rng.StartContainer).closest("p");
				var end_p = $(rng.EndContainer).closest("p");
				do
				{
					if(start_p.tagName == "P")
						start_p.class = "";
					start_p = start_p.nextSibling;
				}
				while (start_p != end_p)
	
				break;
			// subtitle style
			case "subtitle":
				var start_p = $(rng.startContainer).closest("p")[0];
				var end_p = $(rng.endContainer).closest("p")[0];
				window.external.BeginUndoUnit(document, "make subtitle");
				do {
					start_p.className = "subtitle";
					start_p = start_p.nextSibling;
				} while (start_p != end_p);
				end_p.className = "subtitle";
				window.external.EndUndoUnit(document);
				
				break;
			case "code":
				window.external.BeginUndoUnit(document, "make code");
	
				var end_p = $(rng.endContainer).closest("p")[0];
				var new_node;
				var new_node_before;
				var new_node_after;
				var p_node; // current <P> node
				var parent_node;
	
				var rng1 = document.createRange();
				var rng2 = document.createRange();
				// process first P
				var cur_node = rng.startContainer;
			    
				// process internal tags in <P>, split tags if it is necessary
				while (cur_node.parentNode.tagName != "P") {
					parent_node = cur_node.parentNode;
	
					// get left and right parts of node content
					rng1.selectNodeContents(parent_node);
					rng1.setEnd(cur_node, rng.startOffset);
					var content_before = rng1.cloneContents(); 
	
					rng2.selectNodeContents(parent_node);
					rng2.setStart(cur_node, rng.startOffset);
					var content_after = rng2.cloneContents();
				    
					// delete old node and insert two new nodes
					rng1.selectNode(parent_node);
					rng1.deleteContents();
				    
					// don't insert empty nodes
					if(content_beforenode.hasChildNodes()) {
						new_node_before.appendChild(content_before);
						rng1.insertNode(new_node_before);
						rng1.setStartAfter(new_node_before);
						cur_node = new_node_before;
						//move main range start position to split point
						rng.setStartAfter(new_node_before);
					}
					if(content_after.hasChildNodes()) {
						new_node_after = parent_node.cloneNode(false);
						rng1.insertNode(new_node_after);
						cur_node = new_node_after;
						//move main range start position to split point
						rng.setStartBefore(new_node_after);
					}
				}
				p_node = cur_node.parentNode;
	
				// set range from split point to end of <P>
				rng1.selectNodeContents(p_node);
				rng1.setStart(rng.startContainer, rng.startOffset);
	
				// wrap fragment in SPAN
				wrap_node = document.createElement("SPAN");
				wrap_node.className = "code";
				wrap_node.appendChild(rng1.extractContents());
				rng1.insertNode(wrap_node);
	
				// process all full selected <P>
				p_node = p_node.nextSibling;
				while (p_node != end_p) {
					wrap_node = document.createElement("SPAN");
					wrap_node.className = "code";
	
					rng1.selectNodeContents(p_node);
	
					wrap_node.appendChild(rng1.extractContents());
					rng1.insertNode(wrap_node);
					p_node = p_node.nextSibling;
				}
	
				// process end <P>
				cur_node = rng.endContainer;
				alert(cur_node.parentNode.nodeType);
				while (cur_node.parentNode.tagName != "P") {
					parent_node = cur_node.parentNode;
					new_node_before = cur_node.cloneNode(false);
					new_node_after = cur_node.cloneNode(false);
	
					// set ranges
					rng1.selectNodeContents(parent_node);
					rng1.setEnd(cur_node, rng.endOffset);
	
					rng2.selectNodeContents(parent_node);
					rng2.setStart(cur_node, rng.endOffset);
				    
					var content_before = rng1.cloneContents();
	
					rng2.selectNodeContents(parent_node);
					var content_after = rng2.setStart(cur_node, rng.startOffset);
	
					// delete old node and insert two new nodes
					rng1.selectNode(parent_node);
					rng1.deleteContents();
				    
					// don't insert empty nodes
					if(content_beforenode.hasChildNodes()) {
						new_node_before.appendChild(content_before);
						rng1.insertNode(new_node_before);
						rng1.setStartAfter(new_node_before);
						cur_node = new_node_before;
						//move main range start position to split point
						rng.setEndAfter(new_node_before);
					}
					if(content_after.hasChildNodes()) {
						new_node_after = parent_node.cloneNode(false);
						rng1.insertNode(new_node_after);
						cur_node = new_node_after;
						//move main range start position to split point
						rng.setEndAfter(new_node_before);
					}
					cur_node = new_node_before;
				}
	
						p_node = cur_node.parentNode;
	
				// set range from split point to end of <P>
				rng1.selectNodeContents(p_node);
				rng1.setEnd(rng.endContainer, rng.endOffset);
	
				// wrap fragment in SPAN
				wrap_node = document.createElement("SPAN");
				wrap_node.className = "code";
				wrap_node.appendChild(rng1.extractContents());
				rng1.insertNode(wrap_node);
	
				rng1.selectNodeContents(wrap_node);
	
				rng = sel.getRangeAt(0);
				rng.setEndAfter(rng1.startContainer, rng1.endOffset);
	
				sel.removeAllRanges();
				sel.addRange(rng);
				window.external.EndUndoUnit(document);
	
				break;
	
			case "link":
				// can be just in v|th|td|text-author|subtitle|no-class
				if (rng.commonAncestorContainer.nodeType != TEXT_NODE)
					return false;
	
				if (!$(rng.commonAncestorContainer).closest("p:not([class]),.v,.th,.td,.text-author,.subtitle")[0])
					return false;
	
				break;
	
			case "text-author":
				var start_p = $(rng.startContainer).closest("p")[0];
				var end_p = $(rng.endContainer).closest("p")[0];
				window.external.BeginUndoUnit(document, "make text-author");
				do {
					start_p.className = "text-author";
					start_p = start_p.nextSibling;
				} while (start_p != end_p);
				end_p.className = "text-author";
				window.external.EndUndoUnit(document);
	
				break;
	
			case "annotation":
				// can be just first element in section after title/epigraph
				if (!$(rng.commonAncestorContainer).closest("div.section")[0])
					return false;
	
				// search closest P
				var begin_node = $(rng.startContainer).closest("p")[0];
	
				if (!begin_node || (begin_node.previousSibling && begin_node.previousSibling.className.search(/title|epigraph/) < 0))
					return false;
	
				break;
	
			case "epigraph":
				if (!$(rng.commonAncestorContainer).closest("div.section,.body,.poem")[0])
					return false;
				break;
	
				// search closest P
				var begin_node = $(rng.startContainer).closest("p")[0];
	
				if (!begin_node || (begin_node.previousSibling && begin_node.previousSibling.className != "title"))
					return false;
	
				break;
			case "cite":
				if (!$(rng.commonAncestorContainer).closest("div.section,.annotation,.epigraph")[0])
					return false;
				break;
	
			case "poem":
				if (!$(rng.commonAncestorContainer).closest("div.section,.cite,.annotation,.epigraph")[0])
					return false;
				break;
	
				break;
	
			case "title":
				// can be just first element in body|section|poem|stanza
				if (!$(rng.commonAncestorContainer).closest("div.body,.section,.poem,.stanza")[0])
					return false;
	
				var begin_node = $(rng.startContainer).closest("p")[0];
	
				if (!begin_node || begin_node.previousSibling)
					return false;
	
				break;
	
			case "table":
				if (!$(rng.commonAncestorContainer).closest("div.section,.cite,.annotation,.history")[0])
					return false;
				break;
	
			case "inlineimage":
				if (rng.commonAncestorContainer.nodeType != TEXT_NODE)
					return false;
	
				if (!$(rng.commonAncestorContainer).closest("p:not([class]),.v,.th,.td,.text-author,.subtitle")[0])
					return false;
				break;
	
		}
	*/
	return true;
}

///<summary>Check if element is Code</summary>
function IsCode() {
	// as analog for bold check only selection start point
	// can be inside section, body, poem, stanza
	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(true);
	var begin_node = rng.parentElement();
	if (rng.parentElement().tagName == "U")
		return true;
	else
		return false;

//	return ($(cp).parents("SPAN.code").length > 0);
	/*  while(cp && cp.tagName != "DIV")
	  {
		if(cp.tagName == "SPAN" && cp.className == "code")
		  return true;
		cp = cp.parentElement;
	  }
	  return false;*/
}


///<summary>Replace &, < , >  to special HTML entities</summary>
///<param name='s'>text</param>
function TextIntoHTML(s) {
	if (!s) return "";
	var re0 = new RegExp("&");
	var re0_ = "&amp;";
	var re1 = new RegExp("<", "g");
	var re1_ = "&lt;";
	var re2 = new RegExp(">", "g");
	var re2_ = "&gt;";
	return s.replace(re0, re0_).replace(re1, re1_).replace(re2, re2_);
}

//-----------------------------------------------

function GetCP(cp) {
	if (!cp)
		return;

	if (cp.tagName == "P")
		cp = cp.parentElement;

	if (cp.tagName == "DIV" && cp.className == "title")
		cp = cp.parentElement;

	if (cp.tagName != "DIV")
		return;

	return cp;
}
//-----------------------------------------------

function InsBefore(parent, ref, item) {
	if (ref) ref.insertAdjacentElement("beforeBegin", item);
	else parent.insertAdjacentElement("beforeEnd", item);
}
//-----------------------------------------------

function InsBeforeHTML(parent, ref, ht) {
	if (ref) ref.insertAdjacentHTML("beforeBegin", ht);
	else parent.insertAdjacentHTML("beforeEnd", ht);
}

///<summary>Clone container</summary>
///<param name='cp'>Element to clone</param>
function CloneContainer(cp, check) {
	cp = GetCP(cp); if (!cp) return;

	switch (cp.className) {
		case "section": case "poem": case "stanza": case "cite": case "epigraph": break;
		default: return;
	}

	if (check) return true;

	window.external.BeginUndoUnit(document, "clone " + cp.className);
	var ncp = cp.cloneNode(false);

	switch (cp.className) {
		case "section": ncp.innerHTML = '<DIV class="title"><P></P></DIV><P></P>'; break;
		case "poem": ncp.innerHTML = '<DIV class="title"><P></P></DIV><DIV class="stanza"><P></P></DIV>'; break;

		case "stanza": case "cite": case "epigraph": ncp.innerHTML = '<P></P>'; break;
	}
	//InflateIt(ncp);
	cp.insertAdjacentElement("afterEnd", ncp);
	window.external.EndUndoUnit(document);
	GoTo(ncp);
}
//-----------------------------------------------

///<summary>Add title to HTML</summary>
///<param name='cp'>HTML node</param>
function AddTitle(cp) {
	if (!CanApplyStyle("title"))
		return;

	// can be inside section, body, poem, stanza
	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(true);
	var begin_node = $(rng.parentElement()).closest("div.section,.body,.poem")[0];
	if (!begin_node)
		return;

	window.external.BeginUndoUnit(document, "add title");

	var div = document.createElement("DIV");
	div.className = "title";
	div.innerHTML = "<P>" + getTrString("Title", []) + "</P>";
	if (begin_node.firstChild)
		div = begin_node.insertBefore(div, begin_node.firstChild);
	else
		div = begin_node.appendChild(div);

	// select text in P
	rng.moveToElementText(div.firstChild);
	rng.collapse(false);
	rng.select();

	window.external.EndUndoUnit(document);
}

///<summary>Add new epigraph</summary>
function AddEpigraph() {
	if (!CanApplyStyle("epigraph"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(true);
	
	var begin_node = $(rng.parentElement()).closest("div.section,.body,.poem")[0];
	if (!begin_node)
		return;

	window.external.BeginUndoUnit(document, "add epigraph");

	var div = document.createElement("DIV");
	div.className = "epigraph";
	div.innerHTML = "<P>" + getTrString("Epigraph", []) + "</P>";
	var ep = $(begin_node).children("div.title,.epigraph");
	if (ep.length == 0)
		div = begin_node.insertBefore(div, begin_node.firstChild);
	else 
		div = begin_node.insertBefore(div, ep[ep.length-1].nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild);
	rng.collapse(false);
	rng.select();
}

///<summary>Add new epigraph</summary>
function AddPoem() {
	if (!CanApplyStyle("poem"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(false);
	var end_node = $(rng.parentElement()).closest("p")[0];

	if (!end_node)
		return;

	window.external.BeginUndoUnit(document, "add poem");

	var div = document.createElement("DIV");
	div.className = "poem";
	div.innerHTML = "<DIV class='stanza'><P>" + getTrString("Stanza",[]) + "</P></DIV>";

	div = end_node.parentElement.insertBefore(div, end_node.nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild.firstChild);
	rng.collapse(false);
	rng.select();
}

///<summary>Add new cite</summary>
function AddCite() {
	if (!CanApplyStyle("cite"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(false);
	var end_node = $(rng.parentElement()).closest("p")[0];

	if (!end_node)
		return;

	window.external.BeginUndoUnit(document, "add cite");

	var div = document.createElement("DIV");
	div.className = "cite";
	div.innerHTML = "<P>" + getTrString("Cite",[]) +"</P>";

	div = end_node.parentElement.insertBefore(div, end_node.nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild);
	rng.collapse(true);
	rng.select();
}

///<summary>Add new stanza</summary>
function AddStanza() {
	if (!CanApplyStyle("stanza"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(false);
	var end_node = $(rng.parentElement()).closest("div.stanza,div.title,p.subtitle")[0];

	if (!end_node)
		return;

	window.external.BeginUndoUnit(document, "add stanza");

	var div = document.createElement("DIV");
	div.className = "stanza";
	div.innerHTML = "<P>" + getTrString("Stanza", []) + "</P>";

	div = end_node.parentElement.insertBefore(div, end_node.nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild);
	rng.collapse(true);
	rng.select();
}

///<summary>Add new author text</summary>
function AddTextAuthor() {
	if (!CanApplyStyle("text-author"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(false);
	var parent_node = $(rng.parentElement()).closest("div.epigraph,.cite,.poem")[0];

	if (!parent_node)
		return;

	window.external.BeginUndoUnit(document, "add author text");

	var p = document.createElement("P");
	p.className = "text-author";
	p.innerText = getTrString("Author Text", []);

	p = parent_node.appendChild(p);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(p);
	rng.collapse(true);
	rng.select();
}

///<summary>Add new author text</summary>
function AddSubtitle() {
	if (!CanApplyStyle("subtitle"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(false);
	var parent_node = $(rng.parentElement()).closest("div.section,.stanza,.annotation,.cite,.history")[0];

	if (!parent_node)
		return;

	window.external.BeginUndoUnit(document, "add subtitle");

	var p = document.createElement("P");
	p.className = "subtitle";
	p.innerText = getTrString("Subtitle", []);

	p = parent_node.appendChild(p);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(p);
	rng.collapse(true);
	rng.select();
}

///<summary>Add annotation</summary>
function AddAnnotation() {
	if (!CanApplyStyle("annotation"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	rng.collapse(true);

	// can be inside one section only
	var start_range = rng.duplicate();
	start_range.collapse(true);
	var begin_node = $(rng.parentElement()).closest("div.section")[0];

	window.external.BeginUndoUnit(document, "add annotation");

	var div = document.createElement("DIV");
	div.className = "annotation";
	div.innerHTML = "<P>" + getTrString("Annotation", []) +"</P>";

	var ep = $(begin_node).children("div.title,.epigraph,.image");
	if (ep.length == 0)
		div = begin_node.insertBefore(div, begin_node.firstChild);
	else
		div = begin_node.insertBefore(div, ep[ep.length - 1].nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild);
	rng.collapse(false);
	rng.select();
}

///<summary>Add section</summary>
function AddSection() {
	if (!CanApplyStyle("annotation"))
		return;

	var rng = document.selection.createRange();
	if (!rng)
		return false;

	// can be inside one section only
	var begin_node = $(rng.parentElement()).closest("div.section")[0];

	window.external.BeginUndoUnit(document, "add annotation");

	var div = document.createElement("DIV");
	div.className = "section";
	div.innerHTML = "<DIV class='title'><P>" + getTrString("Title", []) + "</P></DIV><P>&nbsp;</P>";
	div = begin_node.parentElement.insertBefore(div, begin_node.nextSibling);

	window.external.EndUndoUnit(document);

	// select text in P
	rng.moveToElementText(div.firstChild.firstChild);
	rng.collapse(true);
	rng.select();
}

///<summary>Add new Body to HTML</summary>
function AddBody(notes) {
	window.external.BeginUndoUnit(document, "add body");
	var newbody = document.createElement("DIV");
	newbody.className = "body";
	if (notes)
	{
		newbody.setAttribute("fbname", "notes");
		newbody.innerHTML = '<DIV class="title"><P>' + getTrString('Notes', []) + '</P></DIV><DIV class="section"><P></P></DIV>';
	}
	else
		newbody.innerHTML = '<DIV class="title"><P></P></DIV><DIV class="section"><P></P></DIV>';
	var body = document.getElementById("fbw_body");
	if (!body) return false;

	body.appendChild(newbody);
	//InflateIt(newbody);
	window.external.EndUndoUnit(document);
	//GoTo(newbody);
}


///<summary>Insert image into HTML</summary>
///<param name='id'>Image id</param>
///<returns>Parent HTML-element or undefined</returns>
function InsImage(id) {
	var imgcode = "<DIV onresizestart='return false' contentEditable='false' class='image' href='#undefined'>"
	"<IMG src='" + imgPrefix + ":#undefined'></DIV>";
	var inlineimgcode = "<SPAN onresizestart='return false' contentEditable='false' class='image' href='#undefined'>" +
		"<IMG src='" + imgPrefix + "#undefined'></SPAN>";

	if (!CanApplyStyle("image"))
			return false;
	var rng = document.selection.createRange();
	if (!rng)
		return false;

	window.external.BeginUndoUnit(document, "insert image");

	var next_el = rng.parentElement().nextSibling;
	var par_el = rng.parentElement().parentElement;
	if (rng.parentElement().innerText =="") {
		//insert full image
		var img_div = document.createElement("div");
		img_div.contentEditable = false;
		img_div.className = "image";
		img_div.setAttribute("onresizestart", "return false");
		img_div.attachEvent("onresizestart", function () { return false; });
		if (id == "") {
			img_div.setAttribute("href", "#undefined");
			img_div.innerHTML = "<IMG src='" + imgPrefix + "#undefined'>";
		}
		else {
			img_div.setAttribute("href", "#" + id);
			img_div.innerHTML = "<IMG src='" + imgPrefix + "#" + id +"'>";
		}
		par_el.removeChild(rng.parentElement());
		img_div = par_el.insertBefore(img_div, next_el);
		rng.moveToElementText(img_div);
		rng.collapse(true);
		rng.select();
	}
	else {
		//insert inline image
		id == "" ? rng.pasteHTML(inlineimgcode) : rng.pasteHTML(inlineimgcode.replace(new RegExp("undefined", 'g'), id));
	}
	window.external.EndUndoUnit(document);
}

//-----------------------------------------------

///<summary>Add author text</summary>
///<param name='cp'>HTML Element/param>
function AddTA(cp, check) {
	if (!CanInsTA())
		return false;

	var sel = document.getSelection();
	if (!sel)
		return false;
	var rng = sel.getRangeAt(0).cloneRange();

	// search closest P boundaries
	var first_node = rng.startContainer;
	while (first_node && first_node.tagName != "P")
		first_node = first_node.parentNode;

	var end_node = rng.endContainer;
	while (end_node && end_node.tagName != "P")
		end_node = end_node.parentNode;

	rng.setStartBefore(first_node);
	rng.setEndAfter(end_node);
	window.external.BeginUndoUnit(document, getTrString("make text-author", []));
	var tt = rng.cloneContents();
	var ep = document.createElement("P");
	ep.className = "text-author";
	ep.innerText = rng.extractContents().textContent;

	// rng.deleteContents();
	rng.insertNode(ep);
	window.external.EndUndoUnit(document);
	/*


	cp = GetCP(cp);

	while (cp) {
		if (cp.tagName == "DIV" && (cp.className == "poem" || cp.className == "epigraph" || cp.className == "cite")) break;

		cp = cp.parentElement;
	}

	if (!cp) return;

	var lc = cp.lastChild; if (lc && lc.tagName == "P" && lc.className == "text-author") return;

	if (check) return true;

	window.external.BeginUndoUnit(document, "add text author");

	var np = document.createElement("P");
	np.className = "text-author";
	//window.external.inflateBlock(np) = true;
	cp.appendChild(np);
	window.external.EndUndoUnit(document);
	GoTo(np);*/
}

///<summary>Check if element has <P>-child</summary>
///<param name='s'>HTML Element</param>
///<returns>true if no <P> elements</returns>
function IsCtSection(s) {
	return $(s).children("P").length == 0;
	/*for(s = s.firstChild; s; s = s.nextSibling)
	  if(s.nodeName == "P") return false;
   
	return true;*/
}

///<summary>Merge two elements (this + next sibling)</summary>
///<param name='cp'>first HTML Element to merge</param>
function MergeContainers(cp, check) {
	cp = GetCP(cp); if (!cp) return;

	if (cp.className != "section" && cp.className != "stanza" && cp.className != "cite") return;

	var nx = cp.nextSibling;

	if (!nx || nx.tagName != "DIV" || nx.className != cp.className) return;

	if (check) return true;

	window.external.BeginUndoUnit(document, "merge " + cp.className + "s");

	if (cp.className == "cite") {
		for (kj = cp.firstChild; kj; kj = kj.nextSibling)
			if (kj.nodeName == "P" && kj.className == "text-author") kj.removeAttribute("className");
	}

	if (!IsCtSection(cp)) {
		// delete all nested sections
		var pi, ii = nx.firstChild;

		while (ii) {
			if (ii.tagName == "DIV") {
				if (ii.className == "title") {
					var pp = ii.getElementsByTagName("P");
					for (var l = 0; l < pp.length; l++) pp.item(l).className = "subtitle";
				}
				if (ii.className == "title" || ii.className == "epigraph" || ii.className == "section" || ii.className == "annotation") {
					ii.removeNode(false);
					ii = pi ? pi.nextSibling : nx.firstChild;
					continue;
				}
			}
			pi = ii; ii = ii.nextSibling;
		}
	}
	else if (!IsCtSection(nx)) {
		// simply move nx under cp
		cp.appendChild(nx);
	}
	else {
		// move nx's content under cp

		// check if there are any header items to save
		var needSave;
		for (var ii = nx.firstChild; ii; ii = ii.nextSibling)
			if (ii.nodeName == "DIV" && (ii.className == "image" || ii.className == "epigraph" ||
				ii.className == "annotation" || ii.className == "title")) {
				needSave = true; break;
			}

		if (needSave) {
			// find a destination section for section header items
			var dst = nx.firstChild;
			while (dst) {
				if (dst.nodeName == "DIV" && dst.className == "section") break;
				dst = dst.nextSibling;
			}

			// create one
			if (!dst) {
				dst = document.createElement("DIV");
				dst.className = "section";
				var pp = document.createElement("P");
				//window.external.inflateBlock(pp) = true;
				dst.appendChild(pp);
				cp.appendChild(dst);
			}

			// move items
			var jj = dst.firstChild;
			for (; ;) {
				var ii = nx.firstChild;
				if (!ii)
					break;
				if (ii.nodeName != "DIV")
					break;
				var stop;
				switch (ii.className) {
					case "title":
						if (jj && jj.nodeName == "DIV" && jj.className == "title") {
							jj.insertAdjacentElement("afterBegin", ii);
							ii.removeNode(false);
						} else
							InsBefore(dst, jj, ii);
						break;
					case "epigraph":
						while (jj && jj.nodeName == "DIV" && (jj.className == "title" || jj.className == "epigraph"))
							jj = jj.nextSibling;
						InsBefore(dst, jj, ii);
						break;
					case "image":
						while (jj && jj.nodeName == "DIV" && (jj.className == "title" || jj.className == "epigraph"))
							jj = jj.nextSibling;
						InsBefore(dst, jj, ii);
						break;
					case "annotation":
						while (jj && jj.nodeName == "DIV" && (jj.className == "title" || jj.className == "epigraph" || jj.className == "image"))
							jj = jj.nextSibling;
						if (jj && jj.nodeName == "DIV" && jj.className == "annotation") {
							jj.insertAdjacentElement("afterBegin", ii);
							ii.removeNode(false);
						} else
							InsBefore(dst, jj, ii);
						break;
					default:
						stop = true;
						break;
				}
				if (stop)
					break;
			}
		}

		// finally merge
		cp.appendChild(nx);
		nx.removeNode(false);
	}

	cp.insertAdjacentElement("beforeEnd", nx);
	nx.removeNode(false);

	window.external.EndUndoUnit(document);
}

///<summary>Move all children up</summary>
///<param name='cp'>HTML Element</param>
function RemoveOuterContainer(cp, check) {
	cp = GetCP(cp); if (!cp) return;

	if ((cp.className != "section" && cp.className != "body") || !IsCtSection(cp)) return;

	if (check) return true;

	window.external.BeginUndoUnit(document, "Remove Outer Section");

	// ok, move all child sections to upper level
	while (cp.firstChild) {
		var cc = cp.firstChild;
		cc.removeNode(true);
		if (cc.nodeName == "DIV" && cc.className == "section")
			cp.insertAdjacentElement("beforeBegin", cc);
	}

	cp.removeNode(true); // delete the section itself

	window.external.EndUndoUnit(document);
}

///<summary>Set image URL</summary>
///<param name='image'>image Element</param>
///<param name='url'>URL</param>
function ImgSetURL(image, url) {
	image.style.width = null;
	image.src = imgPrefix + url;
}

///<summary>Save scroll posution</summary>
function SaveBodyScroll() {
	document.ScrollSave = document.body.scrollTop;
}

///<summary>Show\hide all SPAN with id</summary>
///<param name='id'>SPAN id</param>
///<param name='show'>if show=tru - show SPAN, or hide</param>
function ShowElement(id, show) {
	//    alert(id + "/" +show);
	//	if(!id) return;
	$ff = $("#fbw_desc SPAN#" + id).each(function () {
		$(this).css("display", show ? "block" : "none");
		window.external.DescShowElement(id, show);
	});

    /*	var desc = document.getElementById("fbw_desc");
	var spans = desc.getElementsByTagName("SPAN");
	for(var i=0; i<spans.length; i++)
	{
		if(spans[i].getAttribute("id") == id)
		{
			if(show)
				spans[i].style.display = "block";
			else
				spans[i].style.display = "none";
			window.external.DescShowElement(id, show);
		}
	}
  */
}

///<summary>Display menu and toggle particular info block</summary>
///<param name='button'>Event target</param>
function ShowElementsMenu(button) {
	var elem_id = window.external.DescShowMenu(button, window.event.screenX, window.event.screenY);
	if (elem_id) {
		ShowElement(elem_id, !window.external.GetExtendedStyle(elem_id));
	}
}

///<summary>Show image</summary>
///<param name='id'>SPAN id</param>
///<param name='show'>if show=true - Full mode, if show=false -Preview</param>
function ShowImageOfBin(binDiv, mode) {
	var src = $(binDiv).find("INPUT#id").val();
	if (!src || src == "") return;
	if (mode == "full")
		ShowFullImage(imgPrefix + "#" + src);
	else if (mode == "prev")
		ShowPrevImage(imgPrefix + "#" + src);

	/* HTML DOM variant
		var inps=binDiv.getElementsByTagName("INPUT");
		for (var i=0;i<inps.length;i++)
			if (inps[i].id=="id")
			{
						var src=inps[i].value;
						if (!src || src=="") return;
						if (mode=="full") ShowFullImage("fbw-internal:#"+src);
						else if (mode=="prev") ShowPrevImage("fbw-internal:#"+src);
			}
	*/
}

//======================================
// Internal private functions
//======================================

function MsgBox(str) {
	window.external.MsgBox(str);
}
function AskYesNo(str) {
	return window.external.AskYesNo(str);
}
function InputBox(msg, value, result) {
	result.$ = window.external.InputBox(msg, "FBE script message", value);
	return window.external.GetModalResult();
}
//--------------------------------------
// Our own, less scary error handler

function errorHandler(msg, url, lno) {
	MsgBox(getTrString("Error at line {dvalue1}:\n {dvalue2} ", [lno, msg]));
	return true;
}

function errCantLoad(xd, file) {
	MsgBox(getTrString('"{dvalue1}" loading error:\n\n{dvalue2}\nLine: {dvalue3} , Col: {dvalue4}',
		[file, xd.parseError.reason, xd.parseError.line, xd.parseError.linepos]));
}

function MyMsgWindow(pe) {
	msg = pe.outerHTML;
	var MsgWindow = window.open("Scripts/debug.htm", null, "height=680,width=1014,status=no,toolbar=no,menubar=no,location=no,scrollbars=yes,resizeble=yes");
	MsgWindow.document.body.innerText = msg;
}

///<summary>Return translated string for current interface language</summary>
///<param name='string'>Source string</param>
///<param name='words'>Array of sustitute values</param>
///<returns>String for current language</returns>
function getTrString(string, words) {
	var retString = string;
	for (var i = 0; i < words.length; i++) {
		if (APITranslation[i].eng == string) {
			if (APITranslation[i][curInterfaceLang])
				retString = APITranslation[i][curInterfaceLang];
			break;
		}
	}

	if (typeof words != 'undefined' && words.length > 0) {
		for (var i = 0; i < words.length; i++) {
			retString = retString.replace("{dvalue" + (i + 1) + "}", words[i]);
		}
	}
	return retString;
}

function ShowImageLibrary() {
	$(".thumbnails").empty();
	//		$("#thumbnails").append($("<option>").attr({"value" : "Добавить", "data-img-src" : ""}).text("Добавить"));
	$("#binobj DIV").each(function () {
		var dd = this.base64data;
		alert(bufferToUtf8(dd));
		var ss = '<li><div class="imagetile"><div class="thumbnail" style=' +
			'"background : url(data:image/jpg;base64,) no-repeat; ' +
			'background-size : contain; background-position : center;">' +
			'<input type="image" src="HTML/Images/if_view_126581.png" class="viewoverflow"/></div>' +
			'<div class="container"><input type="image" class="left" src="HTML/Images/if_save_1608823.png"/>' +
			'<input class="right" type="image" src="HTML/Images/if_Cancel_1063907.png""/>' +
			'<span title="image/png: HTML/Images/if_126581.svg" class="center">(19*19)</span></div></div></li>'
		//alert(ss);
		$(".thumbnails").append(ss);
        /*.css({
            "background" : "url(data:image/jpg;"+ $(this).prop("base64data") + " , no-repeat",
            "background-size" : "contain",
            "background-position" : "center"
        });*/
	});
	$('#dialog').dialog("open");
}