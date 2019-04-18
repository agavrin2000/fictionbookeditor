#include "stdafx.h"

#include "utils.h"

#include "ElementDescriptor.h"
#include "FBDoc.h"
#include "Settings.h"
/* Element descriptors and scripts displayed in TreeView*/

extern CSettings _Settings;

CElementDescMnr _EDMnr; // Singleton instance

///<summary>Get short text with ...</summary>
///<params name="elem">HTML-element</summary>
///<returns>Shorted text/returns>
CString GetTitleSc(const MSHTML::IHTMLElementPtr elem)
{
	CString text = elem->innerText;
	text = text.Left(20) + L"...";
	U::NormalizeInplace(text);
	return text;
}

// ----------------CElementDescriptor-----------------------------
///<summary>Init element descriptor</summary>
///<params name="fn1">Pointer to function to check type</params>
///<params name="fn2">Pointer to function to get title</params>
///<params name="imageID">Image ID</params>
///<params name="default_state">Treeview default state</params>
///<params name="caption">Display name in TreeView</returns>
void CElementDescriptor::Init(FN_IsMe fn1, FN_GetTitle fn2, int imageID, bool default_state, CString caption)
{
	// Init standard element descriptor by default
    m_script = false; 
	m_title_from_script = false;

	m_getTitle = fn2;
	m_isMe = fn1;
	m_imageID = imageID;
	m_viewInTree = _Settings.GetDocTreeItemState(caption, default_state);
	m_caption = caption;
}

///<summary>Load script element descriptor</summary>
///<params name="path">Path to script file</params>
///<params name="fn2">Pointer to function to get title</summary>
///<params name="imageID">Image ID</summary>
///<params name="default_state">Treeview default state</summary>
///<returns>true on success</returns>
bool CElementDescriptor::Load(CString path)
{
	m_script = true;
	CString noExt = path.Left(path.ReverseFind(L'.'));
	m_caption = noExt.Right(noExt.GetLength() - noExt.ReverseFind(L'\\') - 1);
	CString folder = path.Left(path.ReverseFind((L'\\'))+ 1);
	m_script_path = path;
	
    m_viewInTree = false;
	m_imageID = 0;
	m_class_name = AskClassName();
	m_getTitle = (FN_GetTitle)GetTitleSc;
	m_picture = 0;
	m_imageID = 9;

	if(m_class_name.IsEmpty())
		return false;
	
	WIN32_FIND_DATA picFd;
	wchar_t* picName = new wchar_t[path.GetLength() + 1];
	wcscpy(picName, m_caption);
    
    //search and load icon for file
	CString picPathNoExt = (folder + picName);
	HANDLE hPicture = FindFirstFile(picPathNoExt + L".bmp", &picFd);
	HANDLE hIcon = FindFirstFile(picPathNoExt + L".ico", &picFd);

	if(hPicture != INVALID_HANDLE_VALUE)
	{
		HBITMAP bitmap = (HBITMAP)LoadImage(NULL, (picPathNoExt + L".bmp").GetBuffer(), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
		if(bitmap != NULL)
		{
			m_picture = bitmap;	
			m_pictType = IMG_BITMAP_TYPE;
		}							
	}
	else if(hIcon != INVALID_HANDLE_VALUE)
	{
		HICON icon = (HICON)LoadImage(NULL, (picPathNoExt + L".ico").GetBuffer(), IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
		if(icon != NULL)
		{
			m_picture = icon;	
			m_pictType = IMG_ICON_TYPE;
		}							
	}

	FindClose(hPicture);
	FindClose(hIcon);
	delete[] picName;

	return true;
}

///<summary>Get classname for current script</summary>
///<returns>Classname/returns>
CString CElementDescriptor::AskClassName()
{
	FB::Doc::m_active_doc->LoadScript(bstr_t(m_script_path));
	CComVariant vtResult;
	CComVariant params(m_script_path);	
	FB::Doc::m_active_doc->InvokeFunc(L"apiGetClassName", &params, 1, vtResult);	
	return vtResult.bstrVal;
}

///<summary>Get view state in TreeView</summary>
///<returns>View state</returns>
bool CElementDescriptor::ViewInTree()
{
	return m_viewInTree;
}

///<summary>Get image ID</summary>
///<returns>Image ID</returns>
int CElementDescriptor::GetDTImageID()
{
	return m_imageID;
}

///<summary>Check has element descriptor</summary>
///<params name="elem">HTML-element</params>
///<returns>true if element has</returns>
bool CElementDescriptor::IsMe(const MSHTML::IHTMLElementPtr elem)
{
	if(m_script)
	{
		if(U::scmp(elem->className, m_class_name) == 0)
			return true;
		return false;
	}
	else
	{
		return m_isMe(elem);
	}
}

///<summary>Get Title for element</summary>
///<params name="elem">HTML-element</params>
///<returns>Title from script, or shorted element text</returns>
CString CElementDescriptor::GetTitle(const MSHTML::IHTMLElementPtr elem)
{
	if(m_script)
	{
		// ??????? ?????? title ????? ??????
		FB::Doc::m_active_doc->LoadScript(bstr_t(m_script_path));
		CComVariant vtResult;
		CComVariant params(m_script_path);
		FB::Doc::m_active_doc->InvokeFunc(L"apiGetTitle", &params, 1, vtResult);
        CString title(vtResult.bstrVal);
        
		if(!title.IsEmpty())
			return title;
	}
	return m_getTitle(elem);
}

///<summary>Get Caption</summary>
///<returns>caption</returns>
CString CElementDescriptor::GetCaption()
{
	return m_caption;
}

///<summary>Set view state in TreeView for standard descriptors.
/// Script descriptors are diaplayed allways</summary>
///<params name="view">view state</params>
void CElementDescriptor::SetViewInTree(bool view)
{
	m_viewInTree = view;
	if(!m_script)
		_Settings.SetDocTreeItemState(m_caption, view);
}

///<summary>Process Script</summary>
void CElementDescriptor::ProcessScript()
{
	if(!m_script)
		return;

	CComVariant vtResult;
	CComVariant params(m_script_path);	
	FB::Doc::m_active_doc->InvokeFunc(L"apiProcessCmd", &params, 1, vtResult);
}

///<summary>Remove standard descriptor from TreeView</summary>
void CElementDescriptor::CleanUp()
{
	if(!m_script)
		return;

	CComVariant vtResult;
	CComVariant params(m_class_name);	
	FB::Doc::m_active_doc->InvokeFunc(L"apiCleanUp", &params, 1, vtResult);
}

///<summary>Remove standard descriptor from TreeView</summary>
bool CElementDescriptor::GetPic(HANDLE& handle, int& type)
{
	if(!m_picture)
		return false;
	handle = m_picture;
	type = m_pictType;
	return true;
}

///<summary>Set image ID</summary>
///<params name="view">image ID</params>
void CElementDescriptor::SetImageID(int ID)
{
	m_imageID = ID;
}

//------------------Functions Set for checking type element descriptors and getting titles-----
bool IsDiv(MSHTML::IHTMLElementPtr elem, CString className)
{
	if(U::scmp(elem->tagName, L"DIV") == 0 && U::scmp(elem->className, className) == 0)
		return true;

	return false;
}

bool IsP(MSHTML::IHTMLElementPtr elem, CString className)
{
	if(U::scmp(elem->tagName, L"P") == 0 && U::scmp(elem->className, className) == 0)
		return true;

	return false;
}

bool IsStylesheet(MSHTML::IHTMLElementPtr elem)
{
	return (U::scmp(elem->className, L"stylesheet") == 0);
}

bool IsStyle(MSHTML::IHTMLElementPtr elem)
{
	CString outerHTML = elem->outerHTML;
	outerHTML.MakeLower();
	return (outerHTML.Find(L"<span class=") == 0);
}

CString GetStylesheetTitle(const MSHTML::IHTMLElementPtr elem)
{
	return L"";
}

bool IsSection(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"section");
}

CString GetSectionTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsBody(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"body");
}

CString GetBodyTitle(const MSHTML::IHTMLElementPtr elem)
{
	CString caption;
	caption = U::GetAttrCS(elem, L"fbname");
	U::NormalizeInplace(caption);
	if (caption.IsEmpty())
		caption = U::FindTitle(elem);
	return caption;
}

bool IsImage(MSHTML::IHTMLElementPtr elem)
{
//	return IsDiv(elem, L"image");
	return (U::scmp(elem->className, L"image") == 0);
}

CString GetImageTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::GetImageFileName(elem);
}

bool IsAnnotation(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"annotation");
}

CString GetAnnotationTitle(const MSHTML::IHTMLElementPtr elem)
{
	return L"";
}

bool IsHistory(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"history");
}

CString GetHistoryTitle(const MSHTML::IHTMLElementPtr elem)
{
	return L"";
}

bool IsPoem(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"poem");
}

CString GetPoemTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsStanza(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"stanza");
}

CString GetStanzaTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsTitle(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"title");
}

CString GetTitleTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsEpigraph(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"epigraph");
}

CString GetEpigraphTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsCite(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"cite");
}

CString GetCiteTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

CString GetCodeTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsSubtitle(MSHTML::IHTMLElementPtr elem)
{
	return IsP(elem, L"subtitle");
}

CString GetSubtitleTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

bool IsTable(MSHTML::IHTMLElementPtr elem)
{
	return IsDiv(elem, L"table");
}

bool IsTH(MSHTML::IHTMLElementPtr elem)
{
	return IsP(elem, L"th");
}

bool IsTD(MSHTML::IHTMLElementPtr elem)
{
	return IsP(elem, L"td");
}

bool IsTR(MSHTML::IHTMLElementPtr elem)
{
	return IsP(elem, L"tr");
}

CString GetTableTitle(const MSHTML::IHTMLElementPtr elem)
{
	return U::FindTitle(elem);
}

//--------------- CElementDescMnr --------------------------
CElementDescMnr::CElementDescMnr()
{
}

// Init uses Settings to get information about each item in its list,
// but until Settings are deserialized, that information is reflected
// by CSettings::GetDocTreeItemState via variable which is set to
// CElementDescriptor constructor. On the other hand information about
// serialized fields and default values is collected after initiation, but not before.
// Though we have to assign really deserialized values during creation of CTreeWithToolBar
// explicitly calling information on CSettings::TREEITEMSHOWINFO instance via the same
// CSettings::GetDocTreeItemState. That is non-obvious but works!

///<summary>Init standard set of element descriptors</summary>
void CElementDescMnr::InitStandartEDs()
{
		CElementDescriptor* stylesheet = new CElementDescriptor;
		stylesheet->Init(IsStylesheet, GetStylesheetTitle, 0, false, L"Stylesheet");
		m_stEDs.Add(stylesheet);

        /*CElementDescriptor* style = new CElementDescriptor;
		style->Init(IsStyle, GetStylesheetTitle, 0, false, L"Style");
		m_stEDs.Add(style);*/

		CElementDescriptor* section = new CElementDescriptor;
		section->Init(IsSection, GetSectionTitle, 0, true, L"Section");
		m_stEDs.Add(section);
		
        CElementDescriptor* body = new CElementDescriptor;
		body->Init(IsBody, GetBodyTitle, 0, true, L"Body");
		m_stEDs.Add(body);
		
        CElementDescriptor* image = new CElementDescriptor;
		image->Init(IsImage, GetImageTitle, 21, true, L"Image");
		m_stEDs.Add(image);
		
        CElementDescriptor* annotation = new CElementDescriptor;
		annotation->Init(IsAnnotation, GetAnnotationTitle, 12, true, L"Annotation");
		m_stEDs.Add(annotation);
		
        CElementDescriptor* history = new CElementDescriptor;
		history->Init(IsHistory, GetHistoryTitle, 15, true, L"History");
		m_stEDs.Add(history);
		
        CElementDescriptor* poem = new CElementDescriptor;
		poem->Init(IsPoem, GetPoemTitle, 3, true, L"Poem");
		m_stEDs.Add(poem);
		
        CElementDescriptor* stanza = new CElementDescriptor;
		stanza->Init(IsStanza, GetStanzaTitle, 3, true, L"Stanza");
		m_stEDs.Add(stanza);
		
        CElementDescriptor* title = new CElementDescriptor;
		title->Init(IsTitle, GetTitleTitle, 27, false, L"Title");
		m_stEDs.Add(title);

		CElementDescriptor* epigraph = new CElementDescriptor;
		epigraph->Init(IsEpigraph, GetEpigraphTitle, 18, false, L"Epigraph");
		m_stEDs.Add(epigraph);

		CElementDescriptor* cite = new CElementDescriptor;
		cite->Init(IsCite, GetCiteTitle, 9, false, L"Cite");
		m_stEDs.Add(cite);

		CElementDescriptor* subtitle = new CElementDescriptor;
		subtitle->Init(IsSubtitle, GetSubtitleTitle, 6, true, L"Subtitle");
		m_stEDs.Add(subtitle);
        
		CElementDescriptor* table = new CElementDescriptor;
		table->Init(IsTable, GetTableTitle, 30, true, L"Table");
		m_stEDs.Add(table);
        
		CElementDescriptor* th = new CElementDescriptor;
		th->Init(IsTH, GetTableTitle, 27, true, L"th");
		m_stEDs.Add(th);
        
		CElementDescriptor* td = new CElementDescriptor;
		td->Init(IsTD, GetTableTitle, 27, true, L"td");
		m_stEDs.Add(td);
        
		CElementDescriptor* tr = new CElementDescriptor;
		tr->Init(IsTR, GetTableTitle, 27, true, L"tr");
		m_stEDs.Add(tr);
}

///<summary>Init set of script element descriptors</summary>
///<params name="askname">true=show FileSave dialog</summary>
///<returns>FILE_OP_STATUS</summary>
void CElementDescMnr::InitScriptEDs()
{
	CElementDescriptor* ED;

	WIN32_FIND_DATA fd;
	CString fff = U::GetDocTReeScriptsDir();
	HANDLE found = FindFirstFile(U::GetDocTReeScriptsDir() + L"\\*.js", &fd);
	if(INVALID_HANDLE_VALUE != found)
	{
		do
		{
			ED = new CElementDescriptor;
			if(ED->Load(U::GetDocTReeScriptsDir() + L"\\" + fd.cFileName))
			{
				m_EDs.Add(ED);
			}
			else
			{
				delete ED;
			}
		}
		while(FindNextFile(found, &fd));
	}
}

///<summary>Get element descriptor for html-element</summary>
///<params name="elem">HTML-element</summary>
///<params name="desc">Output: Pointer to element descriptor</summary>
///<returns>true on success, false if not found</summary>
bool CElementDescMnr::GetElementDescriptor(const MSHTML::IHTMLElementPtr elem, CElementDescriptor **desc) const
{
	// search in standard element descriptors
    for(int i = 0; i < m_stEDs.GetSize(); i++)
	{
		if(m_stEDs[i]->IsMe(elem))
		{
			*desc = m_stEDs[i];
			return true;
		}
	}

	// search in script element descriptors
	for(int i = 0; i < m_EDs.GetSize(); i++)
	{
		if(m_EDs[i]->IsMe(elem))
		{
			*desc = m_EDs[i];
			return true;
		}
	}
	return false;
}

///<summary>Get standard element descriptor by index</summary>
///<params name="index">HTML-element</summary>
///<returns>element descriptor</summary>
CElementDescriptor* CElementDescMnr::GetStED(int index)
{
	return m_stEDs[index];
}

///<summary>Get script element descriptor by index</summary>
///<params name="index">HTML-element</summary>
///<returns>element descriptor</summary>
CElementDescriptor* CElementDescMnr::GetED(int index)
{
	return m_EDs[index];
}

///<summary>Get array size of standard element descriptors</summary>
///<returns>Array size</summary>
int CElementDescMnr::GetEDsCount()
{
	return m_EDs.GetSize();
}

///<summary>Get array size of script element descriptors</summary>
///<returns>Array size</summary>
int CElementDescMnr::GetStEDsCount()
{
	return m_stEDs.GetSize();
}

///<summary>Clear array of scriptelement descriptors</summary>
void CElementDescMnr::CleanUpAll()
{
	for(int i = 0; i < m_EDs.GetSize(); i++)
	{
		m_EDs[i]->CleanUp();
	}
}

///<summary>Clear array of script element descriptors
/// and uncheck ALL (standard and script) in TreeView</summary>
void CElementDescMnr::CleanTree()
{
	for(int i = 0; i < m_EDs.GetSize(); i++)
	{
		m_EDs[i]->CleanUp();
		m_EDs[i]->SetViewInTree(false);
	}
	for(int i = 0; i < m_stEDs.GetSize(); i++)
	{
		m_stEDs[i]->SetViewInTree(false);
	}
}
