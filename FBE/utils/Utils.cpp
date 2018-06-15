#include "stdafx.h"

#include <deque>
#include "utils.h"
#include "../ErrsEx.h"
#include "../Settings.h"
#include "../resource.h"
#include "../res1.h"

#include <iostream>
#include <sstream>

#ifndef NO_EXTRN_SETTINGS
extern CSettings _Settings;
#endif

bool VBErr = false;

namespace U
{
	CString g_strProductInfo;

	CSimpleMap<CString, WORD> keycodes;

	void ChangeAttribute(MSHTML::IHTMLElementPtr elem, const wchar_t* attrib, const wchar_t* value)
	{
		if(U::scmp(attrib, L"class")) elem->setAttribute(attrib, _variant_t(value), 1);
		else elem->className = value;
		// send message to main window
#ifndef NO_EXTRN_SETTINGS
		::SendMessage(_Settings.GetMainWindow(), WM_COMMAND, MAKEWPARAM(0,IDN_TREE_RESTORE), 0);
#endif
	}

HandleStreamPtr	  NewStream(HANDLE& hf,bool fClose) {
  HandleStream  *hs;
  CheckError(HandleStream::CreateInstance(&hs));
  hs->SetHandle(hf,fClose);
  if (fClose)
    hf=INVALID_HANDLE_VALUE;
  return hs;
}

void  NormalizeInplace(CString& s) {
  int	  len=s.GetLength();
  TCHAR	  *p=s.GetBuffer(len);
  TCHAR	  *r=p;
  TCHAR	  *q=p;
  TCHAR	  *e=p+len;
  int	  state=0;

  while (p<e) {
    switch (state) {
    case 0:
      if ((unsigned)*p > 32) {
	*q++=*p;
	state=1;
      }
      break;
    case 1:
      if ((unsigned)*p > 32)
	*q++=*p;
      else
	state=2;
      break;
    case 2:
      if ((unsigned)*p > 32) {
	*q++=_T(' ');
	*q++=*p;
	state=1;
      }
      break;
    }
    ++p;
  }
  s.ReleaseBuffer(q-r);
}

void  RemoveSpaces(wchar_t *zstr) {
  wchar_t  *q=zstr;

  while (*zstr) {
    if (*zstr > 32)
      *q++=*zstr;
    ++zstr;
  }
  *q=L'\0';
}

int	scmp(const wchar_t *s1,const wchar_t *s2) {
  bool f1=!s1 || !*s1;
  bool f2=!s2 || !*s2;
  if (f1 && f2)
    return 0;
  if (f1)
    return -1;
  if (f2)
    return 1;
  return wcscmp(s1,s2);
}

CString GetMimeType(const CString& filename)
{
	CString fn(filename);
	int cp = fn.ReverseFind(_T('.'));
	if(cp < 0)
os:
		return _T("application/octet-stream");
	fn.Delete(0, cp);

	CRegKey rk;
	if(rk.Open(HKEY_CLASSES_ROOT, fn , KEY_READ) != ERROR_SUCCESS)
		goto os;

	CString ret;
	ULONG len = 128;
	TCHAR* rbuf = ret.GetBuffer(len);
	rbuf[0] = _T('\0');
	LONG rv = rk.QueryStringValue(_T("Content Type"), rbuf, &len);
	ret.ReleaseBuffer();

	if(rv != ERROR_SUCCESS)
		goto os;

	return ret;
}
#ifndef NO_EXTRN_SETTINGS
bool is_whitespace(const wchar_t *spc) {
  wchar_t nbsp = _Settings.GetNBSPChar()[0];
  while (*spc) {
    if (!iswspace(*spc) && *spc!=nbsp)
      return false;
    ++spc;
  }
  return true;
}
#endif
CString	GetFileTitle(const TCHAR *filename) {
  CString   ret;
  TCHAR	    *buf=ret.GetBuffer(MAX_PATH);
  if (::GetFileTitle(filename,buf,MAX_PATH))
    ret.ReleaseBuffer(0);
  else
    ret.ReleaseBuffer();
  return ret;
}

DWORD QueryIV(HKEY hKey,const TCHAR *name,DWORD defval) {
  DWORD	dw;
  DWORD	len=sizeof(DWORD);
  DWORD	type=REG_DWORD;
  if (::RegQueryValueEx(hKey,name,NULL,&type,(BYTE*)&dw,&len)!=ERROR_SUCCESS ||
      type!=REG_DWORD || len!=sizeof(DWORD))
    return defval;
  return dw;
}

CString	QuerySV(HKEY hKey, const TCHAR *name, const TCHAR *def) {
  CString   ret;
  DWORD	    type=REG_SZ;
  DWORD	    len=0;
  if (::RegQueryValueEx(hKey,name,NULL,&type,NULL,&len)!=ERROR_SUCCESS ||
      (type!=REG_SZ && type!=REG_EXPAND_SZ))
    return def ? def : CString();
  TCHAR	    *cp=ret.GetBuffer(len/sizeof(TCHAR));
  if (::RegQueryValueEx(hKey,name,NULL,&type,(BYTE*)cp,&len)!=ERROR_SUCCESS) {
    ret.ReleaseBuffer(0);
    return def ? def : CString();
  }
  ret.ReleaseBuffer();
  return ret;
}

HFONT	  CreatePtFont(int sizept,const TCHAR *facename,bool fBold,bool fItalic)
{
  HDC	  hDC=CreateDC(_T("DISPLAY"),NULL,NULL,NULL);
  HFONT	  hFont=::CreateFont(
    -MulDiv(sizept, GetDeviceCaps(hDC, LOGPIXELSY), 72),
    0,
    0,
    0,
    fBold ? FW_BOLD : FW_NORMAL,
    fItalic,
    0,
    0,
    DEFAULT_CHARSET,
    OUT_DEFAULT_PRECIS,
    CLIP_DEFAULT_PRECIS,
    DEFAULT_QUALITY,
    DEFAULT_PITCH|FF_DONTCARE,
    facename);
  DeleteDC(hDC);
  return hFont;
}

CString	GetCharName(int ch) {
  CString   num;
  num.Format(_T("U+%04X"),ch);

#if 0
  static bool	  fTriedDll;
  static HMODULE  hDll;
  if (!hDll) {
    if (fTriedDll)
      return num;
    hDll=LoadLibrary(_T("getuname.dll"));
    fTriedDll=true;
    if (!hDll)
      return num;
  }
  int	    cur=num.GetLength();
  TCHAR	    *buf=num.GetBuffer(cur+101);
  int	    nch=LoadString(hDll,ch,buf+cur+1,100);
  if (nch) {
    buf[cur]=' ';
    num.ReleaseBuffer(cur+1+nch);
  } else
    num.ReleaseBuffer(cur);
#endif
  return num;
}

// path names
CString	GetFullPathName(const CString& filename) {
  CString   ret;
  int	    len=MAX_PATH;
  for (;;) {
    TCHAR	    *buf=ret.GetBuffer(len);
    TCHAR	    *final;
    int nlen=::GetFullPathName(filename,len,buf,&final);
    if (nlen==0) // failed
      return filename;
    if (nlen>len) {
      ret.ReleaseBuffer(0);
      len=nlen;
    } else {
      ret.ReleaseBuffer(nlen);
      break;
    }
  }

  typedef DWORD	(__stdcall *GLPN)(LPCTSTR,LPTSTR,DWORD);
  static bool	checked=false;
  static GLPN	glpn;

  if (!checked) {
    HMODULE hDll=::GetModuleHandle(_T("kernel32.dll"));
    if (hDll) {
      glpn=(GLPN)::GetProcAddress(hDll,"GetLongPathName"
#ifdef UNICODE
	"W"
#else
	"A"
#endif
	);
    }
    checked=true;
  }

  if (glpn) {
    CString   r2;
    for (;;) {
      TCHAR   *buf=r2.GetBuffer(len);
      int	nlen=glpn(ret,buf,len);
      if (nlen==0)
	return ret;
      if (nlen>len) {
	r2.ReleaseBuffer(0);
	len=nlen;
      } else {
	r2.ReleaseBuffer(nlen);
	return r2;
      }
    }
  } else
    return ret;
}

CString	UrlFromPath(const CString& path) {
  CString   ret;
  DWORD	    size=MAX_PATH*3;
  TCHAR	    *cp=ret.GetBuffer(size);
  if (FAILED(UrlCreateFromPath(path,cp,&size,0)))
    return CString();
  ret.ReleaseBuffer(size);
  ret.FreeExtra();
  return ret;
}

CString	Win32ErrMsg(DWORD code) {
  CString ret;
  TCHAR	  *buf=ret.GetBuffer(1024);
  int len=::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM,NULL,code,0,buf,1024,NULL);
  ret.ReleaseBuffer(len);
  if (len==0)
    ret.Format(_T("Unknown error %x"),code);
  return ret;
}

//
// added by SeNS
// fix Vista/Windows 7 issues
CString GetSettingsDir()
{
	wchar_t szPath[MAX_PATH];
	if ( SUCCEEDED( SHGetFolderPath( NULL, CSIDL_LOCAL_APPDATA, NULL, 0, szPath ) ) )
	{
		CString appSettingPath(szPath);
		appSettingPath += L"\\FBE\\";
		::CreateDirectory(appSettingPath, NULL);
		return appSettingPath; 
	}
	else return GetProgDir();
}

CString	GetProgDir() {
  CString     exedir;
  TCHAR	      *cp=exedir.GetBuffer(MAX_PATH);
  DWORD len=::GetModuleFileName(_Module.GetModuleInstance(),cp,MAX_PATH);
  exedir.ReleaseBuffer(len);
  int	      p=len;
  while (p>0 && exedir[p-1]!=_T('\\'))
    --p;
  exedir.Delete(p,len-p);
  return exedir;
}

CString GetDocTReeScriptsDir()
{
	return GetProgDir() + L"TreeCmd";
}

CString	GetProgDirFile(const CString& filename) {
  CString     exedir;
  TCHAR	      *cp=exedir.GetBuffer(MAX_PATH);
  DWORD len=::GetModuleFileName(_Module.GetModuleInstance(),cp,MAX_PATH);
  exedir.ReleaseBuffer(len);
  int	      p=len;
  while (p>0 && exedir[p-1]!=_T('\\'))
    --p;
  exedir.Delete(p,len-p);
  CString     tryname(exedir+filename);
  DWORD	      attr=::GetFileAttributes(tryname);
  if (attr!=INVALID_FILE_ATTRIBUTES)
    goto ok;
  tryname=exedir+_T("..\\")+filename;
  attr=::GetFileAttributes(tryname);
  if (attr!=INVALID_FILE_ATTRIBUTES)
    goto ok;
  tryname=filename;
ok:
  return GetFullPathName(tryname);
}

void  ReportError(HRESULT hr)
{
	CString err(Win32ErrMsg(hr));
	AtlTaskDialog(::GetActiveWindow(), IDS_ERRMSGBOX_CAPTION, (LPCTSTR)err, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
}

void  ReportError(_com_error& e) {
  CString   err;
  HRESULT errCode = e.Error();
  // search for the VBScript error
  for (int i=0; i<(sizeof(vberrs)/sizeof(VB_ERRORS)); i++)
	  if (vberrs[i].Error == errCode)
	  {
		  err.Format(_T("Code: %08lx"),vberrs[i].Error);
		  _bstr_t src(vberrs[i].Source);
		  if (src.length()>0) 
		  {
			err+=_T("\nSource: ");
			err+=(const wchar_t *)src;
		  }
		  src=vberrs[i].Description;
	      if (src.length()>0) 
		  {
		    err+=_T("\nDescription: ");
		    err+=(const wchar_t *)src;
	      }
		  break;
	  }

  if (err.IsEmpty())
  {
	  err.Format(_T("Code: %08lx [%s]"),e.Error(),e.ErrorMessage());
	  _bstr_t src(e.Source());
	  if (src.length()>0) {
		err+=_T("\nSource: ");
		err+=(const wchar_t *)src;
	  }
	  src=e.Description();
	  if (src.length()>0) {
		err+=_T("\nDescription: ");
		err+=(const wchar_t *)src;
	  }
  }
  AtlTaskDialog(::GetActiveWindow(), IDS_COM_ERR_CPT, (LPCTSTR)err, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
  VBErr = true;
}

CString	GetWindowText(HWND hWnd)
{
	CString ret;
	CWindow wnd(hWnd);
	wnd.GetWindowText(ret);
	return ret;
}

CString	GetCBString(HWND hWnd,int idx) {
  int len=ComboBox_GetLBTextLen(hWnd,idx);
  if (len==CB_ERR)
    return CString();
  CString ret;
  TCHAR	  *cp=ret.GetBuffer(len+1);
  len=ComboBox_GetLBText(hWnd,idx,cp);
  if (len==CB_ERR)
    ret.ReleaseBuffer(0);
  else
    ret.ReleaseBuffer(len);
  return ret;
}

MSXML2::IXMLDOMDocument2Ptr  CreateDocument(bool fFreeThreaded)
{
  MSXML2::IXMLDOMDocument2Ptr  doc;
  wchar_t		      *cls=fFreeThreaded ?
    L"Msxml2.FreeThreadedDOMDocument.6.0" : L"Msxml2.DOMDocument.6.0";
  CheckError(doc.CreateInstance(cls));
  return doc;
}

MSXML2::IXSLTemplatePtr    CreateTemplate() {
  MSXML2::IXSLTemplatePtr    tp;
  CheckError(tp.CreateInstance(L"Msxml2.XSLTemplate.6.0"));
  return tp;
}

void  ReportParseError(MSXML2::IXMLDOMDocument2Ptr doc)
{
  try 
  {
    MSXML2::IXMLDOMParseErrorPtr err(doc->parseError);
    long	  line=err->line;
    long	  col=err->linepos;
    _bstr_t	  url(err->url);
    _bstr_t	  reason(err->reason);
    CString msg;
    if (line && col)
	{
		msg.Format(IDS_XML_PARSE_ERR_MSG, (LPCTSTR)url, line, col, (LPCTSTR)reason);
		AtlTaskDialog(::GetActiveWindow(), IDS_XML_PARSE_ERR_CPT, (LPCTSTR)msg, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
    }
    else
	{
		msg.Format(IDS_XML_PARSE_ERRQ_MSG, (LPCTSTR)url, (LPCTSTR)reason);
		AtlTaskDialog(::GetActiveWindow(), IDS_XML_PARSE_ERR_CPT, (LPCTSTR)msg, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
    }
  }
  catch (_com_error& e) 
  {
    ReportError(e);
  }
}

bool  LoadXml(MSXML2::IXMLDOMDocument2Ptr doc,const CString& url)
{
  doc->put_async(VARIANT_FALSE);
  _variant_t  vturl((const TCHAR *)url);
  VARIANT_BOOL	flag;
  HRESULT   hr=doc->raw_load(vturl,&flag);
  if (FAILED(hr)) {
    ReportError(hr);
    return false;
  }
  if (flag!=VARIANT_TRUE) {
    ReportParseError(doc);
    return false;
  }
  return true;
}

HRESULT LoadFile(const TCHAR* filename , VARIANT* vt)
{
	HRESULT hr;

	HANDLE hFile = ::CreateFile(filename,
								GENERIC_READ,
								FILE_SHARE_READ,
								NULL,
								OPEN_EXISTING,
								FILE_FLAG_SEQUENTIAL_SCAN,
								NULL);

	if(hFile == INVALID_HANDLE_VALUE)
		return HRESULT_FROM_WIN32(::GetLastError());

	DWORD fsz = ::GetFileSize(hFile,NULL);
	if(fsz == INVALID_FILE_SIZE)
	{
		hr = HRESULT_FROM_WIN32(::GetLastError());
cfexit:
		CloseHandle(hFile);
		return hr;
	}

	SAFEARRAY* sa = ::SafeArrayCreateVector(VT_UI1, 0, fsz);
	if(sa == NULL)
	{
		hr = E_OUTOFMEMORY;
		goto cfexit;
	}

	void* pv;
	if(FAILED(hr = ::SafeArrayAccessData(sa, &pv)))
	{
		::SafeArrayDestroy(sa);
		goto cfexit;
	}

	DWORD nrd;
	BOOL fRd = ::ReadFile(hFile,pv,fsz,&nrd,NULL);
	DWORD err = ::GetLastError();

	::SafeArrayUnaccessData(sa);
	::CloseHandle(hFile);

	if(!fRd)
	{
		::SafeArrayDestroy(sa);
		return HRESULT_FROM_WIN32(err);
	}

	if(nrd != fsz)
	{
		::SafeArrayDestroy(sa);
		return E_FAIL;
	}

	V_VT(vt) = VT_ARRAY | VT_UI1;
	V_ARRAY(vt) = sa;

	return S_OK;
}

bool GetImageDimsByPath(const wchar_t* pszFileName, int* nWidth, int* nHeight)
{
	CImage image;
	bool result = false;

	if(image.Load(pszFileName) == S_OK)
	{
		if(nWidth)
			*nWidth = image.GetWidth();
		if(nHeight)
			*nHeight = image.GetHeight();
		result = true;
	}

	return result;
}

bool GetImageDimsByData(SAFEARRAY* data, ULONG /*length*/, int* nWidth, int* nHeight)
{
	CImage image;
	bool result = false;

	long elnum = 0;
	void* pData;

	if(::SafeArrayPtrOfIndex(data, &elnum, &pData) != S_OK)
		return false;
	
	IStream* pIStream;
	if(S_OK != ::CreateStreamOnHGlobal(pData, FALSE, &pIStream))
		return false;

	if(image.Load(pIStream) == S_OK)
	{
		if(nWidth)
			*nWidth = image.GetWidth();
		if(nHeight)
			*nHeight = image.GetHeight();
		result = true;
	}

	pIStream->Release();

	return result;
}

WORD StringToKeycode(CString string)
{
	int key = keycodes.FindKey(string);
	if(key >= 0)
		return keycodes.GetValueAt(key);
	else return NULL;
}

CString KeycodeToString(WORD keycode)
{
	int key = keycodes.FindVal(keycode);
	if(key >= 0)
		return keycodes.GetKeyAt(key);
	else return CString();
}

CString AccelToString(ACCEL accel)
{
	CString temp;
	if(accel.fVirt & FCONTROL)
		temp += "Ctrl+";
	if(accel.fVirt & FALT)
		temp += "Alt+";
	if(accel.fVirt & FSHIFT)
		temp += "Shift+";

	temp += KeycodeToString(accel.key);

	return temp;
}

WORD VKToFVirt(WORD virtkey)
{
	switch(virtkey)
	{
	case 0x12: // ALT keycode
		return FALT;
	case VK_SHIFT:
		return FSHIFT;
	case VK_CONTROL:
		return FCONTROL;
	}

	return NULL;
}

void InitSettings() 
{
#ifndef NO_EXTRN_SETTINGS
	_Settings.Init();
	_Settings.Load();

	_Settings.LoadWords();
#endif
}

// DomPath

	bool DomPath::CreatePathFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root, MSHTML::IHTMLDOMNodePtr endNode)
	{
		// лепим путь к элементу начиная от него и раскручивая до корня
		m_path.clear();
		if(!(bool)root || !(bool)endNode)
		{
			return false;
		}
		if(endNode == root)
		{
			m_path.push_front(0);
			return true;
		}

		int id = 0;
		MSHTML::IHTMLDOMNodePtr parentNode = endNode->parentNode;
		MSHTML::IHTMLDOMNodePtr currentNode = endNode;

		if(U::scmp(currentNode->nodeName, L"#text") == 0)
		{
			currentNode = currentNode->previousSibling;
		}
		
		while(currentNode)
		{
			if(currentNode == root)
			{
				return true;
			}
			MSHTML::IHTMLDOMNodePtr previousSibling = currentNode->previousSibling;
			while(previousSibling)
			{
				if(U::scmp(previousSibling->nodeName, L"#text") != 0)
				{
					++id;
				}
				previousSibling = previousSibling->previousSibling;
			}
			m_path.push_front(id);
			id = 0;
			currentNode = currentNode->parentNode;
		}
		m_path.clear();
		return false;
	}

	HRESULT DomPath::GetNodeFromXMLDOM(MSXML2::IXMLDOMNode * pRoot, MSXML2::IXMLDOMNode ** ppNode)
	{
		if (!ppNode)
			return E_POINTER;

		*ppNode = nullptr;

		if (!pRoot)
			return E_INVALIDARG;

		if (m_path.empty())
			return E_NOT_VALID_STATE;
		
		CComPtr<MSXML2::IXMLDOMNode> currentNode(pRoot);
		
		HRESULT hr = E_FAIL;

		size_t size = m_path.size();
	    for (size_t i = 0; i < size; ++i)
		{
			CComPtr<MSXML2::IXMLDOMNode> firstChild;
			 hr = currentNode->get_firstChild(&firstChild);
			 if (hr != S_OK)
			 {
				 return E_FAIL;
			 }

			 currentNode = firstChild;
			 firstChild.Release();

			 CComBSTR nodeName;
			 hr = currentNode->get_nodeName(&nodeName);
			 if (SUCCEEDED(hr))
			 {
				 if (U::scmp(currentNode->nodeName, L"#text") == 0)
				 {
					 CComPtr<MSXML2::IXMLDOMNode> nextSibling;
					 hr = currentNode->get_nextSibling(&nextSibling);
					 if (hr != S_OK)
					 {
						 return E_FAIL;
					 }

					 currentNode = nextSibling;

				 }
				 nodeName.Empty();

				 int numb = m_path[i];
				 for (int j = 0; j < numb; ++j)
				 {
					 CComPtr<MSXML2::IXMLDOMNode> nextSibling;
					 hr = currentNode->get_nextSibling(&nextSibling);
					 if (hr != S_OK)
					 {
						 return E_FAIL;
					 }

					 currentNode = nextSibling;
					 nextSibling.Release();

					 hr = currentNode->get_nodeName(&nodeName);
					 if (SUCCEEDED(hr))
					 {
						 if (U::scmp(nodeName, L"#text") == 0)
						 {
							 hr = currentNode->get_nextSibling(&nextSibling);
							 if (hr = S_OK)
							 {
								 currentNode = nextSibling;
								 nextSibling.Release();
							 }
						 }
						 nodeName.Empty();
					 }
				 }
			 }
	    }

	    return currentNode.QueryInterface(ppNode);
	}

	bool DomPath::CreatePathFromXMLDOM(MSXML2::IXMLDOMNodePtr root, MSXML2::IXMLDOMNodePtr endNode)
	{
		// лепим путь к элементу начиная от него и раскручивая до корня
		m_path.clear();
		if(!(bool)root || !(bool)endNode)
		{
			return false;
		}
		if(endNode == root)
		{
			m_path.push_front(0);
			return true;
		}
		int id = 0;
		MSXML2::IXMLDOMNodePtr parentNode = endNode->parentNode;
		MSXML2::IXMLDOMNodePtr currentNode = endNode;

		if(U::scmp(currentNode->nodeName, L"#text") == 0)
		{
			currentNode = currentNode->previousSibling;
		}
		
		while(currentNode)
		{
			if(currentNode == root)
			{
				return true;
			}
			MSXML2::IXMLDOMNodePtr previousSibling = currentNode->previousSibling;			
			while(previousSibling)
			{
				if(U::scmp(previousSibling->nodeName, L"#text") != 0)
				{
					++id;
				}				
				previousSibling = previousSibling->previousSibling;				
			}
			m_path.push_front(id);
			id = 0;
			currentNode = currentNode->parentNode;			
		}
		m_path.clear();
		return false;
	}

	MSHTML::IHTMLDOMNodePtr DomPath::GetNodeFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root)
	{
		if(!(bool)root || m_path.empty())
		{
			return 0;		
		}
		
		MSHTML::IHTMLDOMNodePtr currentNode = root;
		
		int size = m_path.size();
		for(int i = 0; i < size; ++i)
		{
			currentNode = currentNode->firstChild;
			if (currentNode) 
				if(U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;

			//name = currentNode->nodeName;
			if(!currentNode) return 0;

			int numb = m_path[i];
			for(int j = 0; j < numb; ++j)
			{
				currentNode = currentNode->nextSibling;
				if(!currentNode) return 0;

				if(U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;
				if(!currentNode) return 0;
			}			
		}
		return currentNode;
	}

	bool DomPath::CPFT(const wchar_t* xml, int pos, size_t* char_pos)
	{
		if(!xml)
		{
			return false;
		}

		const wchar_t* selpos = xml + pos;
		size_t virtual_pos = pos;
		// ищем открывающий тег
		int id = 0;

		wchar_t* open_tag_begin = 0;
		wchar_t* open_tag_end = 0;
		wchar_t* close_tag_begin = 0;
		wchar_t* close_tag_end = 0;

		bool ok = tGetFirstXMLNodeParams(xml, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);

		while(ok)
		{
			if(selpos > close_tag_end)
			{
				++id;
				ok = tGetFirstXMLNodeParams(close_tag_end, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);
				virtual_pos = virtual_pos - (close_tag_end - open_tag_begin);
				continue;
			}
			// если курсор попал внутрь тега
			if((selpos >= open_tag_begin && selpos < open_tag_end) ||(selpos > close_tag_begin && selpos <= close_tag_end))
			{
				*char_pos = 0;
				m_path.push_back(id);
				return true;
			}

			// если курсор между открывающим и закрывающим тегом, то рекурсивно заходим в элементы
			if(selpos >= open_tag_end && selpos <= close_tag_begin)
			{
				m_path.push_back(id);
				id = 0;
				return this->CPFT(open_tag_end, selpos - open_tag_end, char_pos);
			}

			// если открывающий тег дальше курсора, значит мы нашли искомый элемент
			if(selpos < open_tag_begin)
			{
				break;
			}
		}

		// если закончились теги, значит мы нашли то что нужно
		*char_pos = virtual_pos;
		return true;		


		/*while(curpos = wcsstr(curpos, L"<"))
		{
			++curpos;			
			// если уже прошли selectedpos
			if(curpos > selpos)
			{
				//m_path.push_back(id);
				*char_pos = 0;
				--selpos;
				while(*selpos != L'>')
				{
					--selpos;
					*char_pos = *char_pos + 1;
				}
				return true;
			}



			// проверяем не коментарий ли это
			// TODO пока считаем, что комментариев не будет

			/*if(*curpos == L'!')
			{
				continue;
			}*/			

			/*// читаем имя тега
			// ищем пробел или закрывающую скобку
			wchar_t* p1 = wcschr(curpos, L' ');
			wchar_t* p2 = wcschr(curpos, L'>');
			if(!p2)
			{
				m_path.clear();
				return false;
			}
			wchar_t* end_tag_name;
			if((p1 > p2) || !p1)
			{
				end_tag_name = p2;
			}
			else
			{
				end_tag_name = p1;
			}

			if(!end_tag_name)
			{
				m_path.clear();
				return false;
			}

			// проверяем нужет ли закрывающий тег
			wchar_t* p3 = wcsstr(curpos, L">");
			wchar_t* p4 = wcsstr(curpos, L"/>");

			if(p4 < p3 && p4)
			{
				if(wcsstr(p3, L"<") > selpos)
				{
					m_path.push_back(id);
					*char_pos = 0;
					return true;
				}	
				++id;
				continue;
			}

			wchar_t* close_tag_name = new wchar_t[end_tag_name - curpos + 2];
			wcscpy(close_tag_name, L"/" );
			wcsncat(close_tag_name, curpos, end_tag_name - curpos);			

			// TODO могут быть вложеные теги с одним названием. типа <my_tag><my_tag></my_tag></my_tag>
			// эту ситуацию надо обязательно обработать*/



			// ищем закрывающий тег
			/*wchar_t* close_tag = wcsstr(curpos, close_tag_name);
			delete[] close_tag_name;
			close_tag_name = 0;

			if(!close_tag)
			{
				m_path.clear();
				return false;
			}

			// если закрывающий тег дальше искомого элемента, то рекурсивно заходим
			if(close_tag > selpos)
			{
				m_path.push_back(id);
				return CPFT(curpos, selpos - curpos, char_pos);
			}	
			++id;
			curpos = close_tag;
		}
		return false;*/
	}


	bool DomPath::CreatePathFromText(const wchar_t* xml, int pos, size_t* char_pos)
	{
		m_path.clear();

		return this->CPFT(xml, pos, char_pos);
	}

	int DomPath::GetNodeFromText(wchar_t* xml, int char_pos)
	{
		size_t nest_len = m_path.size(); //глубина вложенности

		if(!nest_len)
		{
			return 0;
		}

		wchar_t* curpos = 0;
		wchar_t* nested_node = xml;
		for(size_t i = 0; i < nest_len; ++i)
		{
			curpos = nested_node;
			size_t node_id = m_path[i];				
			
			for (size_t j = 0; j < node_id; ++j)
			{
				if(!tGetFirstXMLNodeParams(curpos, 0, 0, 0, &curpos))
				{
					return 0;
				}
			}
			if(!tGetFirstXMLNodeParams(curpos, 0, &nested_node, 0, &curpos))
			{
				return 0;
			}
		}

		// считаем позицию, выкидывая вложеные ноды

		wchar_t* open_tag_begin = 0;
		wchar_t* open_tag_end = 0;
		wchar_t* close_tag_begin = 0;
		wchar_t* close_tag_end = 0;
		int realcount = char_pos;
		int virtualcount = 0;
		curpos = nested_node;

		bool ok = tGetFirstXMLNodeParams(curpos, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);

		while(ok)
		{
			int tmp = virtualcount + (open_tag_begin - curpos);
			if(tmp > char_pos)
			{
				return nested_node - xml + realcount; 
			}
			virtualcount = tmp;
			realcount = realcount + (close_tag_end - open_tag_begin);			
			
			curpos = close_tag_end;
			ok = tGetFirstXMLNodeParams(curpos, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);			
		}

		return nested_node - xml + realcount; 
	}

	bool tGetFirstXMLNodeParams(const wchar_t* xml, wchar_t** open_tag_begin, wchar_t** open_tag_end, wchar_t** close_tag_begin, wchar_t** close_tag_end)
	{
		if(!xml)
		{
			return false;
		}

		// ищем ближайший тег
		wchar_t* optb = wcsstr(const_cast<wchar_t*>(xml), L"<");
		while(optb && (*(optb + 1) == L'/'))
		{
			optb = wcsstr(optb + 1, L"<");
		}
		if(!optb)
		{
			return false;
		}
		wchar_t* tag_name_start = optb + 1;

		// считываем имя тега
		// ищем ближайший пробел или закрывающую скобку
		wchar_t* tag_name_end = wcschr(tag_name_start, L' ');
		wchar_t* opte = wcschr(tag_name_start, L'>');

		// пробела может и не быть
		if((tag_name_end > opte) || (tag_name_end == 0))
		{
			tag_name_end = opte;
		}
		if(!tag_name_end)
		{
			return false;
		}
		++opte;

		wchar_t* tag_name = new wchar_t[tag_name_end - tag_name_start + 1];
		wcsncpy(tag_name, tag_name_start, tag_name_end - tag_name_start);
		tag_name[tag_name_end - tag_name_start] = 0;
		
		wchar_t* cltb = 0;
		wchar_t* clte = 0;
		
		if(*(opte - 2) == L'/')
		{
			cltb = opte;
			clte = opte;
		}
		else
		{
			// ищем закрывающий тег
			if(!tFindCloseTag(opte, tag_name, &cltb))
			{
				delete[] tag_name;
				return false;
			}
			clte = wcschr(cltb, L'>');
			if(!clte)
			{
				delete[] tag_name;
				return false;
			}
			++clte;
		}
		delete[] tag_name;

		// возвращаем затребованные значения
		if(open_tag_begin)
		{
			*open_tag_begin = optb;
		}
		if(open_tag_end)
		{
			*open_tag_end = opte;
		}
		if(close_tag_begin)
		{
			*close_tag_begin = cltb;
		}
		if(close_tag_end)
		{
			*close_tag_end = clte;			
		}

		return true;
	}

	bool tFindCloseTag(wchar_t* start, wchar_t* tag_name, wchar_t** close_tag)
	{
		// added by SeNS: check space in tags
		CString tag(tag_name); tag.Trim(); if (tag.IsEmpty()) return false;

		// будем пропускать пары <tag_name...></tag_name>
		//wchar_t* close_tagname
		wchar_t* next_tag = wcsstr(start, tag_name);
		while(next_tag)
		{
			// если это закрывающий тег, то значит мы нашли что надо
			if((*(next_tag - 1) == L'/')&&(*(next_tag - 2) == L'<'))
			{
				*close_tag = next_tag - 2;
				return true;
			}

			// проверяем таг ли это вообще
			if(*(next_tag - 1) != L'<')
			{
				next_tag = wcsstr(next_tag + wcslen(tag_name), tag_name);
				continue;
			}
			
			// это был открывающий тег
			// ищем его закрывающий тег
			if(!tFindCloseTag(next_tag + wcslen(tag_name), tag_name, &next_tag))
			{
				return false;
			}
			next_tag = wcsstr(next_tag + wcslen(tag_name), tag_name);
		}
		return false;
	}

	/*MSHTML::IHTMLDOMNodePtr GetNodeFromPos(MSHTML::IHTMLElementPtr elem, int pos)
	{
		MSHTML::IHTMLDOMNodePtr node(elem);
		if(!(bool)node)
		{
			return 0;
		}
		node = node->firstChild;

		int index = 0; // поярдковый номер ноды в HTML элементе

		// по тексту определяем поряковый номер ноды
		wchar_t* text = elem->innerHTML;
		wchar_t* openb = 0;
		wchar_t* opene = 0;
		wchar_t* closeb = 0;
		wchar_t* closee = 0;

		wchar_t* curpos = text + pos;

		while(tGetFirstXMLNodeParams(text, &openb, &opene, &closeb, &closee))
		{
			++index;
			if(curpos < openb)
			{
				break;
			}
			if((curpos >= openb && curpos <= opene) || (curpos >= closeb && curpos <= closee))
			{
				pos = 0; 
				break;
			}
			if(curpos > opene && curpos < closeb)
			{
				pos = pos - (opene - text);
				break;
			}
            
			if(text != openb)
			{
				++index;
			}
			pos = pos - (closee - text);
			text = closee;
		}

		for(int i = 0; i < index; ++i)
		{
			if(!(bool)node)
			{
				return 0;
			}
			node = node->nextSibling;			
		}

		return node;
	}*/

	int CountTextNodeChars(wchar_t* node_begin, wchar_t* node_end)
	{
		wchar_t* openb = 0;
		wchar_t* opene = 0;
		wchar_t* closeb = 0;
		wchar_t* closee = 0;
		size_t count = 0;

		bool ok = tGetFirstXMLNodeParams(node_begin, &openb, &opene, &closeb, &closee);
		if(!ok)
		{
			return node_end - node_begin;
		}

		while(ok)
		{
			if(openb >= node_end)
			{
				count = count + node_end - node_begin;
			}
			node_begin = closee;
			count = count + (openb - node_begin);
			count += CountTextNodeChars(opene, closeb);
			ok = tGetFirstXMLNodeParams(node_begin, &openb, &opene, &closeb, &closee);
		}
		return count;
	}

	bool IsParentElement(MSHTML::IHTMLDOMNodePtr elem, MSHTML::IHTMLDOMNodePtr parent)
	{
		if(!(bool)parent || !(bool)elem)
		{
			return false;
		}
		if(parent == elem)
		{
			return true;
		}
		while(elem = elem->parentNode)
		{
			if(elem == parent)
			{
				return true;
			}
		}
		return false;
	}	

	bool IsParentElement(MSXML2::IXMLDOMNodePtr elem, MSXML2::IXMLDOMNodePtr parent)
	{
		if(!(bool)parent || !(bool)elem)
		{
			return false;
		}
		if(parent == elem)
		{
			return true;
		}
		while(elem = elem->parentNode)
		{
			if(elem == parent)
			{
				return true;
			}
		}
		return false;
	}	
#ifndef NO_EXTRN_SETTINGS
	void SaveFileSelectedPos(const CString& filename, int pos)
	{
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _Settings.GetKeyPath() + _T("\\Documents"));
		rk.SetDWORDValue(filename, pos);
		rk.Close();
	}


	int GetFileSelectedPos(const CString& filename)
	{
		CRegKey rk;
		rk.Open(HKEY_CURRENT_USER, _Settings.GetKeyPath() + _T("\\Documents"));
		DWORD pos = 0;
		rk.QueryDWORDValue(filename, pos);		
		rk.Close();
		return pos;
	}
#endif
	MSHTML::IHTMLElementPtr	FindTitleNode(MSHTML::IHTMLDOMNodePtr elem) 
	{
		MSHTML::IHTMLDOMNodePtr node(elem->firstChild);

		if ((bool)node && node->nodeType==1 && U::scmp(node->nodeName,L"DIV")==0) {
			_bstr_t   cls(MSHTML::IHTMLElementPtr(node)->className);
			if (U::scmp(cls,L"image")==0) {
				node=node->nextSibling;
				if (node->nodeType!=1 || U::scmp(node->nodeName,L"DIV"))
					return NULL;
				cls=MSHTML::IHTMLElementPtr(node)->className;
			}
			if (U::scmp(cls,L"title")==0)
				return MSHTML::IHTMLElementPtr(node);
		}
		return NULL;
	}

	CString	FindTitle(MSHTML::IHTMLDOMNodePtr elem) 
	{
		MSHTML::IHTMLElementPtr tn(FindTitleNode(elem));
		if (tn)
			return (const wchar_t *)tn->innerText;
		return CString();
	}

	CString GetImageFileName(MSHTML::IHTMLDOMNodePtr elem)
	{
		_bstr_t   cls(MSHTML::IHTMLElementPtr(elem)->className);
		if (U::scmp(cls,L"image") != 0)
			return L"";

		CString href(MSHTML::IHTMLElementPtr(elem)->getAttribute(L"href", 0));
		if(!href.GetLength())
			return L"";

		CString	name = href.Right(href.GetLength() - 1);
		return name;	
	}


	bool HasSubFolders(const CString path)
	{
		WIN32_FIND_DATA fd;
		HANDLE hFind = FindFirstFile(path + L"*.*", &fd);
		if(hFind != INVALID_HANDLE_VALUE)
		{
			do
			{
				if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					if(wcscmp(fd.cFileName, L".") && wcscmp(fd.cFileName, L".."))
					{
						FindClose(hFind);
						return true;
					}
				}
				else
					if(!(fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT))
					{
						FindClose(hFind);
						return true;
					}
			} while (FindNextFile(hFind, &fd));

			FindClose(hFind);
			return false;
		}
		else return false;
	}

	bool HasFilesWithExt(const CString path, const TCHAR* mask)
	{
		WIN32_FIND_DATA fd;
		HANDLE hFind = FindFirstFile(path + mask, &fd);
		if(hFind != INVALID_HANDLE_VALUE && !(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
		{
			FindClose(hFind);
			return true;
		}
		else return false;
	}

	bool HasScriptsEndpoint(const CString path, TCHAR* ext)
	{
		bool found = false;
		WIN32_FIND_DATA fd;
		HANDLE hFind = FindFirstFile(path + L"*.*", &fd);
		if(hFind == INVALID_HANDLE_VALUE)
		{
			return false;
		}
		else
		{
			if(HasFilesWithExt(path, ext))
			{
				FindClose(hFind);
				return true;
			}
			else
			{
				FindNextFile(hFind, &fd);
				do				
				{
					if(wcscmp(fd.cFileName, L".") && wcscmp(fd.cFileName, L"..") && (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
					{
						found = HasScriptsEndpoint((path + fd.cFileName) + L"\\", ext);						
					}					
				} while (FindNextFile(hFind, &fd));

				FindClose(hFind);
			}
		}
		return found;
	}

	//true - old
	bool CheckScriptsVersion(const wchar_t* Tested)
	{
		wchar_t Name[MAX_PATH];
		wcscpy(Name, Tested);
		Name[wcslen(Name) - 3] = 0;

		wchar_t* pos = wcschr(Name, L'_');
		
		if(*(pos+1) == 0)
			return false;
		
		*pos = 0;
		for(unsigned int i = 0; i < wcslen(Name); ++i)
		{
			if(Name[i] < '0' || Name[i] > '9')
				return false;
		}
		
		return true;
	}

	CString Transliterate (CString src)
	{
		CString dst = L"";
		CString from = L"абвгдеёжзийклмнопрстуфцчшщъыьэюяєіїАБВГДЕЁЖЗИЙКЛМНОПРСТУФЦЧШЩЪЫЬЭЮЯЄІЇ";
		CString to[] = { L"a", L"b", L"v", L"g", L"d", L"e", L"jo", L"zh", L"z", L"i", L"jj", L"k", 
						 L"l", L"m", L"n", L"o", L"p", L"r", L"s", L"t", L"u", L"f", L"c", L"ch", 
						 L"sh", L"shh", L"\"", L"y", L"'", L"eh", L"ju", L"ja", L"je", L"i", L"i'", 
						 L"A", L"B", L"V", L"G", L"D", L"E", L"JO", L"ZH", L"Z", L"I", L"JJ", L"K", 
						 L"L", L"M", L"N", L"O", L"P", L"R", L"S", L"T", L"U", L"F", L"C", L"CH", 
						 L"SH", L"SHH", L"\"", L"Y", L"'", L"EH", L"JU", L"JA", L"JE", L"I", L"I'"};

		for (int i=0; i<src.GetLength(); i++)
		{
			int j = from.Find (src[i]);
			if (j < 0) dst += src[i]; else dst += to[j];
		}
		return dst;
	}

	CString URLDecode(const CString& inStr)
	{
		// decode escape sequences
		CString res;
		for (int x = 0; x < inStr.GetLength(); x++)
		{
			if (inStr.GetAt(x) == _T('%') && x + 2 < inStr.GetLength() && iswxdigit (inStr.GetAt(x+1)) && iswxdigit(inStr.GetAt(x+2)))
			{
				TCHAR hexstr[3];
				_tcsncpy(hexstr, inStr.Mid(x+1, 2), 2);
				hexstr[2] = _T('\0');
				x += 2;

				// Convert the hex to ASCII
				res.AppendChar((TCHAR)_tcstoul(hexstr, NULL, 16));
			}
			else
			{
				res.AppendChar(inStr.GetAt(x));
			}
		}
		return res;
	}

	LPCWSTR GetProductInfo()
	{
		if (g_strProductInfo.IsEmpty())
		{
			CString strRet;
			CString strProductName;
			CString strProductVersion;
			CString strFileName;

			::GetModuleFileNameW(nullptr, strFileName.GetBuffer(MAX_PATH), MAX_PATH);
			strFileName.ReleaseBuffer();

			DWORD dwHandle;
			DWORD dwLen = ::GetFileVersionInfoSizeW(strFileName, &dwHandle);
			if (dwLen > 0)
			{
				CHeapPtr<BYTE> spBlock;
				if (spBlock.Allocate(dwLen))
				{
					if (::GetFileVersionInfoW(strFileName, dwHandle, dwLen, spBlock) == TRUE)
					{
						PLCID *plcid;
						UINT uLen;
						if (VerQueryValueW(spBlock.m_pData, L"\\VarFileInfo\\Translation", (LPVOID*)&plcid, &uLen) == TRUE)
						{
							CString strFormat;
							LPWSTR *pszBuffer;

							strFormat.Format(L"\\StringFileInfo\\%04x%04x\\ProductName", LOWORD(plcid[0]), HIWORD(plcid[0]));
							if (::VerQueryValueW(spBlock, strFormat, (LPVOID*)&pszBuffer, &uLen) == TRUE)
							{
								strProductName = (LPCWSTR)pszBuffer;
							}

							strFormat.Format(L"\\StringFileInfo\\%04x%04x\\ProductVersion", LOWORD(plcid[0]), HIWORD(plcid[0]));
							if (::VerQueryValueW(spBlock, strFormat, (LPVOID*)&pszBuffer, &uLen) == TRUE)
							{
								strProductVersion = (LPCWSTR)pszBuffer;
							}
						}
					}
				}
			}


			g_strProductInfo = strProductName + L" Release " + strProductVersion;
		}

		return g_strProductInfo;
	}

}
