#ifndef UTILS_H
#define UTILS_H

#include <deque>

#ifdef USE_PCRE
#include <atlbase.h>
#include <atlwin.h>
#include <atltypes.h>
#include <atlstr.h>
#include <atlcoll.h>
#include "pcre.h"
#endif

#define UTF8_CHAR_LEN(byte) ((0xE5000000 >> ((byte >> 3) & 0x1e)) & 3) + 1

// command line arguments
extern CSimpleArray<CString> _ARGV;

namespace U // place all utilities into their own namespace
{
// APP messages
enum
{
	WM_POSTCREATE = WM_APP,
	WM_SETSTATUSTEXT,
	WM_TRACKPOPUPMENU
};

struct TRACKPARAMS
{
	HMENU hMenu;
	UINT uFlags;
	int x;
	int y;
};

// command line arguments
struct CmdLineArgs
{
	// no args yet
	bool start_in_desc_mode;

	CmdLineArgs() : start_in_desc_mode(false) {}
};
extern CmdLineArgs _ARGS;
bool ParseCmdLineArgs();

struct ElTextHTML
{
	CString html;
	CString text;

	ElTextHTML(BSTR html, BSTR text)
	{
		this->html = html;
		this->text = text;
	}
};

void InitKeycodes();
void ChangeAttribute(MSHTML::IHTMLElementPtr elem, const wchar_t *attrib, const wchar_t *value);

// loading data into array
HRESULT LoadImageFile(const TCHAR *filename, VARIANT *vt);

// a generic stream
class HandleStreamImpl : public CComObjectRoot, public IStream
{
  private:
	HANDLE m_handle;
	bool m_close;

  public:
	typedef IStream Interface;

	HandleStreamImpl() : m_handle(INVALID_HANDLE_VALUE), m_close(true) {}
	virtual ~HandleStreamImpl()
	{
		if (m_close)
			CloseHandle(m_handle);
	}

	void SetHandle(HANDLE hf, bool cl = true)
	{
		m_handle = hf;
		m_close = cl;
	}

	BEGIN_COM_MAP(HandleStreamImpl)
	COM_INTERFACE_ENTRY(ISequentialStream)
	COM_INTERFACE_ENTRY(IStream)
	END_COM_MAP()

	// ISequentialStream
	STDMETHOD(Read)
	(void *pv, ULONG cb, ULONG *pcbRead)
	{
		DWORD nrd;
		if (ReadFile(m_handle, pv, cb, &nrd, NULL))
		{
			*pcbRead = nrd;
			return S_OK;
		}
		if (GetLastError() == ERROR_BROKEN_PIPE)
		{ // treat as eof
			*pcbRead = 0;
			return S_OK;
		}
		return HRESULT_FROM_WIN32(GetLastError());
	}
	STDMETHOD(Write)
	(const void *pv, ULONG cb, ULONG *pcbWr)
	{
		DWORD nwr;
		if (WriteFile(m_handle, pv, cb, &nwr, NULL))
		{
			*pcbWr = nwr;
			return S_OK;
		}
		return HRESULT_FROM_WIN32(GetLastError());
	}

	// IStream
	STDMETHOD(Seek)
	(LARGE_INTEGER, DWORD, ULARGE_INTEGER *) { return E_NOTIMPL; }
	STDMETHOD(SetSize)
	(ULARGE_INTEGER) { return E_NOTIMPL; }
	STDMETHOD(CopyTo)
	(IStream *, ULARGE_INTEGER, ULARGE_INTEGER *, ULARGE_INTEGER *) { return E_NOTIMPL; }
	STDMETHOD(Commit)
	(DWORD) { return E_NOTIMPL; }
	STDMETHOD(Revert)
	() { return E_NOTIMPL; }
	STDMETHOD(LockRegion)
	(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) { return E_NOTIMPL; }
	STDMETHOD(UnlockRegion)
	(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) { return E_NOTIMPL; }
	STDMETHOD(Stat)
	(STATSTG *, DWORD) { return E_NOTIMPL; }
	STDMETHOD(Clone)
	(IStream **) { return E_NOTIMPL; }
};

typedef CComObject<HandleStreamImpl> HandleStream;
typedef CComPtr<HandleStream> HandleStreamPtr;

// strings
int scmp(const wchar_t *s1, const wchar_t *s2);
CString GetMimeType(const CString &filename);
bool GetImageDimsByPath(const wchar_t *pszFileName, int *nWidth, int *nHeight);
bool GetImageDimsByData(SAFEARRAY *data, ULONG length, int *nWidth, int *nHeight);
bool is_whitespace(const wchar_t *spc);
void NormalizeInplace(CString &s);
void RemoveSpaces(wchar_t *zstr);
extern inline CString Normalize(const CString &s)
{
	CString p(s);
	NormalizeInplace(p);
	return p;
}
CString GetFileTitle(const TCHAR *filename);
CString UrlFromPath(const CString &path);
CString PrepareDefaultId(const CString &path);

// settings in the registry
CString QuerySV(HKEY hKey, const TCHAR *name, const TCHAR *def = NULL);
DWORD QueryIV(HKEY hKey, const TCHAR *name, DWORD defval = 0);

template <class T>
T QueryBV(HKEY hKey, const TCHAR *name, T def)
{
	BYTE *buff = new BYTE[sizeof(T)];
	ZeroMemory(buff, sizeof(T));
	DWORD type = REG_BINARY;
	DWORD len = 0;

	if (::RegQueryValueEx(hKey, name, NULL, &type, NULL, &len) != ERROR_SUCCESS || (type != REG_BINARY))
		return def;
	if (::RegQueryValueEx(hKey, name, NULL, &type, buff, &len) != ERROR_SUCCESS)
		return def;

	T ret;
	ZeroMemory(&ret, sizeof(T));
	CopyMemory(&ret, buff, sizeof(T));
	delete[] buff;

	return ret;
}
void InitSettings();
void InitSettingsHotkeyGroups();

// windows api
HFONT CreatePtFont(int sizept, const TCHAR *facename, bool fBold = false, bool fItalic = false);
//CString GetFullPathName(const CString &filename);
CString Win32ErrMsg(DWORD code);
CString GetWindowText(HWND hWnd);
void ReportError(HRESULT hr);
void ReportError(_com_error &e);
void ReportError(const CString &message);
CString GetProgDir();
CString GetSettingsDir();
CString GetDocTReeScriptsDir();
CString GetProgDirFile(const CString &filename);
CString GetCBString(HWND hCB, int idx);
bool HasSubFolders(const CString);
bool HasFilesWithExt(const CString, const TCHAR *);
bool HasScriptsEndpoint(const CString, TCHAR *);
bool CheckScriptsVersion(const wchar_t *);
WORD StringToKeycode(CString);
WORD VKToFVirt(WORD);
CString KeycodeToString(WORD);
CString AccelToString(ACCEL);

// a generic inputbox
int InputBox(CString &result, const wchar_t *title, const wchar_t *prompt);

// html
CString GetAttrCS(MSHTML::IHTMLElement *elem, const wchar_t *attr);
_bstr_t GetAttrB(MSHTML::IHTMLElement *elem, const wchar_t *attr);

// unicode char names (win2k/xp only)
CString GetCharName(int ch);
// convert multibyte string to UTF-8
char *ToUtf8(const CString &s, int &patlen);

// msxml support
MSXML2::IXMLDOMDocument2Ptr CreateDocument(bool fFreeThreaded = false);
void ReportParseError(MSXML2::IXMLDOMDocument2Ptr doc);
//bool LoadXml(MSXML2::IXMLDOMDocument2Ptr doc, const CString &url);
MSXML2::IXSLTemplatePtr CreateTemplate();

void SaveFileSelectedPos(const CString &filename, int pos);
int GetFileSelectedPos(const CString &filename);

// Class to construct HTML-XML DOM path;
class DomPath
{
  private:
	std::deque<int> m_path;

  public:
	DomPath(){};
	~DomPath(){};
	bool CreatePathFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root, MSHTML::IHTMLDOMNodePtr EndNode);
	bool CreatePathFromXMLDOM(MSXML2::IXMLDOMNodePtr root, MSXML2::IXMLDOMNodePtr EndNode);
	bool CreatePathFromText(const wchar_t *xml, int pos, int *char_pos);

	MSXML2::IXMLDOMNodePtr GetNodeFromXMLDOM(MSXML2::IXMLDOMNodePtr root);
	MSHTML::IHTMLDOMNodePtr GetNodeFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root);
	MSHTML::IHTMLDOMNodePtr FindSelectedNodeInXMLDOM(MSXML2::IXMLDOMNodePtr root);
	int GetNodeFromText(wchar_t *xml, int char_pos);

	operator bstr_t()
	{
		bstr_t ret;
		for (unsigned int i = 0; i < m_path.size(); ++i)
		{
			wchar_t numb[10];
			wsprintf(numb, L"%d", m_path[i]);
			ret = ret + numb + L"/";
		}
		return ret;
	}

  private:
	bool CPFT(const wchar_t *xml, int pos, int *char_pos);
};

bool tGetFirstXMLNodeParams(const wchar_t *xml, wchar_t **open_tag_begin, wchar_t **open_tag_end, wchar_t **close_tag_begin, wchar_t **close_tag_end);
bool tFindCloseTag(wchar_t *start, wchar_t *tag_name, wchar_t **close_tag);
MSHTML::IHTMLDOMNodePtr GetNodeFromPos(MSHTML::IHTMLElementPtr elem, int pos);
int CountTextNodeChars(wchar_t *NodeBegin, wchar_t *NodeEnd);

bool IsParentElement(MSHTML::IHTMLDOMNodePtr elem, MSHTML::IHTMLDOMNodePtr parent);
bool IsParentElement(MSXML2::IXMLDOMNodePtr elem, MSXML2::IXMLDOMNodePtr parent);

MSHTML::IHTMLElementPtr FindTitleNode(MSHTML::IHTMLDOMNodePtr elem);
CString FindTitle(MSHTML::IHTMLDOMNodePtr elem);
CString GetImageFileName(MSHTML::IHTMLDOMNodePtr elem);

CString Transliterate(CString src);
CString URLDecode(const CString &inStr);

// persistent wait cursor
// allowing the wait cursor to persist across messages.
class CPersistentWaitCursor : public CMessageFilter
{
  public:
	CPersistentWaitCursor() : m_wait(0), m_old(0)
	{
		CMessageLoop *pLoop = _Module.GetMessageLoop();
		if (pLoop)
			pLoop->AddMessageFilter(this);
		m_wait = ::LoadCursor(NULL, IDC_WAIT);
		if (m_wait)
			m_old = ::SetCursor(m_wait);
	}

	~CPersistentWaitCursor()
	{
		CMessageLoop *pLoop = _Module.GetMessageLoop();
		if (pLoop)
			pLoop->RemoveMessageFilter(this);
		if (m_old)
			::SetCursor(m_old);
	}

	BOOL PreTranslateMessage(MSG *msg)
	{
		if (msg->message == WM_SETCURSOR && m_wait)
		{
			::SetCursor(m_wait);
			return TRUE;
		}
		return FALSE;
	}

  protected:
	HCURSOR m_wait, m_old;
};
// Regular expression classes
#ifdef USE_PCRE

// VB-compatible wrappers
typedef CSimpleArray<CString> CStrings;

struct ISubMatches
{
  public:
	__declspec(property(get = GetItem)) CString Item[];
	__declspec(property(get = GetCount)) long Count;
	CString GetItem(long index) { return m_strs[index]; }
	long GetCount() { return m_strs.GetSize(); }
	void AddItem(CString item) { m_strs.Add(item); }

  private:
	CStrings m_strs;
};

struct IMatch2
{
  public:
	IMatch2(CString str, int index) : m_str(str), m_index(index) {}
	__declspec(property(get = GetValue)) CString Value;
	__declspec(property(get = GetFirstIndex)) long FirstIndex;
	__declspec(property(get = GetLength)) long Length;
	__declspec(property(get = GetSubMatches)) ISubMatches *SubMatches;
	CString GetValue() { return m_str; }
	long GetFirstIndex() { return m_index; }
	long GetLength() { return m_str.GetLength(); }
	ISubMatches *GetSubMatches() { return &m_submatches; }
	void AddSubMatch(CString item) { m_submatches.AddItem(item); }

  private:
	CString m_str;
	int m_index;
	ISubMatches m_submatches;
};

struct IMatchCollection
{
  public:
	IMatchCollection() { m_matches.RemoveAll(); }
	__declspec(property(get = GetCount)) long Count;
	__declspec(property(get = GetItem)) IMatch2 *Item[];
	long GetCount() { return m_matches.GetSize(); }
	IMatch2 *GetItem(long index)
	{
		int count = GetCount();
		if ((count == 0) || (index >= count))
			return NULL;
		return &m_matches[index];
	}
	void AddItem(IMatch2 *item) { m_matches.Add(*item); }

  private:
	CSimpleArray<IMatch2> m_matches;
};

// wrapper classes compatible with MS VBScrip
struct IRegExp2
{
  public:
	__declspec(property(get = GetPattern, put = PutPattern)) CString Pattern;
	__declspec(property(get = GetIgnoreCase, put = PutIgnoreCase)) VARIANT_BOOL IgnoreCase;
	__declspec(property(get = GetGlobal, put = PutGlobal)) VARIANT_BOOL Global;
	__declspec(property(get = GetMultiline, put = PutMultiline)) VARIANT_BOOL Multiline;
	IMatchCollection *Execute(CString sourceString);
	IMatchCollection *Execute(_bstr_t sourceString)
	{
		CString src;
		src.SetString((wchar_t *)sourceString);
		return Execute(src);
	};
	CString GetPattern() { return m_pattern; }
	void PutPattern(CString val) { m_pattern.SetString(val); }
	VARIANT_BOOL GetIgnoreCase() { return m_ignorecase; }
	void PutIgnoreCase(VARIANT_BOOL val) { m_ignorecase = val; }
	VARIANT_BOOL GetGlobal() { return m_global; }
	void PutGlobal(VARIANT_BOOL val) { m_global = val; }
	VARIANT_BOOL GetMultiline() { return m_multiline; }
	void PutMultiline(VARIANT_BOOL val) { m_multiline = val; }

  private:
	CString m_pattern;
	VARIANT_BOOL m_ignorecase, m_global, m_multiline;
};

typedef IRegExp2 *RegExp;
typedef IMatchCollection *ReMatches;
typedef IMatch2 *ReMatch;
typedef ISubMatches *ReSubMatches;
#else
// visual basic RegExps
typedef VBScript_RegExp_55::IRegExp2Ptr RegExp;
typedef VBScript_RegExp_55::IMatchCollection2Ptr ReMatches;
typedef VBScript_RegExp_55::IMatch2Ptr ReMatch;
typedef VBScript_RegExp_55::ISubMatchesPtr ReSubMatches;
#endif
CString GetFileExtension(const CString &inStr);
CString GetStringEncoding(const char *buffer);
CString GetInputElementValue(MSHTML::IHTMLElementPtr m_cur_input_element);

} // namespace U

#endif
