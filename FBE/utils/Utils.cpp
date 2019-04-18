#include "stdafx.h"

#include <deque>
#include "utils.h"
#include "../ModelessDialog.h"
#include "../ErrsEx.h"
#include "../Settings.h"
#include "../resource.h"
#include "../res1.h"

#include <iostream>
#include <sstream>

#ifndef NO_EXTRN_SETTINGS
extern CSettings _Settings;
#endif

extern RECT dialogRect;

bool VBErr = false;

namespace U
{
	CSimpleMap<CString, WORD> keycodes;
	///<summary>Fill keycodes list</summary>
	void InitKeycodes()
	{
		keycodes.RemoveAll();

		keycodes.Add(L"0", 0x30);
		keycodes.Add(L"1", 0x31);
		keycodes.Add(L"2", 0x32);
		keycodes.Add(L"3", 0x33);
		keycodes.Add(L"4", 0x34);
		keycodes.Add(L"5", 0x35);
		keycodes.Add(L"6", 0x36);
		keycodes.Add(L"7", 0x37);
		keycodes.Add(L"8", 0x38);
		keycodes.Add(L"9", 0x39);
		keycodes.Add(L"A", 0x41);
		keycodes.Add(L"B", 0x42);
		keycodes.Add(L"C", 0x43);
		keycodes.Add(L"D", 0x44);
		keycodes.Add(L"E", 0x45);
		keycodes.Add(L"F", 0x46);
		keycodes.Add(L"G", 0x47);
		keycodes.Add(L"H", 0x48);
		keycodes.Add(L"I", 0x49);
		keycodes.Add(L"J", 0x4A);
		keycodes.Add(L"K", 0x4B);
		keycodes.Add(L"L", 0x4C);
		keycodes.Add(L"M", 0x4D);
		keycodes.Add(L"N", 0x4E);
		keycodes.Add(L"O", 0x4F);
		keycodes.Add(L"P", 0x50);
		keycodes.Add(L"Q", 0x51);
		keycodes.Add(L"R", 0x52);
		keycodes.Add(L"S", 0x53);
		keycodes.Add(L"T", 0x54);
		keycodes.Add(L"U", 0x55);
		keycodes.Add(L"V", 0x56);
		keycodes.Add(L"W", 0x57);
		keycodes.Add(L"X", 0x58);
		keycodes.Add(L"Y", 0x59);
		keycodes.Add(L"Z", 0x5A);

		keycodes.Add(L"F1", VK_F1);
		keycodes.Add(L"F2", VK_F2);
		keycodes.Add(L"F3", VK_F3);
		keycodes.Add(L"F4", VK_F4);
		keycodes.Add(L"F5", VK_F5);
		keycodes.Add(L"F6", VK_F6);
		keycodes.Add(L"F7", VK_F7);
		keycodes.Add(L"F8", VK_F8);
		keycodes.Add(L"F9", VK_F9);
		keycodes.Add(L"F10", VK_F10);
		keycodes.Add(L"F11", VK_F11);
		keycodes.Add(L"F12", VK_F12);

		keycodes.Add(L"Tab", VK_TAB);
		keycodes.Add(L"Backspace", VK_BACK);
		keycodes.Add(L"Spacebar", 0x20);
		keycodes.Add(L"Alt", 0x12);
		keycodes.Add(L"Esc", VK_ESCAPE);
		keycodes.Add(L"Ctrl", VK_CONTROL);
		keycodes.Add(L"Shift", VK_SHIFT);
		keycodes.Add(L"Enter", 0x0D);
		keycodes.Add(L"Delete", VK_DELETE);
		// added by SeNS
		keycodes.Add(L"Ins", VK_INSERT);
		keycodes.Add(L"Home", VK_HOME);
		keycodes.Add(L"End", VK_END);
		keycodes.Add(L"PageUp", VK_PRIOR);
		keycodes.Add(L"PageDn", VK_NEXT);

		keycodes.Add(L"Num 0", VK_NUMPAD0);
		keycodes.Add(L"Num 1", VK_NUMPAD1);
		keycodes.Add(L"Num 2", VK_NUMPAD2);
		keycodes.Add(L"Num 3", VK_NUMPAD3);
		keycodes.Add(L"Num 4", VK_NUMPAD4);
		keycodes.Add(L"Num 5", VK_NUMPAD5);
		keycodes.Add(L"Num 6", VK_NUMPAD6);
		keycodes.Add(L"Num 7", VK_NUMPAD7);
		keycodes.Add(L"Num 8", VK_NUMPAD8);
		keycodes.Add(L"Num 9", VK_NUMPAD9);

		keycodes.Add(L"*", VK_MULTIPLY);
		keycodes.Add(L"+", VK_ADD);
		keycodes.Add(L"-", VK_SUBTRACT);
		keycodes.Add(L".", VK_DECIMAL);
		keycodes.Add(L"/", VK_DIVIDE);
		keycodes.Add(L"SEP", VK_SEPARATOR);

		keycodes.Add(L":", VK_OEM_1);
		keycodes.Add(L"?", VK_OEM_2);

		keycodes.Add(L"[", VK_OEM_4);
		keycodes.Add(L"]", VK_OEM_6);
		keycodes.Add(L"\"", VK_OEM_7);
		keycodes.Add(L"<", VK_OEM_COMMA);
		keycodes.Add(L">", VK_OEM_PERIOD);
		keycodes.Add(L"-", VK_OEM_MINUS);
		keycodes.Add(L"~", VK_OEM_3);
		keycodes.Add(L"+", VK_OEM_PLUS);

		keycodes.Add(L"←", VK_LEFT);
		keycodes.Add(L"↑", VK_UP);
		keycodes.Add(L"→", VK_RIGHT);
		keycodes.Add(L"↓", VK_DOWN);
	}

	///<summary>Change attribute value</summary>
	///<param name="elem">HTML element</param>
	///<param name="attrib">Attribute name</param>
	///<param name="value">Value</param>
	void ChangeAttribute(MSHTML::IHTMLElementPtr elem, const wchar_t* attrib, const wchar_t* value)
	{
		if (U::scmp(attrib, L"class"))
			elem->setAttribute(attrib, _variant_t(value), 1);
		else
			elem->className = value;
		// send message to main window
#ifndef NO_EXTRN_SETTINGS
		::SendMessage(_Settings.GetMainWindow(), WM_COMMAND, MAKEWPARAM(0, IDN_TREE_RESTORE), 0);
#endif
	}

	HandleStreamPtr	NewStream(HANDLE& hf, bool fClose)
	{
		HandleStream  *hs;
		CheckError(HandleStream::CreateInstance(&hs));
		hs->SetHandle(hf, fClose);
		if (fClose)
			hf = INVALID_HANDLE_VALUE;
		return hs;
	}


	///<summary>Replace non-print characters in string on space</summary>
	///<param name="s">String (in\out)</param>
	void  NormalizeInplace(CString& s)
	{
		s.TrimLeft();
		int	  len = s.GetLength();
		TCHAR  *p = s.GetBuffer(len);
		TCHAR  *r = p;
		TCHAR  *q = p;
		TCHAR  *e = p + len;

		while (p < e)
		{
			if (iswgraph(*p))
				*q++ = *p;
			else if (*q != L' ')
				*q++ = L' ';
			++p;
		}
		s.ReleaseBuffer(q - r);
/*		Old variant
int	  len = s.GetLength();
		TCHAR	  *p = s.GetBuffer(len);
		TCHAR	  *r = p;
		TCHAR	  *q = p;
		TCHAR	  *e = p + len;
		int	  state = 0;

		while (p < e)
		{
			switch (state)
			{
			case 0: //skip leading space
				if (iswgraph(*p)) //32=SPACE 
				{
					*q++ = *p;
					state = 1;
				}
				break;
			case 1:
				if (iswgraph(*p))
					*q++ = *p;
				else
					state = 2;
				break;
			case 2: //replace non-print characters on space
				if (iswgraph(*p))
				{
					*q++ = L' ';
					*q++ = *p;
					state = 1;
				}
				break;
			}
			++p;
		}
		s.ReleaseBuffer(q - r);*/
	}

	///<summary>Remove non-print characters from string on space</summary>
	///<param name="zstr">Null-terminated string (in\out)</param>
	void  RemoveSpaces(wchar_t *zstr)
	{
		wchar_t  *q = zstr;

		while (*zstr) {
			if (iswgraph(*zstr))
				*q++ = *zstr;
			++zstr;
		}
		*q = L'\0';
	}

	///<summary>Compare two null-terminated strings</summary>
	///<param name="s1">First sting</param>
	///<param name="s2">Second sting</param>
	///<returns>0 - if strings equal or both null or both empty</returns>
	int	scmp(const wchar_t *s1, const wchar_t *s2)
	{
		bool f1 = !s1 || !*s1;
		bool f2 = !s2 || !*s2;
		if (f1 && f2)
			return 0;
		if (f1)
			return -1;
		if (f2)
			return 1;
		return wcscmp(s1, s2);
	}

	///<summary>Get mime type of file content</summary>
	///<param name="filename">Filename</param>
	///<returns>MIME-type string</returns>
	CString GetMimeType(const CString& filename)
	{
		CString ext = GetFileExtension(filename);
		if (ext == L"jpg" || ext == L"jpeg")
		{
			return L"image/jpeg";
		}
		if (ext == L"png")
		{
			return L"image/png";
		}
		return L"application/octet-stream";
/*		  Old variant
		  CString fn(filename);
		int cp = fn.ReverseFind(_T('.'));
		if (cp < 0)
			os:
		return _T("application/octet-stream");
		fn.Delete(0, cp);

		CRegKey rk;
		if (rk.Open(HKEY_CLASSES_ROOT, fn, KEY_READ) != ERROR_SUCCESS)
			goto os;

		CString ret;
		ULONG len = 128;
		TCHAR* rbuf = ret.GetBuffer(len);
		rbuf[0] = _T('\0');
		LONG rv = rk.QueryStringValue(_T("Content Type"), rbuf, &len);
		ret.ReleaseBuffer();

		if (rv != ERROR_SUCCESS)
			goto os;

		return ret;*/
	}

	///<summary>Get filename without extension</summary>
	///<param name="filename">Full path</param>
	///<returns>Filename without extension</returns>
	CString	GetFileTitle(const TCHAR *filename) 
    {
		CString   ret;
		TCHAR	    *buf = ret.GetBuffer(MAX_PATH);
		if (::GetFileTitle(filename, buf, MAX_PATH))
			ret.ReleaseBuffer(0);
		else
			ret.ReleaseBuffer();
		return ret;
	}
	
	DWORD QueryIV(HKEY hKey, const TCHAR *name, DWORD defval) {
		DWORD	dw;
		DWORD	len = sizeof(DWORD);
		DWORD	type = REG_DWORD;
		if (::RegQueryValueEx(hKey, name, NULL, &type, (BYTE*)&dw, &len) != ERROR_SUCCESS ||
			type != REG_DWORD || len != sizeof(DWORD))
			return defval;
		return dw;
	}
	
	CString	QuerySV(HKEY hKey, const TCHAR *name, const TCHAR *def) {
		CString   ret;
		DWORD	    type = REG_SZ;
		DWORD	    len = 0;
		if (::RegQueryValueEx(hKey, name, NULL, &type, NULL, &len) != ERROR_SUCCESS ||
			(type != REG_SZ && type != REG_EXPAND_SZ))
			return def ? def : CString();
		TCHAR	    *cp = ret.GetBuffer(len / sizeof(TCHAR));
		if (::RegQueryValueEx(hKey, name, NULL, &type, (BYTE*)cp, &len) != ERROR_SUCCESS) {
			ret.ReleaseBuffer(0);
			return def ? def : CString();
		}
		ret.ReleaseBuffer();
		return ret;
	}
	///<summary>Get list item from combobox</summary>
	///<param name="hWnd">Combobox handle</param>
	///<param name="idx">Index of item</param>
	///<returns>List item string or empty string if bad index</returns>
	CString	GetCBString(HWND hWnd, int idx) {
		// get length at first
		LRESULT len = ::SendMessage(hWnd, CB_GETLBTEXTLEN, idx, 0);
		if (len == CB_ERR) //bad index
			return CString();

		CString ret;
		TCHAR	*cp = ret.GetBuffer(len + 1);
		len = ::SendMessage(hWnd, CB_GETLBTEXT, idx, (LPARAM)cp);
		if (len == CB_ERR)
			ret.ReleaseBuffer(0);
		else
			ret.ReleaseBuffer(len);
		return ret;
	}

	///<summary>Create font</summary>
	///<param name="sizept">Font height</param>
	///<param name="facename">Font facename</param>
	///<param name="fBold">Boldness</param>
	///<param name="fItalic">Italicness</param>
	///<returns>Handle to a logical font</returns>
	/*HFONT  CreatePtFont(int sizept, const TCHAR *facename, bool fBold, bool fItalic)
	{
		HDC	  hDC = CreateDC(L"DISPLAY", NULL, NULL, NULL);
		HFONT	  hFont = ::CreateFont(
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
			DEFAULT_PITCH | FF_DONTCARE,
			facename);
		DeleteDC(hDC);
		return hFont;
	}*/

	///<summary>Get Unicode code of character</summary>
	///<param name="ch">unicode character</param>
	///<returns>UNICODE code</returns>
/*	CString	GetCharName(int ch) {
		CString num;
		num.Format(_T("U+%04X"), ch);
		return num;
	}
	*/
/*
	///<summary>Get Full pathname</summary>
	///<param name="filename">File name</param>
	///<returns>Full path name</returns>
	CString	GetFullPathName(const CString& filename)
	{
		CString   ret;
		DWORD	    size = MAX_PATH;
		for (;;)
		{
			TCHAR	    *buf = ret.GetBuffer(size);
			TCHAR	    *final;
			DWORD nlen = ::GetFullPathNameW(filename, size, buf, &final);
			if (nlen == 0) // failed
				return filename;
			if (nlen > size)
			{
				// ret.ReleaseBuffer(0);
				size = nlen;
			}
			else
			{
				ret.ReleaseBuffer(nlen);
				break;
			}
		}
		return ret;
		/* Not needed for Windows 10
		// Convert to long form
		CString   r2;
		for (;;)
		{
		  TCHAR   *buf = r2.GetBuffer(size);
			DWORD	nlen = GetLongPathNameW(ret, buf, size);
			if (nlen==0)
			  return ret;
			if (nlen>len) {
			//  r2.ReleaseBuffer(0);
			  size = nlen;
			}
			else
			{
			  r2.ReleaseBuffer(nlen);
			  return r2;
			}
		}
	  */
	  /*
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
	}*/
	/*
	CString	UrlFromPath(const CString& path) {
	  CString   ret;
	  DWORD	    size = MAX_PATH*3;
	  TCHAR	    *cp = ret.GetBuffer(size);
	  if (FAILED(UrlCreateFromPath(path,cp,&size,0)))
		return CString();
	  ret.ReleaseBuffer(size);
	  ret.FreeExtra();
	  return ret;
	}
	*/


	///<summary>Get settings directory</summary>
	///<returns>Settings directory</returns>
	CString GetSettingsDir()
	{
		//shlobj_core.h (include Shlobj.h)
		PWSTR wszPath = NULL;
		if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_CREATE, NULL, &wszPath)))
		{
			CString appSettingPath(wszPath);
			appSettingPath += L"\\FBE\\";

			::CreateDirectory(appSettingPath, NULL);

			CoTaskMemFree(wszPath);
			return appSettingPath;
		}
		else
			return GetProgDir();
		/*
		wchar_t szPath[MAX_PATH];
		if (SUCCEEDED(SHGetFolderPath( NULL, CSIDL_APPDATA, NULL, 0, szPath )))
		{
			CString appSettingPath(szPath);
			appSettingPath += L"\\FBE\\";
			::CreateDirectory(appSettingPath, NULL);
			return appSettingPath;
		}
		else return GetProgDir();*/
	}

	///<summary>Get exe module directory</summary>
	///<returns>Programm directory</returns>
	CString	GetProgDir()
	{
		CString exedir;
		size_t size = MAX_PATH;
		DWORD r = 0;
		for (; ; )
		{
			TCHAR	*cp = exedir.GetBuffer(size);
			r = GetModuleFileName(_Module.GetModuleInstance(), cp, size);
			if (r < size && r != 0)
				break;
			if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
				size *= 2;
			else
				return NULL;
		}
		exedir.ReleaseBuffer(r);

		int pos = exedir.ReverseFind('\\');
		if (pos != -1)
		{
			exedir = exedir.Left(pos);
		}
		return exedir;
		/*
		  CString     exedir;
		  TCHAR	      *cp = exedir.GetBuffer(MAX_PATH);
		  DWORD len=::GetModuleFileName(_Module.GetModuleInstance(),cp,MAX_PATH);
		  exedir.ReleaseBuffer(len);
		  int	      p=len;
		  while (p>0 && exedir[p-1]!=_T('\\'))
			--p;
		  exedir.Delete(p,len-p);
		  return exedir;*/
	}

	/// <summary>Get file extension</summary>
	/// <param name="inStr">filename</param>  
	/// <returns>Extension (last chars after '.')</returns> 
	CString GetFileExtension(const CString& inStr)
	{
		CString file_extension = L"";
		int idx = inStr.ReverseFind(L'.');
		if (idx != -1) {
			file_extension = inStr.Right(inStr.GetLength() - idx - 1);
		}
		return file_extension;
	}

	/// <summary>Define string encoding</summary>
	/// <param name="filename">fb2 filename</param>  
	/// <returns>Encoding name or unknown</returns> 
	CString GetStringEncoding(const char * buffer)
	{
		CString enc(L"unknown");
		// Check first bytes at first
		if (buffer[0] == '\xEF' && buffer[1] == '\xBB' && buffer[2] == '\xBF')
		{
			enc = L"utf-8";
		}
		else if (buffer[0] == '\xFF' && buffer[1] == '\xFE')
		{
			enc = L"utf-16";
		}
		else if (buffer[0] == '\xFE' && buffer[1] == '\xFF')
		{
			enc = L"utf-16";
		}
		// Check encoding attribute in xml-header
		else
		{
			// get first line
			char str[] = "\r\n";
			int ii = strcspn(buffer, str);
			CString firstline(buffer, ii);

			firstline.MakeLower();
			int pos = firstline.Find(L"encoding");
			if (pos >= 0)
			{
				int pos2 = firstline.Find(L"\"", pos + 10);
				enc = firstline.Mid(pos + 10, pos2 - pos - 10);
			}
			/*

			enc = buffer;
			enc.MakeLower();
			int pos = enc.Find(L"encoding");
			if (pos >= 0)
			{
				int pos2 = enc.Find(L"\"", pos + 10);
				enc = enc.Mid(pos + 10, pos2 - pos - 10);
			}
			else 
				enc = L"";
				*/
		}
		return enc;
	}

	///<summary>Get directory for treeview scripts</summary>
	///<returns>Treeview scripts directory</returns>
	CString GetDocTReeScriptsDir()
	{
		return GetProgDir() + L"TreeCmd";
	}

	///<summary>Check if file exists</summary>
	///<param name="filename">Full filename</param>
	///<returns>true if exists</returns>
	bool IsFileExists(const CString& filename)
	{
		return ::GetFileAttributes(filename) != INVALID_FILE_ATTRIBUTES;
	}

	// prepare a default id for image file
	/// <summary>Prepare a default ID for image file</summary>
	/// <params name="path">File Path</params>
	/// <returns>File ID</returns>
	CString PrepareDefaultId(const CString &path)
	{

		CString fileName = path.Mid(path.ReverseFind('\\') + 1);
		fileName = L"_" + Transliterate(fileName);
		return fileName;
	}

	///<summary>Get full path to file in Prog directory</summary>
	///<param name="filename">Full filename</param>
	///<returns>Full path if file exists or filename if file isn't exists</returns>
	CString	GetProgDirFile(const CString& filename)
	{

		CString fullname = GetProgDir() + L"\\"+ filename;
		if (IsFileExists(fullname))
			return fullname;
		else
			return filename;
		/*CString     exedir;
		TCHAR	      *cp=exedir.GetBuffer(MAX_PATH);
		DWORD len=::GetModuleFileName(_Module.GetModuleInstance(),cp,MAX_PATH);
		exedir.ReleaseBuffer(len);
		int	      p=len;
		while (p>0 && exedir[p-1]!=_T('\\'))
		  --p;
		exedir.Delete(p,len-p);

		DWORD	      attr=::GetFileAttributes(tryname);
		if (attr!=INVALID_FILE_ATTRIBUTES)
		  goto ok;
		tryname=exedir+_T("..\\")+filename;
		attr=::GetFileAttributes(tryname);
		if (attr!=INVALID_FILE_ATTRIBUTES)
		  goto ok;
		tryname=filename;
	  ok:
		return GetFullPathName(tryname);*/
	}

	///<summary>Check string if it contains whitespaces only</summary>
	///<param name="spc">Null terminating string</param>
	///<returns>true if string contains whitespaces only</returns>
	bool is_whitespace(const wchar_t *spc) 
	{
		wchar_t nbsp = _Settings.GetNBSPChar()[0];
		while (*spc) {
			if (!iswspace(*spc) && *spc != nbsp)
				return false;
			++spc;
		}
		return true;
	}

	///<summary>Get Win-message for message id</summary>
	///<param name="code">message Id</param>
	///<returns>Message string</returns>
	CString	Win32ErrMsg(DWORD code)
	{
		CString ret;
		TCHAR *buf = ret.GetBuffer(1024);
		int len = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, code, 0, buf, 1024, NULL);
		ret.ReleaseBuffer(len);
		if (len == 0)
			ret.Format(L"Unknown error %x", code);
		return ret;
	}

	///<summary>Display error message dialog</summary>
	///<param name="Message">Message string</param>
	void ReportError(const CString& Message)
	{
		AtlTaskDialog(::GetActiveWindow(), IDS_ERRMSGBOX_CAPTION, (LPCTSTR)Message, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
	}

	///<summary>Display error message dialog</summary>
	///<param name="hr">HRESULT code</param>
	void  ReportError(HRESULT hr)
	{
		CString err(Win32ErrMsg(hr));
		ReportError(err);
	}

	///<summary>Display error message dialog</summary>
	///<param name="e">_com_error</param>
	void  ReportError(_com_error& e) {
		CString   err;
		HRESULT errCode = e.Error();
		// search for the VBScript error
		for (int i = 0; i < (sizeof(vberrs) / sizeof(VB_ERRORS)); i++)
			if (vberrs[i].Error == errCode)
			{
				err.Format(_T("Code: %08lx"), vberrs[i].Error);
				_bstr_t src(vberrs[i].Source);
				if (src.length() > 0)
				{
					err += _T("\nSource: ");
					err += (const wchar_t *)src;
				}
				src = vberrs[i].Description;
				if (src.length() > 0)
				{
					err += _T("\nDescription: ");
					err += (const wchar_t *)src;
				}
				break;
			}

		if (err.IsEmpty())
		{
			err.Format(_T("Code: %08lx [%s]"), e.Error(), e.ErrorMessage());
			_bstr_t src(e.Source());
			if (src.length() > 0) {
				err += _T("\nSource: ");
				err += (const wchar_t *)src;
			}
			src = e.Description();
			if (src.length() > 0) {
				err += _T("\nDescription: ");
				err += (const wchar_t *)src;
			}
		}
		ReportError(err);

		//AtlTaskDialog(::GetActiveWindow(), IDS_COM_ERR_CPT, (LPCTSTR)err, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
		VBErr = true;
	}

	CString	GetWindowText(HWND hWnd)
	{
		CString ret;
		CWindow wnd(hWnd);
		wnd.GetWindowText(ret);
		return ret;
	}


	MSXML2::IXMLDOMDocument2Ptr CreateDocument(bool fFreeThreaded)
	{
		MSXML2::IXMLDOMDocument2Ptr  doc;
		wchar_t *cls = fFreeThreaded ? L"Msxml2.FreeThreadedDOMDocument.6.0" : L"Msxml2.DOMDocument.6.0";
		CheckError(doc.CreateInstance(cls));
		return doc;
	}

	MSXML2::IXSLTemplatePtr CreateTemplate() {
		MSXML2::IXSLTemplatePtr    tp;
		CheckError(tp.CreateInstance(L"Msxml2.XSLTemplate.6.0"));
		return tp;
	}
	/*
	static void  ReportParseError(MSXML2::IXMLDOMDocument2Ptr doc)
	{
		try
		{
			MSXML2::IXMLDOMParseErrorPtr err(doc->parseError);
			long	  line = err->line;
			long	  col = err->linepos;
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
	*/
	// Not used??
	/*
	bool  LoadXml(MSXML2::IXMLDOMDocument2Ptr doc, const CString& url)
	{
		doc->put_async(VARIANT_FALSE);
		_variant_t  vturl((const TCHAR *)url);
		VARIANT_BOOL	flag;
		HRESULT   hr = doc->raw_load(vturl, &flag);
		if (FAILED(hr)) {
			ReportError(hr);
			return false;
		}
		if (flag != VARIANT_TRUE) {
			ReportParseError(doc);
			return false;
		}
		return true;
	}
	*/
	///<summary>Load image file to variant array</summary>
	///<param name="filename">Filename</param>
	///<param name="vt">Target variant array</param>
	///<returns>HRESULT code</returns>
	HRESULT LoadImageFile(const TCHAR* filename, VARIANT* vt)
	{
		HRESULT hr;

		HANDLE hFile = ::CreateFile(filename,
			GENERIC_READ,
			FILE_SHARE_READ,
			NULL,
			OPEN_EXISTING,
			FILE_FLAG_SEQUENTIAL_SCAN,
			NULL);

		if (hFile == INVALID_HANDLE_VALUE)
			return HRESULT_FROM_WIN32(::GetLastError());

		DWORD fsz = ::GetFileSize(hFile, NULL);
		if (fsz == INVALID_FILE_SIZE)
		{
			hr = HRESULT_FROM_WIN32(::GetLastError());
			CloseHandle(hFile);
			return hr;
		}

		SAFEARRAY* sa = ::SafeArrayCreateVector(VT_UI1, 0, fsz);
		if (sa == NULL)
		{
			hr = E_OUTOFMEMORY;
			CloseHandle(hFile);
			return hr;
		}

		void* pv;
		if (FAILED(hr = ::SafeArrayAccessData(sa, &pv)))
		{
			::SafeArrayDestroy(sa);
			CloseHandle(hFile);
			return hr;
		}

		DWORD nrd;
		BOOL fRd = ::ReadFile(hFile, pv, fsz, &nrd, NULL);
		DWORD err = ::GetLastError();

		::SafeArrayUnaccessData(sa);
		::CloseHandle(hFile);

		if (!fRd)
		{
			::SafeArrayDestroy(sa);
			return HRESULT_FROM_WIN32(err);
		}

		if (nrd != fsz)
		{
			::SafeArrayDestroy(sa);
			return E_FAIL;
		}

		V_VT(vt) = VT_ARRAY | VT_UI1;
		V_ARRAY(vt) = sa;

		return S_OK;
	}

	///<summary>Get image dimensions from file</summary>
	///<param name="pszFileName">Filename</param>
	///<param name="nWidth">Image width</param>
	///<param name="nHeight">Image height</param>
	///<returns>true if succcess</returns>
	bool GetImageDimsByPath(const wchar_t* pszFileName, int* nWidth, int* nHeight)
	{
		CImage image;

		if (image.Load(pszFileName) == S_OK)
		{
			if (nWidth)
				*nWidth = image.GetWidth();
			if (nHeight)
				*nHeight = image.GetHeight();
			return true;
		}
		return false;
	}

	///<summary>Get image dimensions from variant array</summary>
	///<param name="pszFileName">Filename</param>
	///<param name="nWidth">Image width</param>
	///<param name="nHeight">Image height</param>
	///<returns>true if succcess</returns>
	bool GetImageDimsByData(SAFEARRAY* data, ULONG length, int* nWidth, int* nHeight)
	{
		CImage image;
		bool result = false;

		long elnum = 0;
		void* pData = nullptr;

		if (::SafeArrayPtrOfIndex(data, &elnum, &pData) != S_OK)
			return false;

		IStream* pIStream;
		if (S_OK != ::CreateStreamOnHGlobal(pData, FALSE, &pIStream))
			return false;

		if (image.Load(pIStream) == S_OK)
		{
			if (nWidth)
				*nWidth = image.GetWidth();
			if (nHeight)
				*nHeight = image.GetHeight();
			return true;
		}
		pIStream->Release();

		return result;
	}

	///<summary>Convert name of keycode to keycode value</summary>
	///<param name="string">Name of keycode</param>
	///<returns>Keycode value or NULL if not found</returns>
	WORD StringToKeycode(CString string)
	{
		int key = keycodes.FindKey(string);
		if (key >= 0)
			return keycodes.GetValueAt(key);
		else 
            return NULL;
	}

	///<summary>Convert keycode value to name of keycode</summary>
	///<param name="keycode">Keycode value</param>
	///<returns>Name of keycode or empty string if not found</returns>
	CString KeycodeToString(WORD keycode)
	{
		int key = keycodes.FindVal(keycode);
		if (key >= 0)
			return keycodes.GetKeyAt(key);
		else 
            return CString();
	}

	///<summary>Get string representation of accelerator key</summary>
	///<param name="accel">Accelerator</param>
	///<returns>String representation or empty string if not found</returns>
	CString AccelToString(ACCEL accel)
	{
		CString temp(L"");
		if (accel.fVirt & FCONTROL)
			temp += "Ctrl+";
		if (accel.fVirt & FALT)
			temp += "Alt+";
		if (accel.fVirt & FSHIFT)
			temp += "Shift+";

		temp += KeycodeToString(accel.key);

		return temp;
	}

	///<summary>Get accelerator behavior of virtual keycode</summary>
	///<param name="virtkey">Virtual keycode</param>
	///<returns>Accelerator behavior or NULL if not found</returns>
	WORD VKToFVirt(WORD virtkey)
	{
		switch (virtkey)
		{
		case VK_MENU: // ALT keycode
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

	void InitSettingsHotkeyGroups()
	{
#ifndef NO_EXTRN_SETTINGS
		_Settings.InitHotkeyGroups();
#endif
	}

//---------------------- DomPath--------------------------------------------

///<summary>Fill path queue</summary>
///<param name="root">Topmost HTML node</param>
///<param name="endNode">Start HTML node</param>
///<returns>true-if success, false - if error</returns>
	bool DomPath::CreatePathFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root, MSHTML::IHTMLDOMNodePtr endNode)
	{
		// лепим путь к элементу начиная от него и раскручивая до корня
		m_path.clear();
		if (!root || !endNode)
			return false;
		
		if (endNode == root)
		{
			m_path.push_front(0);
			return true;
		}

		int id = 0;
		MSHTML::IHTMLDOMNodePtr parentNode = endNode->parentNode;
		MSHTML::IHTMLDOMNodePtr currentNode = endNode;

		if (U::scmp(currentNode->nodeName, L"#text") == 0)
			currentNode = currentNode->previousSibling;
		
		while (currentNode)
		{
			if (currentNode == root)
				return true;
			
			MSHTML::IHTMLDOMNodePtr previousSibling = currentNode->previousSibling;
			// count siblings
            while (previousSibling)
			{
				if (U::scmp(previousSibling->nodeName, L"#text") != 0)
					++id;
				previousSibling = previousSibling->previousSibling;
			}
			m_path.push_front(id);
			id = 0;
			currentNode = currentNode->parentNode;
		}
		m_path.clear();
		return false;
	}

///<summary>Get node using builded path</summary>
///<param name="root">Topmost XML node</param>
///<param name="endNode">Start XML node</param>
///<returns>XML node or NULL if not found</returns>
	MSXML2::IXMLDOMNodePtr DomPath::GetNodeFromXMLDOM(MSXML2::IXMLDOMNodePtr root)
	{
		if (!root || m_path.empty())
			return nullptr;

		MSXML2::IXMLDOMNodePtr currentNode = root;

		// int size = m_path.size();
		for (unsigned int i = 1; i < m_path.size(); ++i) //skip root
		{
			currentNode = currentNode->firstChild;
			if (!currentNode)
                return nullptr;
			CString ss1 = currentNode->nodeName;

			if (U::scmp(currentNode->nodeName, L"#text") == 0)
			{
				currentNode = currentNode->nextSibling;
				if (!currentNode)
					return nullptr;
			}

			CString ss2 = currentNode->nodeName;

			// int numb = m_path[i];
			for (int j = 0; j < m_path[i]; ++j)
			{
				currentNode = currentNode->nextSibling;
				CString ss3 = currentNode->nodeName;
				if (U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;

				if (!currentNode)
					return 0;
				CString ss4 = currentNode->nodeName;
			}
		}
		return currentNode;
	}

///<summary>Fill path queue</summary>
///<param name="root">Topmost XML node</param>
///<param name="endNode">Start XML node</param>
///<param name="char_pos">Position (out)</param>
///<returns>true-if success, false - if error</returns>
	bool DomPath::CreatePathFromXMLDOM(MSXML2::IXMLDOMNodePtr root, MSXML2::IXMLDOMNodePtr endNode)
	{
		m_path.clear();
		if (!root || !endNode)
			return false;

		if (endNode == root)
		{
			m_path.push_front(0);
			return true;
		}
		int id = 0;

		// from end Node climb up
		MSXML2::IXMLDOMNodePtr parentNode = endNode->parentNode;
		MSXML2::IXMLDOMNodePtr currentNode = endNode;

		if (U::scmp(currentNode->nodeName, L"#text") == 0)
		{
			currentNode = currentNode->previousSibling;
		}

		while (currentNode)
		{
			if (currentNode == root)
			{
				return true;
			}
			MSXML2::IXMLDOMNodePtr previousSibling = currentNode->previousSibling;
			while (previousSibling)
			{
				if (U::scmp(previousSibling->nodeName, L"#text") != 0)
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

///<summary>Get node using builded path</summary>
///<param name="root">Topmost HTML node</param>
///<param name="endNode">Start HTML node</param>
///<returns>HTML node or NULL if not found</returns>
	MSHTML::IHTMLDOMNodePtr DomPath::GetNodeFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root)
	{
		if (!root || m_path.empty())
			return nullptr;

		MSHTML::IHTMLDOMNodePtr currentNode = root;

		int size = m_path.size();
		for (int i = 0; i < size; ++i)
		{
			currentNode = currentNode->firstChild;
			if (currentNode)
				if (U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;

			//name = currentNode->nodeName;
			if (!currentNode)
                return nullptr;

			for (int j = 0; j < m_path[i]; ++j)
			{
				currentNode = currentNode->nextSibling;
				if (!currentNode) 
                    return nullptr;

				if (U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;
				if (!currentNode) 
                    return nullptr;
			}
		}
		return currentNode;
	}

// TO_DO Not used?
/*	MSHTML::IHTMLDOMNodePtr DomPath::FindSelectedNodeInXMLDOM(MSXML2::IXMLDOMNodePtr root)
	{
		// рекурсивно обходим дерево и ищем элемент с атрибутом selected		
		MSXML2::IXMLDOMElementPtr currentNode = root;
		while (currentNode)
		{
			if (MSXML2::IXMLDOMElementPtr(currentNode)->getAttribute(L"selected"))
				return currentNode;

			MSXML2::IXMLDOMElementPtr foundNode = FindSelectedNodeInXMLDOM(MSXML2::IXMLDOMNodePtr(currentNode)->firstChild);
			if (foundNode)
				return foundNode;
			currentNode = MSXML2::IXMLDOMNodePtr(currentNode)->nextSibling;
		}
		return nullptr;
	}
	*/
///<summary>Recursively build path queue</summary>
///<param name="xml">String for searching</param>
///<param name="open_tag_begin">Select position in string</param>
///<param name="char_pos">Position (out)</param>
///<returns>true-if success, false - if xml=NULL</returns>
	bool DomPath::CPFT(const wchar_t* xml, int pos, int* char_pos)
	{
		if (!xml)
		{
			return false;
		}

		const wchar_t* selpos = xml + pos;
		int virtual_pos = pos;
		// ищем открывающий тег
		int id = 0; // index of child element 

		wchar_t* open_tag_begin = nullptr;
		wchar_t* open_tag_end = nullptr;
		wchar_t* close_tag_begin = nullptr;
		wchar_t* close_tag_end = nullptr;

		bool ok = tGetFirstXMLNodeParams(xml, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);

		while (ok)
		{
			if (selpos > close_tag_end)
			{
				++id;
				virtual_pos = virtual_pos - (close_tag_end - open_tag_begin);
				ok = tGetFirstXMLNodeParams(close_tag_end, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end);
				continue;
			}
			// если курсор попал внутрь тега
			if ((selpos >= open_tag_begin && selpos < open_tag_end) || (selpos > close_tag_begin && selpos <= close_tag_end))
			{
				*char_pos = 0;
				m_path.push_back(id);
				return true;
			}

			// если курсор между открывающим и закрывающим тегом, то рекурсивно заходим в элементы
			if (selpos >= open_tag_end && selpos <= close_tag_begin)
			{
				m_path.push_back(id);
				id = 0;
				return this->CPFT(open_tag_end, selpos - open_tag_end, char_pos);
			}

			// если открывающий тег дальше курсора, значит мы нашли искомый элемент
			if (selpos < open_tag_begin)
			{
				break;
			}
		}

		// если закончились теги, значит мы нашли то что нужно
		*char_pos = virtual_pos;
		return true;
	}

///<summary>Fill path queue</summary>
///<param name="xml">String for searching</param>
///<param name="open_tag_begin">Select position in string</param>
///<param name="char_pos">Position (out)</param>
///<returns>true-if success, false - if xml=NULL</returns>
	bool DomPath::CreatePathFromText(const wchar_t* xml, int pos, int* char_pos)
	{
		m_path.clear();
		return this->CPFT(xml, pos, char_pos);
	}

	int DomPath::GetNodeFromText(wchar_t* xml, int char_pos)
	{
		int nest_len = m_path.size(); //queue depth

		if (!nest_len)
		{
			return 0;
		}

		wchar_t* curpos = 0;
		wchar_t* nested_node = xml;
		for (int i = 0; i < nest_len; ++i)
		{
			curpos = nested_node;
			int node_id = m_path[i];

			for (int j = 0; j < node_id; ++j)
			{
				if (!tGetFirstXMLNodeParams(curpos, 0, 0, 0, &curpos))
				{
					return 0;
				}
			}
			if (!tGetFirstXMLNodeParams(curpos, 0, &nested_node, 0, &curpos))
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

		while (ok)
		{
			int tmp = virtualcount + (open_tag_begin - curpos);
			if (tmp > char_pos)
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
    
///<summary>Get opening& closing xml tags</summary>
///<param name="xml">String for searching</param>
///<param name="open_tag_begin">Pointer to begin of opening tag (out)</param>
///<param name="open_tag_end">Pointer to end of opening tag (out)</param>
///<param name="close_tag_begin">Pointer to begin of closing tag (out)</param>
///<param name="close_tag_end">Pointer to end of closing tag (out)</param>
///<param name="xml">String for searching</param>
///<returns>true-if tag found, false -if not found</returns>
	bool tGetFirstXMLNodeParams(const wchar_t* xml, wchar_t** open_tag_begin, wchar_t** open_tag_end, wchar_t** close_tag_begin, wchar_t** close_tag_end)
	{
		if (!xml)
			return false;

		// serach first xml opening tag
        
        //search opening bracket
		wchar_t* optb = wcsstr(const_cast<wchar_t*>(xml), L"<");
        //skip closing tags & xml-header
		while (optb && (*(optb + 1) == L'/') || ((*(optb + 1) == L'?')))
		{
			optb = wcsstr(optb + 1, L"<");
		}
		if (!optb)
		{
			return false;
		}
		        
		// get tag name
        wchar_t* tag_name_start = optb + 1; // pointer to tag name
        
		// search closest space or closing bracket
		wchar_t* tag_name_end = wcschr(tag_name_start, L' '); //spaces are important for tags with attributes
		wchar_t* opte = wcschr(tag_name_start, L'>');

		if ((tag_name_end > opte) || (tag_name_end == 0))
		{
			tag_name_end = opte;
		}
		if (!tag_name_end)
		{
			return false;
		}
		++opte;
		wchar_t* tag_name = new wchar_t[tag_name_end - tag_name_start + 1];
		wcsncpy(tag_name, tag_name_start, tag_name_end - tag_name_start);
		tag_name[tag_name_end - tag_name_start] = L'\0';

        // Weknow now name. So will search closing tag
		wchar_t* cltb = 0;  //pointer to begin of closing tag
		wchar_t* clte = 0;  //pointer to end of closing tag

		if (*(opte - 2) == L'/')
		{
			//Found short form. No more search need
            cltb = opte;
			clte = opte;
		}
		else
		{
			// search begin of full closing tag
			if (!tFindCloseTag(opte, tag_name, &cltb))
			{
				delete[] tag_name;
				return false;
			}
            //search end of closing tag
			clte = wcschr(cltb, L'>');
			if (!clte)
			{
				delete[] tag_name;
				return false;
			}
			++clte;
		}
		delete[] tag_name;

		if (open_tag_begin)
		{
			*open_tag_begin = optb;
		}
		if (open_tag_end)
		{
			*open_tag_end = opte;
		}
		if (close_tag_begin)
		{
			*close_tag_begin = cltb;
		}
		if (close_tag_end)
		{
			*close_tag_end = clte;
		}

		return true;
	}

///<summary>Find closing xml tag</summary>
///<param name="start">String for searching</param>
///<param name="tag_name">Tag name</param>
///<param name="open_tag_end">Pointer to end of opening tag</param>
///<param name="close_tag">Pointer to begin of closing tag (out)</param>
///<returns>true-if tag found, false -if not found</returns>
	bool tFindCloseTag(wchar_t* start, wchar_t* tag_name, wchar_t** close_tag)
	{
		// added by SeNS: check space in tags
        //TO_DO Is possible spaces at all?
		CString tag(tag_name); 
        tag.Trim(); 
        if (tag.IsEmpty()) 
            return false;

		// Will skip pairs of <tag_name...></tag_name>
		//wchar_t* close_tagname
		wchar_t* next_tag = wcsstr(start, tag_name);
		while (next_tag)
		{
			// если это закрывающий тег, то значит мы нашли что надо
			if ((*(next_tag - 1) == L'/') && (*(next_tag - 2) == L'<'))
			{
				*close_tag = next_tag - 2;
				return true;
			}

			// проверяем таг ли это вообще
			if (*(next_tag - 1) != L'<')
			{
				next_tag = wcsstr(next_tag + wcslen(tag_name), tag_name);
				continue;
			}

			// это был открывающий тег
			// ищем его закрывающий тег
			if (!tFindCloseTag(next_tag + wcslen(tag_name), tag_name, &next_tag))
			{
				return false;
			}
			next_tag = wcsstr(next_tag + wcslen(tag_name), tag_name);
		}
		return false;
	}

	//TO_DO Not used??
/*	int CountTextNodeChars(wchar_t* node_begin, wchar_t* node_end)
	{
		wchar_t* openb = 0;
		wchar_t* opene = 0;
		wchar_t* closeb = 0;
		wchar_t* closee = 0;
		int count = 0;

		bool ok = tGetFirstXMLNodeParams(node_begin, &openb, &opene, &closeb, &closee);
		if (!ok)
		{
			return node_end - node_begin;
		}

		while (ok)
		{
			if (openb >= node_end)
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
	*/
///<summary>Check if HTML-element is parent of another</summary>
///<param name="elem">HTML element</param>
///<param name="parent">Possible parent element</param>
///<returns>true-if <parent> is parent element or elements are equal</returns>
	bool IsParentElement(MSHTML::IHTMLDOMNodePtr elem, MSHTML::IHTMLDOMNodePtr parent)
	{
		if (!parent || !elem)
		{
			return false;
		}
		if (parent == elem)
		{
			return true;
		}
		
        while (elem = elem->parentNode)
		{
			if (elem == parent)
			{
				return true;
			}
		}
		return false;
	}

///<summary>Check if XML-element is parent of another</summary>
///<param name="elem">XML element</param>
///<param name="parent">Possible parent element</param>
///<returns>true-if <parent> is parent element or elements are equal</returns>
	bool IsParentElement(MSXML2::IXMLDOMNodePtr elem, MSXML2::IXMLDOMNodePtr parent)
	{
		if (!parent || !elem)
		{
			return false;
		}
		if (parent == elem)
		{
			return true;
		}
		while (elem = elem->parentNode)
		{
			if (elem == parent)
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
	
///<summary>Find HTML element with class=title</summary>
///<param name="elem">HTML element for search inside</param>
///<returns>HTML element or NULL if not found</returns>
    MSHTML::IHTMLElementPtr	FindTitleNode(MSHTML::IHTMLDOMNodePtr elem)
	{
		MSHTML::IHTMLDOMNodePtr node(elem->firstChild);

		// _bstr_t   cls(MSHTML::IHTMLElementPtr(elem)->className);

		if (node && node->nodeType == 1 && U::scmp(node->nodeName, L"DIV") == 0) 
        {
			_bstr_t cls(MSHTML::IHTMLElementPtr(node)->className);
			if (U::scmp(cls, L"image") == 0) 
            {
				node = node->nextSibling;
				if (node->nodeType != 1 || U::scmp(node->nodeName, L"DIV"))
					return nullptr;
				cls = MSHTML::IHTMLElementPtr(node)->className;
			}
			if (U::scmp(cls, L"title") == 0)
				return MSHTML::IHTMLElementPtr(node);
		}
		return NULL;
	}

///<summary>Get inner text of HTML element with class=title</summary>
///<param name="elem">HTML element for search inside</param>
///<returns>HTML element or NULL if not found</returns>
	CString	FindTitle(MSHTML::IHTMLDOMNodePtr elem)
	{
		MSHTML::IHTMLElementPtr tn(FindTitleNode(elem));
		if (tn)
			return CString(tn->innerText.GetBSTR());
		return CString();
	}

///<summary>Get image href</summary>
///<param name="elem">HTML element with class=image</param>
///<returns>href value or empty string if no href attribute</returns>
	CString GetImageFileName(MSHTML::IHTMLDOMNodePtr elem)
	{
		_bstr_t cls(MSHTML::IHTMLElementPtr(elem)->className);
		if (U::scmp(cls, L"image") != 0)
			return L"";

		CString href(MSHTML::IHTMLElementPtr(elem)->getAttribute(L"href", 0));
		if (!href.GetLength())
			return L"";

		CString	name = href.Right(href.GetLength() - 1);
		return name;
	}


///<summary>Check is directory has subfolders</summary>
///<param name="path">Directory Path</param>
///<returns>true if has subfolders</returns>
	bool HasSubFolders(const CString path)
	{
		WIN32_FIND_DATA fd;
		HANDLE hFind = FindFirstFile(path + L"*.*", &fd);
		if (hFind != INVALID_HANDLE_VALUE)
		{
			do
			{
				if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					if (wcscmp(fd.cFileName, L".") && wcscmp(fd.cFileName, L".."))
					{
						FindClose(hFind);
						return true;
					}
				}
				else
					if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT))
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
		if (hFind != INVALID_HANDLE_VALUE && !(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
		{
			FindClose(hFind);
			return true;
		}
		else 
            return false;
	}

	bool HasScriptsEndpoint(const CString path, TCHAR* ext)
	{
		bool found = false;
		WIN32_FIND_DATA fd;
		HANDLE hFind = FindFirstFile(path + L"*.*", &fd);
		if (hFind == INVALID_HANDLE_VALUE)
		{
			return false;
		}
		else
		{
			if (HasFilesWithExt(path, ext))
			{
				FindClose(hFind);
				return true;
			}
			else
			{
				FindNextFile(hFind, &fd);
				do
				{
					if (wcscmp(fd.cFileName, L".") && wcscmp(fd.cFileName, L"..") && (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
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

		if (*(pos + 1) == 0)
			return false;

		*pos = 0;
		for (unsigned int i = 0; i < wcslen(Name); ++i)
		{
			if (Name[i] < '0' || Name[i] > '9')
				return false;
		}

		return true;
	}

///<summary>Transliterate string</summary>
///<param name="src">Source string</param>
///<returns>Transliterated string</returns>
	CString Transliterate(CString src)
	{
		CString dst = L"";
		CString from = L"\u0430\u0431\u0432\u0433\u0434\u0435\u0451\u0436\u0437\u0438\u0439\u043A\u043B"
						"\u043C\u043D\u043E\u043F\u0440\u0441\u0442\u0443\u0444\u0445\u0446\u0447\u0448"
						"\u0449\u044A\u044B\u044C\u044D\u044E\u044F\u0454\u0456\u0457\u0410\u0411\u0412"
						"\u0413\u0414\u0415\u0401\u0416\u0417\u0418\u0419\u041A\u041B\u041C\u041D\u041E"
						"\u041F\u0420\u0421\u0422\u0423\u0424\u0425\u0426\u0427\u0428\u0429\u042A\u042B"
						"\u042C\u042D\u042E\u042F\u0404\u0406\u0407";
		CString to[] = { L"a", L"b", L"v", L"g", L"d", L"e", L"jo", L"zh", L"z", L"i", L"jj", L"k",
						 L"l", L"m", L"n", L"o", L"p", L"r", L"s", L"t", L"u", L"f", L"h", L"c", L"ch",
						 L"sh", L"shh", L"\"", L"y", L"'", L"eh", L"ju", L"ja", L"je", L"i", L"i'",
						 L"A", L"B", L"V", L"G", L"D", L"E", L"JO", L"ZH", L"Z", L"I", L"JJ", L"K",
						 L"L", L"M", L"N", L"O", L"P", L"R", L"S", L"T", L"U", L"F", L"H", L"C", L"CH",
						 L"SH", L"SHH", L"\"", L"Y", L"'", L"EH", L"JU", L"JA", L"JE", L"I", L"I'" };

		for (int i = 0; i < src.GetLength(); i++)
		{
			int j = from.Find(src.GetAt(i));
			if (j < 0) 
                dst += src.GetAt(i);
            else 
                dst += to[j];
		}
		return dst;
	}

	// a getopt implementation
	static int   xgetopt(
		CSimpleArray<CString>& argv,
		const TCHAR *ospec,
		int&	    argp,
		const TCHAR *&state,
		const TCHAR *&arg)
	{
		const TCHAR	  *cp;
		TCHAR		  opt;

		if (!state || !state[0]) { // look a the next arg
			if (argp >= argv.GetSize() || argv[argp][0] != '-') // no more options
				return 0;
			if (!argv[argp][1]) // a lone '-', treat as an end of list
				return 0;
			if (argv[argp][1] == '-') { // '--', ignore rest of text and stop
				++argp;
				return 0;
			}
			state = (const TCHAR *)argv[argp] + 1;
			++argp;
		}
		// we are in a middle of an arg
		opt = (unsigned)*state++;
		bool found = false;
		for (cp = ospec; *cp; ++cp) {
			if (*cp == opt)
			{
				found = true;
				break;
			}
			if (cp[1] == ':')
				++cp;
		}
		if (!found)
		{
			CString strMessage;
			strMessage.Format(IDS_INVALID_CML_MSG, opt);
			AtlTaskDialog(::GetActiveWindow(), IDS_ERRMSGBOX_CAPTION, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
			return -1; // error
		}

		if (cp[1] == ':') { // option requires an argument
			if (*state) { // use rest of string
				arg = state;
				state = NULL;
				return opt;
			}
			// use next arg if available
			if (argp < argv.GetSize()) {
				arg = argv[argp];
				++argp;
				return opt;
			}
			// barf about missing args
			CString strMessage;
			strMessage.Format(IDS_CML_ARGS_MSG, opt);
			AtlTaskDialog(::GetActiveWindow(), IDS_ERRMSGBOX_CAPTION, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
			return -1;
		}
		// just return current option
		return opt;
	}

	CmdLineArgs   _ARGS;

	bool  ParseCmdLineArgs() {
		const TCHAR	*arg, *state = NULL;
		int		argp = 0;
		int		ch;
		for (;;) {
			switch ((ch = xgetopt(_ARGV, _T("d"), argp, state, arg))) {
			case 0: // end of options
				while (argp--)
					_ARGV.RemoveAt(0);
				return true;
			case _T('d'):
				_ARGS.start_in_desc_mode = true;
				break;
			case -1: // error
				return false;
				// just ignore options for now :)
			}
		}
	}

	// an input box
	class CInputBox : public CDialogImpl<CInputBox>,
		public CWinDataExchange<CInputBox>
	{
	public:
		enum { IDD = IDD_INPUTBOX };
		CString	m_text;
		const wchar_t	*m_title;
		const wchar_t	*m_prompt;

		CInputBox(const wchar_t *title, const wchar_t *prompt) : m_title(title), m_prompt(prompt) { }

		BEGIN_DDX_MAP(CInputBox)
			DDX_TEXT(IDC_INPUT, m_text)
		END_DDX_MAP()

		BEGIN_MSG_MAP(CInputBox)
			COMMAND_ID_HANDLER(IDYES, OnYes)
			COMMAND_ID_HANDLER(IDNO, OnCancel)
			COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
			MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		END_MSG_MAP()

		LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
			DoDataExchange(FALSE);  // Populate the controls
			SetDlgItemText(IDC_PROMPT, m_prompt);
			SetWindowText(m_title);
			if (dialogRect.left != -1)
				::SetWindowPos(m_hWnd, HWND_TOPMOST, dialogRect.left, dialogRect.top, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
			return 0;
		}

		BOOL EndDialog(int nRetCode)
		{
			::GetWindowRect(m_hWnd, &dialogRect);
			return ::EndDialog(m_hWnd, nRetCode);
		}

		LRESULT OnYes(WORD, WORD wID, HWND, BOOL&) {
			DoDataExchange(TRUE);  // Populate the data members
			EndDialog(wID);
			return 0L;
		}

		LRESULT OnCancel(WORD, WORD wID, HWND, BOOL&) {
			EndDialog(wID);
			return 0L;
		}
	};

///<summary>Show input dialog and return enetered value</summary>
///<param name="result">Value as string (IN/OUT)</param>
///<param name="title">Dialog title</param>
///<param name="prompt">Input prompt</param>
///<returns>Dialog result</returns>
	int	InputBox(CString& result, const wchar_t *title, const wchar_t *prompt) 
    {
		CInputBox   dlg(title, prompt);
		dlg.m_text = result;
		int dlgResult = dlg.DoModal();
		if (dlgResult == IDYES) 
            result = dlg.m_text;
		return dlgResult;
	}

	// html
	CString GetAttrCS(MSHTML::IHTMLElement *elem, const wchar_t *attr) {
		if (!elem) 
            return L"";
		_variant_t va(elem->getAttribute(attr, 2));
		if (V_VT(&va) != VT_BSTR)
			return CString();
		return V_BSTR(&va);
	}

	_bstr_t GetAttrB(MSHTML::IHTMLElement *elem, const wchar_t *attr) {
		if (!elem) 
            return L"";
		_variant_t	    va(elem->getAttribute(attr, 2));
		if (V_VT(&va) != VT_BSTR)
			return _bstr_t();
		return _bstr_t(va.Detach().bstrVal, false);
	}

	char	*ToUtf8(const CString& s, int& patlen) {
		DWORD   len = ::WideCharToMultiByte(CP_UTF8, 0, s, s.GetLength(), NULL, 0, NULL, NULL);
		char    *tmp = (char *)malloc(len + 1);
		if (tmp) 
        {
			::WideCharToMultiByte(CP_UTF8, 0, s, s.GetLength(),	tmp, len, NULL, NULL);
			tmp[len] = '\0';
		}
		patlen = len;
		return tmp;
	}

#ifdef USE_PCRE
	// Regexp constructor
	IMatchCollection* IRegExp2::Execute(CString sourceString)
	{
#define OVECCOUNT 300

		pcre *re;
		const char *error;
		bool is_error = true;
		int options;
		int erroffset;
		int ovector[OVECCOUNT];
		int subject_length;
		int rc, offset, char_offset;
		IMatchCollection* matches;
		// fix for issue #145
		char dst[0xFFFF];

		matches = new IMatchCollection();

		options = IgnoreCase ? PCRE_CASELESS : 0;
		options |= PCRE_UTF8;

		// convert pattern to UTF-8
		CT2A pat(m_pattern, CP_UTF8);

		re = pcre_compile(
			pat,				  /* the pattern */
			options,              /* default options */
			&error,               /* for error message */
			&erroffset,           /* for error offset */
			NULL);                /* use default character tables */

		if (re)
		{
			is_error = false;
			CT2A subj(sourceString, CP_UTF8);
			subject_length = strlen(subj);

			offset = char_offset = 0;

			do
			{
				rc = pcre_exec(
					re,                   /* the compiled pattern */
					NULL,                 /* no extra data - we didn't study the pattern */
					subj,					/* the subject string */
					subject_length,       /* the length of the subject */
					offset,                /* start at offset 0 in the subject */
					0,					/* default options */
					ovector,              /* output vector for substring information */
					OVECCOUNT);           /* number of elements in the output vector */

				if (rc > 0)
				{
					// add match
					char *substring_start = subj + ovector[0];
					int substring_length = ovector[1] - ovector[0];
					if (substring_length < sizeof(dst))
					{
						// convert substring to Unicode
						strncpy(dst, substring_start, substring_length);
						dst[substring_length] = '\0';
						// calculate character position
						while (offset < ovector[0])
						{
							offset += UTF8_CHAR_LEN(subj[offset]);
							char_offset++;
						}

						CString str = CString(CA2T(dst, CP_UTF8));
						IMatch2* item = new IMatch2(str, char_offset);

						// add submatches (including match)
						for (int i = 1; i < rc; i++)
						{
							substring_start = subj + ovector[i * 2];
							substring_length = ovector[i * 2 + 1] - ovector[i * 2];
							if (substring_length < sizeof(dst))
							{
								// convert substring to Unicode
								strncpy(dst, substring_start, substring_length);
								dst[substring_length] = '\0';
								item->AddSubMatch(CString(CA2T(dst, CP_UTF8)));
							}
						}
						matches->AddItem(item);
						char_offset += str.GetLength();
						offset = ovector[1];
						// empty line
						if (ovector[0] == ovector[1])
						{
							offset++;
							char_offset++;
						}
					}
				}
			} while (rc > 0);

			pcre_free(re);     /* Release memory used for the compiled pattern */
		}

		if (is_error)
		{
			// Raise COM-compatible exception
			ICreateErrorInfoPtr cerrinf = 0;
			if (::CreateErrorInfo(&cerrinf) == S_OK)
			{
				cerrinf->SetDescription(CString(error).AllocSysString());
				cerrinf->SetSource(L"Perl Compatible Regular Expressions");
				cerrinf->SetGUID(GUID_NULL);
				IErrorInfoPtr ei = cerrinf;
				throw _com_error(MK_E_SYNTAX, ei, true);
			}
		}

		return matches;
	}
#endif

}
