// MainFrm.cpp : implmentation of the CMainFrame class
#include "stdafx.h"

#include "mainfrm.h"
#include "AboutBox.h"
#include "SettingsDlg.h"
#include "xmlMatchedTagsHighlighter.h"

// Vista file dialogs client GUIDs
// {CF7C097D-0A2D-47EF-939D-1760BF4D0154}
static const GUID GUID_FB2Dialog =
{
	0xcf7c097d, 0xa2d, 0x47ef, {0x93, 0x9d, 0x17, 0x60, 0xbf, 0x4d, 0x1, 0x54} };
static const int SCRIPT_MENU_ORDER = 6;

LRESULT CCustomEdit::OnChar(UINT, WPARAM wParam, LPARAM, BOOL &bHandled)
{
	if (wParam == VK_RETURN)
		::PostMessage(::GetParent(GetParent()), WM_COMMAND, MAKELONG(GetDlgCtrlID(), IDN_ED_RETURN), (LPARAM)m_hWnd);

	bHandled = FALSE;
	return S_OK;
}

void CCustomStatic::DoPaint(CDCHandle dc)
{
	RECT rc;
	GetClientRect(&rc);

	HFONT oldFont = (HFONT)SelectObject(dc, m_font);

	UINT iFlags = DT_SINGLELINE | DT_CENTER | DT_VCENTER;

	int len = GetWindowTextLength();
	wchar_t *text = new wchar_t[len + 1];
	GetWindowText(text, len + 1);

	dc.SetBkMode(TRANSPARENT);
	if (m_enabled)
	{
		dc.SetTextColor(GetSysColor(COLOR_BTNTEXT));
	}
	else
	{
		dc.SetTextColor(GetSysColor(COLOR_GRAYTEXT));
	}
	dc.DrawText(text, -1, &rc, iFlags);
	SelectObject(dc, oldFont);
	delete[] text;
}

LRESULT CCustomStatic::OnPaint(UINT, WPARAM wParam, LPARAM, BOOL &)
{
	if (wParam != NULL)
	{
		DoPaint((HDC)wParam);
	}
	else
	{
		CPaintDC dc(m_hWnd);
		DoPaint(dc.m_hDC);
	}
	return S_OK;
}

void CCustomStatic::SetFont(HFONT pFont)
{
	m_font = pFont;
}

void CCustomStatic::SetEnabled(bool Enabled = true)
{
	m_enabled = Enabled;
	Invalidate();
}

// search&replace in scintilla
CString SciSelection(CWindow source)
{
	int start = source.SendMessage(SCI_GETSELECTIONSTART);
	int end = source.SendMessage(SCI_GETSELECTIONEND);

	if (start >= end)
		return CString();

	char *buffer = (char *)malloc(end - start + 1);
	if (buffer == NULL)
		return CString();
	source.SendMessage(SCI_GETSELTEXT, 0, (LPARAM)buffer);

	char *p = buffer;
	while (*p && *p != '\r' && *p != '\n')
		++p;

	int wlen = ::MultiByteToWideChar(CP_UTF8, 0, buffer, p - buffer, NULL, 0);
	CString ret;
	wchar_t *wp = ret.GetBuffer(wlen);
	if (wp)
		::MultiByteToWideChar(CP_UTF8, 0, buffer, p - buffer, wp, wlen);
	free(buffer);
	if (wp)
		ret.ReleaseBuffer(wlen);
	return ret;
}

class CSciFindDlg : public CFindDlgBase
{
public:
	CWindow m_source;

	CSciFindDlg(CFBEView *view, HWND src) : CFindDlgBase(view), m_source(src)
	{
	}
	void UpdatePattern()
	{
		m_view->m_fo.pattern = SciSelection(m_source);
	}

	virtual void DoFind()
	{
		GetData();
		if (m_view->SciFindNext(m_source, false, true))
		{
			SaveString();
			SaveHistory();
		}
	}
};

class CSciReplaceDlg : public CReplaceDlgBase
{
public:
	CWindow m_source;

	CSciReplaceDlg(CFBEView *view, HWND src) : CReplaceDlgBase(view), m_source(src)
	{
	}

	void UpdatePattern()
	{
		m_view->m_fo.pattern = SciSelection(m_source);
	}

	virtual void DoFind()
	{
		if (!m_view->SciFindNext(m_source, true, false))
		{
			CString strMessage;
			strMessage.Format(IDS_SEARCH_END_MSG, (LPCTSTR)m_view->m_fo.pattern);
			AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_WARNING_ICON);
		}
		else
		{
			SaveString();
			SaveHistory();
			m_selvalid = true;
			MakeClose();
		}
	}
	virtual void DoReplace()
	{
		//TO DO repeat search in selection to refill patterns if search by regexp (??)
		if (m_selvalid)
		{ // replace
			m_source.SendMessage(SCI_TARGETFROMSELECTION);
			DWORD len = ::WideCharToMultiByte(CP_UTF8, 0,
				m_view->m_fo.replacement, m_view->m_fo.replacement.GetLength(),
				NULL, 0, NULL, NULL);
			char *tmp = (char *)malloc(len + 1);
			if (tmp)
			{
				::WideCharToMultiByte(CP_UTF8, 0,
					m_view->m_fo.replacement, m_view->m_fo.replacement.GetLength(),
					tmp, len, NULL, NULL);
				tmp[len] = '\0';
				if (m_view->m_fo.fRegexp)
					m_source.SendMessage(SCI_REPLACETARGETRE, len, (LPARAM)tmp);
				else
					m_source.SendMessage(SCI_REPLACETARGET, len, (LPARAM)tmp);
				free(tmp);
			}
			m_selvalid = false;
		}
		DoFind();
	}
	virtual void DoReplaceAll()
	{
		if (m_view->m_fo.pattern.IsEmpty())
			return;

		// setup search flags
		int flags = 0;
		if (m_view->m_fo.flags & CFBEView::FRF_WHOLE)
			flags |= SCFIND_WHOLEWORD;
		if (m_view->m_fo.flags & CFBEView::FRF_CASE)
			flags |= SCFIND_MATCHCASE;
		if (m_view->m_fo.fRegexp)
			flags |= SCFIND_REGEXP | SCFIND_REGEXP;
		m_source.SendMessage(SCI_SETSEARCHFLAGS, flags, 0);

		// setup target range
		int end = m_source.SendMessage(SCI_GETLENGTH);
		m_source.SendMessage(SCI_SETTARGETSTART, 0);
		m_source.SendMessage(SCI_SETTARGETEND, end);

		// convert search pattern and replacement to utf8
		int patlen = 0, num_pat_nbsp = 0, num_rep_nbsp = 0;
		// added by SeNS
		if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
			num_pat_nbsp = m_view->m_fo.pattern.Replace(L"\u00A0", _Settings.GetNBSPChar());
		char *pattern = U::ToUtf8(m_view->m_fo.pattern, patlen);
		if (pattern == NULL)
			return;
		int replen;
		// added by SeNS
		if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
			num_rep_nbsp = m_view->m_fo.replacement.Replace(L"\u00A0", _Settings.GetNBSPChar());
		char *replacement = U::ToUtf8(m_view->m_fo.replacement, replen);
		if (replacement == NULL)
		{
			free(pattern);
			return;
		}

		// find first match
		int pos = m_source.SendMessage(SCI_SEARCHINTARGET, patlen, (LPARAM)pattern);

		int num_repl = 0;

		if (pos != -1 && pos <= end)
		{
			int last_match = pos;

			m_source.SendMessage(SCI_BEGINUNDOACTION);
			while (pos != -1)
			{
				int matchlen = m_source.SendMessage(SCI_GETTARGETEND) - m_source.SendMessage(SCI_GETTARGETSTART);
				matchlen -= num_pat_nbsp * 2;

				int mvp = 0;
				if (matchlen <= 0)
				{
					char ch = (char)m_source.SendMessage(SCI_GETCHARAT, m_source.SendMessage(SCI_GETTARGETEND));
					if (ch == '\r' || ch == '\n')
						mvp = 1;
				}
				int rlen = matchlen;
				if (m_view->m_fo.fRegexp)
					rlen = m_source.SendMessage(SCI_REPLACETARGETRE, replen, (LPARAM)replacement);
				else
					m_source.SendMessage(SCI_REPLACETARGET, replen, (LPARAM)replacement);

				end += rlen - matchlen;
				last_match = pos + rlen + mvp + num_rep_nbsp * 2;
				if (last_match >= end)
					pos = -1;
				else
				{
					m_source.SendMessage(SCI_SETTARGETSTART, last_match);
					m_source.SendMessage(SCI_SETTARGETEND, end);
					pos = m_source.SendMessage(SCI_SEARCHINTARGET, patlen, (LPARAM)pattern);
				}
				++num_repl;
			}
			m_source.SendMessage(SCI_ENDUNDOACTION);
		}

		free(pattern);
		free(replacement);

		CString strMessage;
		if (num_repl > 0)
		{
			SaveString();
			SaveHistory();

			strMessage.Format(IDS_REPL_DONE_MSG, num_repl);
			AtlTaskDialog(*this, IDS_REPL_ALL_CAPT, (LPCTSTR)strMessage, (LPCTSTR)NULL);
			MakeClose();
			m_selvalid = false;
		}
		else
		{
			strMessage.Format(IDS_SEARCH_END_MSG, (LPCTSTR)m_view->m_fo.pattern);
			AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_WARNING_ICON);
		}
	}
};

//--------------- CCustomSaveDialog----------------------
LRESULT CCustomSaveDialog::OnInitDialog(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	m_hDlg = hWnd;

	TCHAR buf[1024];

	if (::LoadString(_Module.GetResourceInstance(), IDS_ENCODINGS, buf, sizeof(buf) / sizeof(buf[0])) == 0)
		return TRUE;

	TCHAR *cp = buf;
	while (*cp)
	{
		size_t len = _tcscspn(cp, L",");
		if (cp[len])
			cp[len++] = L'\0';
		if (*cp)
			::SendDlgItemMessage(hWnd, IDC_ENCODING, CB_ADDSTRING, 0, (LPARAM)cp);
		cp += len;
	}

	//::SetDlgItemText(hWnd,IDC_ENCODING,m_encoding);
	::SendMessage(::GetDlgItem(hWnd, IDC_ENCODING), CB_SELECTSTRING, 0, (LPARAM)m_encoding.GetBuffer());

	return TRUE;
}

LRESULT CCustomSaveDialog::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	// make combobox the same size as std controls
	RECT rc_std, rc_my, rc_static, rc_static_my;
	HWND hCB = ::GetDlgItem(m_hDlg, IDC_ENCODING);
	HWND hST = ::GetDlgItem(m_hDlg, IDC_STATIC);
	::GetWindowRect(hCB, &rc_my);
	::GetWindowRect(GetFileDialogWindow().GetDlgItem(cmb1), &rc_std);

	::GetWindowRect(hST, &rc_static_my);
	::GetWindowRect(GetFileDialogWindow().GetDlgItem(stc2), &rc_static);

	POINT pt = { rc_std.left, rc_my.top };
	POINT pt1 = { rc_static.left, rc_static_my.top };
	::ScreenToClient(m_hDlg, &pt);
	::ScreenToClient(m_hDlg, &pt1);

	::MoveWindow(hST, pt1.x, pt1.y, rc_static_my.right - rc_static_my.left, rc_my.bottom - rc_my.top, TRUE);
	::MoveWindow(hCB, pt.x, pt.y, rc_std.right - rc_std.left, rc_my.bottom - rc_my.top, TRUE);

	return S_OK;
}

BOOL CCustomSaveDialog::OnFileOK(LPOFNOTIFY on)
{
	m_encoding = U::GetWindowText(::GetDlgItem(m_hDlg, IDC_ENCODING));
	return TRUE;
}

//--------------- CMainFrame----------------------
CMainFrame::CMainFrame() : m_doc(0), m_cb_updated(false),
m_sel_changed(false), m_want_focus(0),
m_ignore_cb_changes(false), m_incsearch(0), m_cb_last_images(false),
m_last_ie_ovr(true), m_last_sci_ovr(true), m_sci_find_dlg(0), m_sci_replace_dlg(0),
m_last_script(0), m_last_plugin(0), m_restore_pos_cmdline(false), m_strSuggestions(nullptr)
{
	m_Speller = new CSpeller(U::GetProgDir() + L"\\dict\\");
}

CMainFrame::~CMainFrame()
{
	delete m_doc;
	if (m_sci_find_dlg)
	{
		delete m_sci_find_dlg;
	}
	delete m_Speller;
}

// utility methods
bool CMainFrame::IsBandVisible(int id)
{
	int nBandIndex = m_rebar.IdToIndex(id);
	REBARBANDINFO rbi;
	rbi.cbSize = sizeof(rbi);
	rbi.fMask = RBBIM_STYLE;
	m_rebar.GetBandInfo(nBandIndex, &rbi);
	return (rbi.fStyle & RBBS_HIDDEN) == 0;
}

///<summary>Open System open dialog and return filename</summary>
CString CMainFrame::GetOpenFileName()
{
	CString strFileName;
	if (RunTimeHelper::IsVista())
	{
		const COMDLG_FILTERSPEC arrFilterSpec[] =
		{
			{L"FictionBook files", L"*.fb2"},
			{L"FB2ZIP files", L"*.zip"},
			{L"All files", L"*.*"} };
		CShellFileOpenDialog dlg(NULL, FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_FILEMUSTEXIST, L"fb2",
			arrFilterSpec, ARRAYSIZE(arrFilterSpec));
		dlg.GetPtr()->SetClientGuid(GUID_FB2Dialog);
		if (dlg.DoModal() == IDOK)
		{
			dlg.GetFilePath(strFileName);
		}
	}
	else
	{

		CFileDialog dlg(TRUE, L"fb2", NULL, OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_EXPLORER,
			L"FictionBook files (*.fb2)\0*.fb2\0FB2ZIP files (*.zip)\0*.fb2.zip\0All files (*.*)\0*.*\0\0");
		dlg.m_ofn.Flags &= ~OFN_ENABLEHOOK;
		dlg.m_ofn.lpfnHook = NULL;
		if (dlg.DoModal(*this) == IDOK)
			strFileName = dlg.m_szFileName;
	}
	return strFileName;
}

///<summary>Open System Save dialog and return filename and encoding!</summary>
CString CMainFrame::GetSaveFileName(CString &encoding)
{
	CString strFileName = m_doc->m_filename;

	const COMDLG_FILTERSPEC arrFilterSpec[] =
	{
		{L"FictionBook files", L"*.fb2"},
		{L"FB2ZIP files", L"*.zip"},
		{L"All files", L"*.*"} };

	CShellFileSaveDialog dlg(NULL, FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT,
		(U::GetFileExtension(m_doc->m_filename) == L"zip") ? L"zip" : L"fb2", arrFilterSpec, ARRAYSIZE(arrFilterSpec));
	dlg.GetPtr()->SetClientGuid(GUID_FB2Dialog);

	CComPtr<IFileDialog> spFileDialog;
	dlg.GetPtr()->QueryInterface(&spFileDialog);
	spFileDialog->SetFileTypeIndex((U::GetFileExtension(m_doc->m_filename) == L"zip") ? 2 : 1);

	CComPtr<IFileDialogCustomize> spFileDialogCustomize;
	HRESULT hr = dlg.GetPtr()->QueryInterface(&spFileDialogCustomize);
	if (SUCCEEDED(hr))
	{
		CString strTitle;
		strTitle.LoadString(IDS_ENCODING);
		spFileDialogCustomize->StartVisualGroup(1000, strTitle);
		spFileDialogCustomize->AddComboBox(1001);

		CString strEncodings;
		strEncodings.LoadString(IDS_ENCODINGS);
		CSimpleArray<CString> lstEncodings;
		CString strEncoding;
		int iStart = 0;
		do
		{
			strEncoding = strEncodings.Tokenize(L",", iStart);
			if (strEncoding.IsEmpty())
				break;
			lstEncodings.Add(strEncoding);
		} while (true);

		for (int i = 0; i < lstEncodings.GetSize(); i++)
		{
			spFileDialogCustomize->AddControlItem(1001, 1100 + i, lstEncodings[i]);
		}

		CString strSelectedEncodding = _Settings.m_keep_encoding ? m_doc->m_file_encoding : _Settings.GetDefaultEncoding();
		int nEncodingIndex = lstEncodings.Find(strSelectedEncodding.MakeLower());
		spFileDialogCustomize->SetSelectedControlItem(1001, 1100 + nEncodingIndex);

		spFileDialogCustomize->EndVisualGroup();

		//For new files set prompt to Untitled.fb2
		//For exists files use IShellItem
		if (!m_spShellItem)
			dlg.GetPtr()->SetFileName(strFileName);
		else
			hr = dlg.GetPtr()->SetSaveAsItem(m_spShellItem);
		if (dlg.DoModal() == IDOK)
		{
			dlg.GetFilePath(strFileName);
			m_spShellItem.Release();
			dlg.GetPtr()->GetResult(&m_spShellItem);
			CComHeapPtr<WCHAR> szSelectedEncoding;
			DWORD dwItem;
			spFileDialogCustomize->GetSelectedControlItem(1001, &dwItem);
			encoding = lstEncodings[dwItem - 1100];
		}
		else
			strFileName = L"";
	}
	return strFileName;
}

///<summary>If doc was changed ask user about saving</summary>
///<returns>true - discard allowed, false - no discard allowed</summary>
bool CMainFrame::DiscardChanges()
{
	if (m_doc->IsChanged())
	{
		CString strMessage;
		strMessage.Format(IDS_SAVE_DLG_MSG, (LPCTSTR)m_doc->m_filename);
		switch (AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL,
			TDCBF_YES_BUTTON | TDCBF_NO_BUTTON | TDCBF_CANCEL_BUTTON, TD_WARNING_ICON))
		{
		case IDYES:
			return (SaveFile(false) == FILE_OP_STATUS::OK);
		case IDNO:
			return true;
		case IDCANCEL:
			return false;
		}
	}
	return true;
}

void CMainFrame::AttachDocument(FB::Doc *doc)
{
	m_view.AttachWnd(doc->m_body);
	m_current_view = DESC;
	m_last_view = BODY;
	m_last_ctrl_tab_view = BODY;
	m_cb_updated = false;
	m_need_title_update = m_sel_changed = true;
	m_Speller->AttachDocument(doc->m_body.Document());

	//	MSHTML::IHTMLElementPtr fbwBody = MSHTML::IHTMLDocument3Ptr(m_doc->m_body.Document())->getElementById(L"fbw_body");
	//	MSHTML::IHTMLDOMNodePtr el = MSHTML::IHTMLDOMNodePtr(fbwBody)->firstChild;

		// Set selection to first P
	MSHTML::IHTMLElementPtr el = MSHTML::IElementSelectorPtr(m_doc->m_body.Document()->body)->querySelector(L"div#fbw_body p");
	m_doc->m_body.GoTo(el);

	ShowView(BODY);
	if (_Settings.ViewDocumentTree())
	{
		GetDocumentStructure();
		m_document_tree.HighlightItemAtPos(doc->m_body.SelectionContainer());
	}

	m_doc->m_body.ScrollToSelection();
	m_doc->m_body.SetFocus();

	m_doc->m_body.Synched(false);
	m_doc->Changed(false);
	m_doc->m_body.Init();
	m_doc->m_body.TraceChanges(true);
}

///<summary>save doc in file</summary>
///<params name="askname">true=show FileSave dialog</summary>
///<returns>FILE_OP_STATUS</summary>
CMainFrame::FILE_OP_STATUS CMainFrame::SaveFile(bool askname)
{
	ATLASSERT(m_doc != NULL);

	if (!m_doc->m_body.IsSynched())
	{
		m_doc->CreateDOM();
	}
	int err_col = 0;
	int err_line = 0;
	CString ErrMsg;
	if (!m_doc->Validate(err_col, err_line, ErrMsg))
	{
		if (m_current_view != SOURCE)
			ShowView(SOURCE);
		SourceGoTo(err_line, err_col);
		DisplayXMLError(ErrMsg);
		return FAIL;
	}

	CString encoding = m_doc->m_file_encoding;
	CString filename(L"");

	if (askname || m_doc->GetDocFileName() == L"")
	{
		// ask user about save file name
		filename.SetString(GetSaveFileName(encoding));
		m_doc->m_file_encoding = encoding;
		if (filename.IsEmpty())
			return CANCELLED;
	}

	if (m_doc->Save(filename))
	{
		m_file_age = FileAge(m_doc->m_filename);
		if (m_current_view == SOURCE)
			m_source.SendMessage(SCI_SETSAVEPOINT); //empty undo log in Scintilla
		//TO_DO empty undo in BODY!

		// set current directory
		wchar_t str[MAX_PATH];
		wcscpy(str, (const wchar_t *)m_doc->m_filename);
		PathRemoveFileSpec(str);
		SetCurrentDirectory(str);

		return OK;
	}
	else
	{
		return FAIL;
	}
}

///<summary>Load file into doc</summary>
///<params name="initfilename">file name</summary>
///<returns>FILE_OP_STATUS</summary>
CMainFrame::FILE_OP_STATUS CMainFrame::LoadFile(const wchar_t *initfilename)
{
	if (!DiscardChanges())
		return CANCELLED;

	CString filename(initfilename);
	// ask for filename, if was canceled by user, also return CANCELLED
	if (filename.IsEmpty())
		filename = GetOpenFileName();

	if (filename.IsEmpty())
		return CANCELLED;

	//EnableWindow(FALSE);
	FB::Doc *doc = new FB::Doc(*this);
	FB::Doc::m_active_doc = doc;

	if (!doc->Load(m_view, filename))
	{
		delete doc;
		FB::Doc::m_active_doc = m_doc;
		return FAIL;
	}
	if ((filename.ReverseFind(L'\\') + 1) != -1 && (filename.ReverseFind(L'\\') + 1) < filename.GetLength() - 1)
	{
		doc->m_body.m_file_path = filename.Mid(0, filename.ReverseFind(L'\\') + 1);
		doc->m_body.m_file_name = filename.Mid(filename.ReverseFind(L'\\') + 1, filename.GetLength() - 1);
	}
	//doc->m_body.m_file_name  = U::GetFileTitle(filename);
	m_file_age = FileAge(filename);

	delete m_doc;
	m_doc = doc;
	AttachDocument(m_doc);

	//ShowView(BODY);
	//  m_bad_xml = false;

	//Save IShellItem
	m_spShellItem.Release();
	PIDLIST_ABSOLUTE pidl = { 0 };
	HRESULT hr = SHParseDisplayName(m_doc->m_filename, NULL, &pidl, 0, NULL);
	if (SUCCEEDED(hr))
	{
		SHCreateShellItem(NULL, NULL, pidl, &m_spShellItem);
		ILFree(pidl);
	}

	return OK;
}

///<summary>Reload file if it was changed outside</summary>
///<returns>true if reloaded</summary>
bool CMainFrame::ReloadFile()
{
	FB::Doc *doc = new FB::Doc(*this);
	FB::Doc::m_active_doc = doc;

	//EnableWindow(FALSE);
	m_status.SetPaneText(ID_DEFAULT_PANE, _T("Loading..."));
	bool fLoaded = doc->Load(m_view, m_doc->m_filename);
	//EnableWindow(TRUE);
	if (!fLoaded)
	{
		delete doc;
		FB::Doc::m_active_doc = m_doc;
		return false;
	}
	m_file_age = FileAge(m_doc->m_filename);

	delete m_doc;
	m_doc = doc;
	AttachDocument(doc);
	return true;
}

///<summary>Get file last write time</summary>
///<params name="FileName">file name</summary>
///<returns>Last write time</summary>
unsigned __int64 CMainFrame::FileAge(LPCTSTR FileName)
{
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (::GetFileAttributesEx(FileName, GetFileExInfoStandard, &data))
	{
		return *((unsigned __int64 *)&data.ftLastWriteTime);
	}
	return unsigned __int64(~0);
}

///<summary>Check file timestamp
///If file was changed ask user and reload file</summary>
///<returns>true if successfully reload, false if not changed or user declined reload</summary>
bool CMainFrame::CheckFileTimeStamp()
{
	if (m_file_age == FileAge(m_doc->m_filename))
		return false;

	CString strMessage;
	strMessage.Format(IDS_FILE_CHANGED_MSG, (LPCTSTR)m_doc->m_filename);
	if (IDYES == AtlTaskDialog(*this, IDS_FILE_CHANGED_CPT, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, TD_WARNING_ICON))
		return ReloadFile();
	else
		m_file_age = FileAge(m_doc->m_filename); // reset timestamp

	return false;
}

void CMainFrame::GetDocumentStructure()
{
	m_document_tree.GetDocumentStructure(m_doc->m_body.Document());
}

BOOL CMainFrame::PreTranslateMessage(MSG *pMsg)
{
	// reset ctrl tab
	if (pMsg->message == WM_KEYUP && pMsg->wParam == VK_CONTROL)
	{
		m_ctrl_tab = false;
	}

	// well, if we are doing an incremental search, then swallow WM_CHARS
	if (m_incsearch && pMsg->hwnd != *this)
	{
		BOOL tmp;
		if (pMsg->message == WM_CHAR)
		{
			OnChar(WM_CHAR, pMsg->wParam, 0, tmp);
			return TRUE;
		}
		if ((pMsg->message == WM_KEYDOWN || pMsg->message == WM_KEYUP) &&
			(pMsg->wParam == VK_BACK || pMsg->wParam == VK_RETURN))
		{
			if (pMsg->message == WM_KEYDOWN)
				OnChar(WM_CHAR, pMsg->wParam, 0, tmp);
			return TRUE;
		}
	}

	// let other windows do their translations
	if (CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg))
		return TRUE;

	// this is needed to pass certain keys to the web browser
	HWND hWndFocus = ::GetFocus();
	if (m_doc)
	{
		if (::IsChild(m_doc->m_body, hWndFocus))
		{
			if (m_doc->m_body.PreTranslateMessage(pMsg))
				return TRUE;
			/*    } else if (::IsChild(m_doc->m_desc,hWndFocus)) {
			if (m_doc->m_desc.PreTranslateMessage(pMsg))
			return TRUE;*/
		}
	}

	return FALSE;
}

void CMainFrame::UIUpdateViewCmd(CFBEView &view, WORD wID, OLECMD &oc, const wchar_t *hk)
{
	CString fbuf;
	fbuf.Format(L"%s\t%s", (const TCHAR *)view.QueryCmdText(oc.cmdID), hk);
	UISetText(wID, fbuf);
	UIEnable(wID, (oc.cmdf & OLECMDF_ENABLED) != 0);
}

///<summary>Set enable state of command (toolbars& menu)</summary>
void CMainFrame::UIUpdateViewCmd(CFBEView &view, WORD wID)
{
	UIEnable(wID, view.CheckCommand(wID));
}

//Set checked state of command (toolbars& menu)
void CMainFrame::UISetCheckCmd(CFBEView &view, WORD wID)
{
	UISetCheck(wID, view.CheckSetCommand(wID));
}

// Functions for Initialization
void CMainFrame::AddTbButton(HWND hWnd, const TCHAR *text, const int idCommand, const BYTE bState, const HICON icon)
{
	CToolBarCtrl tb = hWnd;
	int iImage = I_IMAGENONE;
	BYTE bStyle = BTNS_BUTTON | BTNS_AUTOSIZE;
	if (icon)
	{
		CImageList iList = tb.GetImageList();
		if (iList)
			iImage = iList.AddIcon(icon);
	}

	tb.AddButton(idCommand, bStyle, bState, iImage, text, 0);
	// custom added command
	if (icon)
	{
		int idx = tb.CommandToIndex(idCommand);
		TBBUTTON tbButton;
		tb.GetButton(idx, &tbButton);
		AddToolbarButton(tb, tbButton, text);
		// move button to unassigned
		tb.DeleteButton(idx);
	}
	tb.AutoSize();
}

static void SubclassBox(HWND hWnd, RECT &rc, const int pos, CComboBox &box, DWORD dwStyle, CCustomEdit &custedit, const int resID, HFONT &hFont)
{
	::SendMessage(hWnd, TB_GETITEMRECT, pos, (LPARAM)&rc);
	rc.bottom--;
	box.Create(hWnd, rc, NULL, dwStyle, WS_EX_CLIENTEDGE, resID);
	box.SetFont(hFont);
	custedit.SubclassWindow(box.ChildWindowFromPoint(CPoint(3, 3)));
}

void CMainFrame::AddStaticText(CCustomStatic &st, HWND toolbarHwnd, int id, const TCHAR *text, HFONT hFont)
{
	RECT rect;
	SendMessage(toolbarHwnd, TB_GETITEMRECT, id, (LPARAM)&rect);
	rect.bottom--;

	st.Create(toolbarHwnd, rect, NULL, WS_CHILD | WS_VISIBLE, WS_EX_TRANSPARENT, IDC_ID);
	st.SetFont(hFont);
	st.SetWindowText(text);
	st.SetEnabled(true);
}

void CMainFrame::InitPluginsType(HMENU hMenu, const TCHAR *type, UINT cmdbase, CSimpleArray<CLSID> &plist)
{
	CRegKey rk;
	int ncmd = 0;

	if (rk.Open(HKEY_CURRENT_USER, _Settings.GetKeyPath() + L"\\Plugins") == ERROR_SUCCESS)
	{
		for (int i = 0; ncmd < 20; ++i)
		{
			CString name;
			DWORD size = 128; // enough for GUIDs
			TCHAR *cp = name.GetBuffer(size);
			FILETIME ft;
			if (::RegEnumKeyEx(rk, i, cp, &size, 0, 0, 0, &ft) != ERROR_SUCCESS)
				break;
			name.ReleaseBuffer(size);
			CRegKey pk;
			if (pk.Open(rk, name) != ERROR_SUCCESS)
				continue;
			CString pt(U::QuerySV(pk, L"Type"));
			CString ms(U::QuerySV(pk, L"Menu"));
			if (pt.IsEmpty() || ms.IsEmpty() || pt != type)
				continue;
			CLSID clsid;
			if (::CLSIDFromString((TCHAR *)(const TCHAR *)name, &clsid) != NOERROR)
				continue;

			// all checks pass, add to menu and remember clsid
			plist.Add(clsid);
			::AppendMenu(hMenu, MF_STRING, cmdbase + ncmd, ms);
			CString hs = ms;
			hs.Remove(L'&');
			InitPluginHotkey(name, cmdbase + ncmd, pt + CString(L" | ") + hs);
			// check if an icon is available
			CString icon(U::QuerySV(pk, L"Icon"));
			if (!icon.IsEmpty())
			{
				int cp1 = icon.ReverseFind(L',');
				int iconID;
				if (cp1 > 0 && _stscanf((const TCHAR *)icon + cp1, L",%d", &iconID) == 1)
					icon.Delete(cp1, icon.GetLength() - cp1);
				else
					iconID = 0;

				// try load from file first
				HICON hIcon;
				if (::ExtractIconEx(icon, iconID, NULL, &hIcon, 1) > 0 && hIcon)
				{
					m_MenuBar.AddIcon(hIcon, cmdbase + ncmd);
					::DestroyIcon(hIcon);
				}
			}
			++ncmd;
		}
	}

	// Old path to provide searching of old plugins
	CRegKey oldRk;
	if (oldRk.Open(HKEY_LOCAL_MACHINE, L"Software\\Haali\\FBE\\Plugins", KEY_READ) != ERROR_SUCCESS)
		goto skip;
	else
	{
		for (int i = ncmd; ncmd < 20; ++i)
		{
			CString name;
			DWORD size = 128; // enough for GUIDs
			TCHAR *cp = name.GetBuffer(size);
			FILETIME ft;
			if (::RegEnumKeyEx(oldRk, i, cp, &size, 0, 0, 0, &ft) != ERROR_SUCCESS)
				break;
			name.ReleaseBuffer(size);
			CRegKey pk;
			if (pk.Open(oldRk, name, KEY_READ) != ERROR_SUCCESS)
				continue;
			CString pt(U::QuerySV(pk, L"Type"));
			CString ms(U::QuerySV(pk, L"Menu"));
			if (pt.IsEmpty() || ms.IsEmpty() || pt != type)
				continue;
			CLSID clsid;
			if (::CLSIDFromString((TCHAR *)(const TCHAR *)name, &clsid) != NOERROR)
				continue;

			// all checks pass, add to menu and remember clsid
			plist.Add(clsid);
			::AppendMenu(hMenu, MF_STRING, cmdbase + ncmd, ms);
			CString hs = ms;
			hs.Remove(L'&');
			InitPluginHotkey(name, cmdbase + ncmd, pt + CString(L" | ") + hs);
			// check if an icon is available
			CString icon(U::QuerySV(pk, L"Icon"));
			if (!icon.IsEmpty())
			{
				int cp2 = icon.ReverseFind(L',');
				int iconID;
				if (cp2 > 0 && _stscanf((const TCHAR *)icon + cp2, L",%d", &iconID) == 1)
					icon.Delete(cp2, icon.GetLength() - cp2);
				else
					iconID = 0;

				// try load from file first
				HICON hIcon;
				if (::ExtractIconEx(icon, iconID, NULL, &hIcon, 1) > 0 && hIcon)
				{
					m_MenuBar.AddIcon(hIcon, cmdbase + ncmd);
					::DestroyIcon(hIcon);
				}
			}
			++ncmd;
		}
	}
skip:
	if (ncmd > 0) // delete placeholder from menu
		::RemoveMenu(hMenu, 0, MF_BYPOSITION);
}

void CMainFrame::InitPlugins()
{
	CollectScripts(_Settings.GetScriptsFolder(), L"*.js", 1, L"0");
	QuickScriptsSort(m_scripts, 0, m_scripts.GetSize() - 1);
	UpScriptsFolders(m_scripts);

	HMENU file = ::GetSubMenu(m_MenuBar.GetMenu(), 0);
	HMENU sub = ::GetSubMenu(file, 6);
	InitPluginsType(sub, L"Import", ID_IMPORT_BASE, m_import_plugins);

	sub = ::GetSubMenu(file, 7);
	InitPluginsType(sub, L"Export", ID_EXPORT_BASE, m_export_plugins);

	sub = ::GetSubMenu(file, 9);
	m_mru.SetMenuHandle(sub);
	m_mru.ReadFromRegistry(_Settings.GetKeyPath());
	m_mru.SetMaxEntries(m_mru.m_nMaxEntries_Max - 1);

	// Scripts
	HMENU ManMenu = m_MenuBar.GetMenu();
	HMENU scripts = GetSubMenu(ManMenu, 6);

	while (::GetMenuItemCount(scripts) > 0)
		::RemoveMenu(scripts, 0, MF_BYPOSITION);

	if (m_scripts.GetSize())
	{
		AddScriptsSubMenu(scripts, L"0", m_scripts);
	}
	else
	{
		wchar_t buf[MAX_LOAD_STRING + 1];
		::LoadString(_Module.GetResourceInstance(), IDS_NO_SCRIPTS, buf, MAX_LOAD_STRING);
		AppendMenu(scripts, MF_STRING | MF_DISABLED | MF_GRAYED, IDCANCEL, buf);
	}
}

// message handlers
// Display Overwrite\Insert status on Status panel
void CMainFrame::DisplayOvrInsStatus()
{
	// insert/overwrite mode
	CFBEView &view = ActiveView();
	OLECMD oc = { IDM_OVERWRITE };
	view.QueryStatus(&oc, 1);
	bool fOvr = (oc.cmdf & OLECMDF_LATCHED) != 0;

	if (fOvr != m_last_ie_ovr)
	{
		m_last_ie_ovr = fOvr;
		m_status.SetPaneText(ID_PANE_INS, fOvr ? strOVR : strINS);
	}
}

BOOL CMainFrame::OnIdle()
{
	if (CheckFileTimeStamp())
		return true;

	if (m_current_view == SOURCE)
	{
		// disable/enable paste menu items
		// check if there is active text selection
		bool fCanCC = m_source.SendMessage(SCI_GETSELECTIONSTART) != m_source.SendMessage(SCI_GETSELECTIONEND);

		UIEnable(ID_EDIT_COPY, fCanCC);
		UIEnable(ID_EDIT_CUT, fCanCC);

		UIEnable(ID_EDIT_PASTE, m_source.SendMessage(SCI_CANPASTE));
		UIEnable(ID_EDIT_UNDO, m_source.SendMessage(SCI_CANUNDO));
		UIEnable(ID_EDIT_REDO, m_source.SendMessage(SCI_CANREDO));

		// Overwrite/Insert mode
		m_last_sci_ovr = m_source.SendMessage(SCI_GETOVERTYPE);
		m_status.SetPaneText(ID_PANE_INS, m_last_sci_ovr ? strOVR : strINS);
		
		// enable\disable zoom
		int zoom = m_source.SendMessage(SCI_GETZOOM);
		UIEnable(ID_VIEW_ZOOMIN, zoom =20 ? FALSE : TRUE);
		UIEnable(ID_VIEW_ZOOMOUT, zoom = -10 ? FALSE : TRUE);

		// Display charcode on status panel
		DisplayCharCode();
	}

	if (m_current_view == DESC || m_current_view == BODY)
	{
		// enable\disable zoom
		SHD::IWebBrowser2Ptr brws = m_doc->m_body.Browser();
		CComVariant ret;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, NULL, &ret);
		int zoom = ret.intVal;
		UIEnable(ID_VIEW_ZOOMIN, zoom >= 1000 ? FALSE : TRUE);
		UIEnable(ID_VIEW_ZOOMOUT, zoom <= 25 ? FALSE : TRUE);
	}

	if (m_current_view == DESC)
	{
		// check if editing commands can be performed
		CString fbuf;
		CFBEView &view = ActiveView();

		static OLECMD mshtml_commands[] =
		{
		{ IDM_REDO },			// 0
		{ IDM_UNDO },			// 1
		{ IDM_COPY },			// 2
		{ IDM_CUT },			// 3
		{ IDM_PASTE }			// 4
		};
		view.QueryStatus(mshtml_commands, sizeof(mshtml_commands) / sizeof(mshtml_commands[0]));

		static WORD fbe_commands[] =
		{
			ID_EDIT_REDO,
			ID_EDIT_UNDO,
			ID_EDIT_COPY,
			ID_EDIT_CUT,
			ID_EDIT_PASTE
		};

		for (int jj = 0; jj < sizeof(mshtml_commands) / sizeof(mshtml_commands[0]); ++jj)
		{
			DWORD flags = mshtml_commands[jj].cmdf;
			WORD cmd = fbe_commands[jj];
			UIEnable(cmd, (flags & OLECMDF_ENABLED) != 0);
			UISetCheck(cmd, (flags & OLECMDF_LATCHED) != 0);
		}
		UIUpdateViewCmd(view, ID_EDIT_REDO, mshtml_commands[0], L"Ctrl+Y");
		UIUpdateViewCmd(view, ID_EDIT_UNDO, mshtml_commands[1], L"Ctrl+Z");

		DisplayOvrInsStatus();
	}
	// BODY view
	if (m_current_view == BODY)
	{
		// check if editing commands can be performed
		CString fbuf;
		CFBEView &view = ActiveView();

		static OLECMD mshtml_commands[] =
		{
			{IDM_REDO},			 // 0
			{IDM_UNDO},			 // 1
			{IDM_COPY},			 // 2
			{IDM_CUT},			 // 3
			{IDM_PASTE},		 // 4
			{IDM_UNLINK},		 // 5
			{IDM_BOLD},			 // 6
			{IDM_ITALIC},		 // 7
			{IDM_STRIKETHROUGH}, // 8
			{IDM_SUPERSCRIPT},   // 9
			{IDM_SUBSCRIPT},	 // 10
			{IDM_UNDERLINE},	 // 11
			{IDM_REMOVEFORMAT}	 // 12
		};
		view.QueryStatus(mshtml_commands, sizeof(mshtml_commands) / sizeof(mshtml_commands[0]));

		static WORD fbe_commands[] =
		{
			ID_EDIT_REDO,
			ID_EDIT_UNDO,
			ID_EDIT_COPY,
			ID_EDIT_CUT,
			ID_EDIT_PASTE,
			ID_STYLE_NOLINK,
			ID_STYLE_BOLD,
			ID_STYLE_ITALIC,
			ID_STYLE_STRIKETROUGH,
			ID_STYLE_SUPERSCRIPT,
			ID_STYLE_SUBSCRIPT,
			ID_STYLE_CODE,
			ID_STYLE_NORMAL
		};
		for (int jj = 0; jj < sizeof(mshtml_commands) / sizeof(mshtml_commands[0]); ++jj)
		{
			DWORD flags = mshtml_commands[jj].cmdf;
			WORD cmd = fbe_commands[jj];
			UIEnable(cmd, (flags & OLECMDF_ENABLED) != 0);
			UISetCheck(cmd, (flags & OLECMDF_LATCHED) != 0);
		}

		UIUpdateViewCmd(view, ID_EDIT_REDO, mshtml_commands[0], L"Ctrl+Y");
		UIUpdateViewCmd(view, ID_EDIT_UNDO, mshtml_commands[1], L"Ctrl+Z");

		UIEnable(ID_EDIT_FINDNEXT, view.CanFindNext());

		UIUpdateViewCmd(view, ID_STYLE_LINK);
		UIUpdateViewCmd(view, ID_STYLE_NOTE);
		//UIUpdateViewCmd(view, ID_STYLE_NORMAL);
		UIUpdateViewCmd(view, ID_ADD_SUBTITLE);
		UIUpdateViewCmd(view, ID_ADD_TEXTAUTHOR);
		UIUpdateViewCmd(view, ID_ADD_SECTION);
		UIUpdateViewCmd(view, ID_ADD_TITLE);
		UIUpdateViewCmd(view, ID_ADD_BODY);
		UIUpdateViewCmd(view, ID_ADD_BODY_NOTES);
		UIUpdateViewCmd(view, ID_ADD_ANN);
		UIUpdateViewCmd(view,ID_STYLE_CODE);
		//		UINT fEnable = view.CheckCommand(ID_ADD_BODY_NOTES) ? MF_ENABLED : MF_BYCOMMAND | MF_GRAYED;
//		::EnableMenuItem(m_MenuBar.GetMenu(), ID_ADD_BODY_NOTES, fEnable);
		
//		UIUpdateViewCmd(view, ID_ADD_TA);
		UIUpdateViewCmd(view, ID_EDIT_CLONE);
		UIUpdateViewCmd(view, ID_ADD_IMAGE);
		UIUpdateViewCmd(view, ID_ATTACH_BINARY);
		UIUpdateViewCmd(view, ID_ADD_EPIGRAPH);
		UIUpdateViewCmd(view, ID_EDIT_SPLIT);
		UIUpdateViewCmd(view, ID_ADD_POEM);
		UIUpdateViewCmd(view, ID_ADD_STANZA);
		UIUpdateViewCmd(view, ID_ADD_CITE);
		//UIUpdateViewCmd(view, ID_EDIT_CODE);
		//UISetCheckCmd(view, ID_EDIT_CODE);
		UIUpdateViewCmd(view, ID_INS_TABLE);
		UIUpdateViewCmd(view, ID_GOTO_FOOTNOTE);
		UIUpdateViewCmd(view, ID_GOTO_REFERENCE);
		UIUpdateViewCmd(view, ID_EDIT_MERGE);
		UIUpdateViewCmd(view, ID_EDIT_REMOVE_OUTER);

		// Added by SeNS: process bitmap paste
//		UIEnable(ID_EDIT_PASTE, m_source.SendMessage(SCI_CANPASTE) || BitmapInClipboard());

		if (m_sel_changed)
		{
			// update path in Statusbar
			m_status.SetPaneText(ID_DEFAULT_PANE, m_doc->m_body.SelPath());

			// update links and IDs
			try
			{
				MSHTML::IHTMLElementPtr an(m_doc->m_body.SelectionsStruct(L"anchor"));
				_variant_t href;

				if (an)
					href.Attach(an->getAttribute(L"href", 2).GetVARIANT());

				if ((bool)an && V_VT(&href) == VT_BSTR)
				{
					m_href_box.EnableWindow();
					m_href_caption.SetEnabled();
					m_ignore_cb_changes = true;

					if (!(m_href == ::GetFocus()))
					{
						// changed by SeNS: fix hrefs
						CString tmp(V_BSTR(&href));
						if (tmp.Find(L"file") == 0)
							tmp = tmp.Mid(tmp.ReverseFind(L'#'), 1024);
						m_href.SetWindowText(tmp);
						m_href.SetSel(tmp.GetLength(), tmp.GetLength(), FALSE);
					}

					m_ignore_cb_changes = false;

					// SeNS - images
					bool img = (U::scmp(an->tagName, L"DIV") == 0) || (U::scmp(an->tagName, L"SPAN") == 0);
					if (img != m_cb_last_images)
						m_cb_updated = false;
					m_cb_last_images = img;
				}
				else
				{
					m_href_box.SetWindowText(L"");
					m_href_box.EnableWindow(FALSE);
					m_href_caption.SetEnabled(false);
				}

				MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStruct(L""));
				if (sc)
				{
					m_id_box.EnableWindow();
					m_id_caption.SetEnabled(true);
					m_ignore_cb_changes = true;

					if (U::scmp(sc->id, L"fbw_body"))
						m_id.SetWindowText(sc->id);
					else
						m_id.SetWindowText(L"");

					m_ignore_cb_changes = false;
				}
				else
				{
					m_id_box.EnableWindow(FALSE);
					m_id_caption.SetEnabled(false);
				}

				MSHTML::IHTMLElementPtr im(m_doc->m_body.SelectionsStruct(L"image"));
				if (im)
				{
					m_image_title_box.EnableWindow();
					m_image_title_caption.SetEnabled();
					m_ignore_cb_changes = true;
					m_image_title.SetWindowText(im->title);
					_bstr_t title = im->title;
					if (title.length())
						m_image_title.SetSel(wcslen(im->title), wcslen(im->title), FALSE);
					m_ignore_cb_changes = false;
				}
				else
				{
					m_image_title_box.SetWindowText(L"");
					m_image_title_box.EnableWindow(FALSE);
					m_image_title_caption.SetEnabled(false);
				}

				// отображение ID для тегов <section>
				MSHTML::IHTMLElementPtr scstn(m_doc->m_body.SelectionsStruct(L"section"));
				if (scstn)
				{
					m_section_box.EnableWindow(TRUE);
					m_section_id_caption.SetEnabled();
					m_ignore_cb_changes = true;
					m_section.SetWindowText(scstn->id);
					m_ignore_cb_changes = false;
				}
				else
				{
					m_section_box.SetWindowText(L"");
					m_section_box.EnableWindow(FALSE);
					m_section_id_caption.SetEnabled(false);
				}
				// отображение ID для тегов <table>
				MSHTML::IHTMLElementPtr sct(m_doc->m_body.SelectionsStruct(L"table"));
				if (sct)
				{
					m_id_table_id_box.EnableWindow(TRUE);
					m_table_id_caption.SetEnabled();
					m_ignore_cb_changes = true;
					m_id_table_id.SetWindowText(sct->id);
					m_ignore_cb_changes = false;
				}
				else
				{
					m_id_table_id_box.SetWindowText(L"");
					m_id_table_id_box.EnableWindow(FALSE);
					m_table_id_caption.SetEnabled(false);
				}

				// отображение ID для тегов <tr>, <th>, <td>
				MSHTML::IHTMLElementPtr sctc(m_doc->m_body.SelectionsStruct(L"th"));
				if (sctc)
				{
					m_id_table_box.EnableWindow(TRUE);
					m_id_table_caption.SetEnabled();
					m_ignore_cb_changes = true;
					m_id_table.SetWindowText(sctc->id);
					m_ignore_cb_changes = false;
				}
				else
				{
					m_id_table_box.SetWindowText(L"");
					m_id_table_box.EnableWindow(FALSE);
					m_id_table_caption.SetEnabled(false);
				}

				// отображение style для тегов <table>
				_bstr_t styleT("");
				MSHTML::IHTMLElementPtr scsT(m_doc->m_body.SelectionsStructB(L"table", styleT));
				if (scsT)
				{
					m_styleT_table_box.EnableWindow(TRUE);
					m_table_style_caption.SetEnabled();
					if (U::scmp(styleT, L"") != 0)
					{
						m_styleT_table_box.EnableWindow(TRUE);
						m_table_style_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_styleT_table.SetWindowText(styleT);
						m_ignore_cb_changes = false;
					}
					else
					{
						m_styleT_table_box.SetWindowText(L"");
					}
				}
				else
				{
					m_styleT_table_box.SetWindowText(L"");
					m_table_style_caption.SetEnabled(false);
					m_styleT_table_box.EnableWindow(FALSE);
				}

				// отображение style для тегов <th>, <td>
				_bstr_t style("");
				MSHTML::IHTMLElementPtr scs(m_doc->m_body.SelectionsStructB(L"th", style));
				if (scs)
				{
					m_style_table_box.EnableWindow(TRUE);
					m_style_caption.SetEnabled();
					if (U::scmp(style, L"") != 0)
					{
						m_style_table_box.EnableWindow(TRUE);
						m_style_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_style_table.SetWindowText(style);
						m_ignore_cb_changes = false;
					}
					else
					{
						m_style_table_box.SetWindowText(L"");
					}
				}
				else
				{
					m_style_table_box.SetWindowText(L"");
					m_style_table_box.EnableWindow(FALSE);
					m_style_caption.SetEnabled(false);
				}

				// отображение colspan для тегов <th>, <td>
				_bstr_t colspan("");
				MSHTML::IHTMLElementPtr scc(m_doc->m_body.SelectionsStructB(L"colspan", colspan));
				if (scc)
				{
					m_colspan_table_box.EnableWindow(TRUE);
					m_colspan_caption.SetEnabled();
					if (U::scmp(colspan, L"") != 0)
					{
						m_colspan_table_box.EnableWindow(TRUE);
						m_colspan_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_colspan_table.SetWindowText(colspan);
						m_ignore_cb_changes = false;
					}
					else
					{
						m_colspan_table_box.SetWindowText(L"");
					}
				}
				else
				{
					m_colspan_table_box.SetWindowText(_T(""));
					m_colspan_table_box.EnableWindow(FALSE);
					m_colspan_caption.SetEnabled(false);
				}

				// отображение rowspan для тегов <th>, <td>
				_bstr_t rowspan("");
				MSHTML::IHTMLElementPtr scr(m_doc->m_body.SelectionsStructB(L"rowspan", rowspan));
				if (scr)
				{
					m_rowspan_table_box.EnableWindow(TRUE);
					m_rowspan_caption.SetEnabled();
					if (U::scmp(rowspan, L"") != 0)
					{
						m_rowspan_table_box.EnableWindow(TRUE);
						m_rowspan_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_rowspan_table.SetWindowText(rowspan);
						m_ignore_cb_changes = false;
					}
					else
					{
						m_rowspan_table_box.SetWindowText(L"");
					}
				}
				else
				{
					m_rowspan_table_box.SetWindowText(L"");
					m_rowspan_table_box.EnableWindow(FALSE);
					m_rowspan_caption.SetEnabled(false);
				}

				// отображение align для тегов <tr>
				_bstr_t alignTR("");
				MSHTML::IHTMLElementPtr scaTR(m_doc->m_body.SelectionsStructB(L"tralign", alignTR));
				if (scaTR)
				{
					m_alignTR_table_box.EnableWindow(TRUE);
					m_tr_allign_caption.SetEnabled();
					if (U::scmp(alignTR, L"") != 0)
					{
						m_alignTR_table_box.EnableWindow(TRUE);
						m_tr_allign_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_alignTR_table_box.SetCurSel(m_alignTR_table_box.FindString(0, alignTR));
						m_ignore_cb_changes = false;
					}
					else
					{
						m_alignTR_table_box.SetCurSel(m_alignTR_table_box.FindString(0, L""));
					}
				}
				else
				{
					m_alignTR_table_box.SetCurSel(m_alignTR_table_box.FindString(0, L""));
					m_alignTR_table_box.EnableWindow(FALSE);
					m_tr_allign_caption.SetEnabled(false);
				}

				// отображение align для тегов <th>, <td>
				_bstr_t align("");
				MSHTML::IHTMLElementPtr sca(m_doc->m_body.SelectionsStructB(L"align", align));
				if (sca)
				{
					m_align_table_box.EnableWindow(TRUE);
					m_th_allign_caption.SetEnabled();
					if (U::scmp(align, L"") != 0)
					{
						m_align_table_box.EnableWindow(TRUE);
						m_th_allign_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_align_table_box.SetCurSel(m_align_table_box.FindString(0, align));
						m_ignore_cb_changes = false;
					}
					else
					{
						m_align_table_box.SetCurSel(m_align_table_box.FindString(0, L""));
					}
				}
				else
				{
					m_align_table_box.SetCurSel(m_align_table_box.FindString(0, L""));
					m_align_table_box.EnableWindow(FALSE);
					m_th_allign_caption.SetEnabled(false);
				}

				// отображение valign для тегов <th>, <td>
				_bstr_t valign("");
				MSHTML::IHTMLElementPtr scva(m_doc->m_body.SelectionsStructB(L"valign", valign));
				if (scva)
				{
					m_valign_table_box.EnableWindow(TRUE);
					m_valign_caption.SetEnabled();
					if (U::scmp(valign, L"") != 0)
					{
						m_valign_table_box.EnableWindow(TRUE);
						m_valign_caption.SetEnabled();
						m_ignore_cb_changes = true;
						m_valign_table_box.SetCurSel(m_valign_table_box.FindString(0, valign));
						m_ignore_cb_changes = false;
					}
					else
					{
						m_valign_table_box.SetCurSel(m_valign_table_box.FindString(0, L""));
					}
				}
				else
				{
					m_valign_table_box.SetCurSel(m_valign_table_box.FindString(0, L""));
					m_valign_table_box.EnableWindow(FALSE);
					m_valign_caption.SetEnabled(false);
				}
			}
			catch (_com_error &)
			{
			}

			// update current tree node
			if (!m_doc->IsChanged() && _Settings.ViewDocumentTree())
				m_document_tree.HighlightItemAtPos(m_doc->m_body.SelectionContainer()); // locate appropriate tree node

			m_sel_changed = false;
		}

		// insert/overwrite mode
		DisplayOvrInsStatus();

		// added by SeNS
		// detect page scrolling, run a background spellcheck if necessary
		// hide\show spellcheck toolbar button
		if (!m_Speller->Enabled())
			UIEnable(ID_TOOLS_SPELLCHECK, false, true);
		else
		{
			UIEnable(ID_TOOLS_SPELLCHECK, true, true);
			m_Speller->CheckScroll();
		}
	}

	// update UI
	UIUpdateToolBar();
	UIUpdateMenuBar();

	// focus some stupid control if requested
	BOOL tmp;
	switch (m_want_focus)
	{
	case IDC_ID:
		OnSelectCtl(0, ID_SELECT_ID, 0, tmp);
		break;
	case IDC_HREF:
		OnSelectCtl(0, ID_SELECT_HREF, 0, tmp);
		break;
	case IDC_IMAGE_TITLE:
		OnSelectCtl(0, ID_SELECT_IMAGE, 0, tmp);
		break;
	case IDC_SECTION:
		OnSelectCtl(0, ID_SELECT_SECTION, 0, tmp);
		break;
	case IDC_IDT:
		OnSelectCtl(0, ID_SELECT_IDT, 0, tmp);
		break;
	case IDC_STYLET:
		OnSelectCtl(0, ID_SELECT_STYLET, 0, tmp);
		break;
	case IDC_STYLE:
		OnSelectCtl(0, ID_SELECT_STYLE, 0, tmp);
		break;
	case IDC_COLSPAN:
		OnSelectCtl(0, ID_SELECT_COLSPAN, 0, tmp);
		break;
	case IDC_ROWSPAN:
		OnSelectCtl(0, ID_SELECT_ROWSPAN, 0, tmp);
		break;
	case IDC_ALIGNTR:
		OnSelectCtl(0, ID_SELECT_ALIGNTR, 0, tmp);
		break;
	case IDC_ALIGN:
		OnSelectCtl(0, ID_SELECT_ALIGN, 0, tmp);
		break;
	case IDC_VALIGN:
		OnSelectCtl(0, ID_SELECT_VALIGN, 0, tmp);
		break;
	default:
		break;
	}
	m_want_focus = 0;

	// install a posted status line message
	if (!m_status_msg.IsEmpty())
	{
		m_status.SetPaneText(ID_DEFAULT_PANE, m_status_msg);
		m_status_msg.Empty();
	}

	m_doc->Changed(m_doc->IsChanged() || m_doc->m_body.IsDescFormChanged() || (m_current_view == SOURCE && m_source.SendMessage(SCI_GETMODIFY)));

	// see if we need to update title
	if (m_need_title_update || m_change_state != m_doc->IsChanged())
	{
		m_need_title_update = false;
		//m_change_state = m_doc->IsChanged();
		CString tt(U::GetFileTitle(m_doc->m_filename));
		SetWindowText((m_change_state ? L"*" : L"") + tt + L"-FB Editor");
	}

	return FALSE;
}

/// <summary>WM Handler(WM_CREATE)</summary>
/// <returns>S_OK</returns>
LRESULT CMainFrame::OnCreate(UINT, WPARAM, LPARAM, BOOL &)
{
	m_ctrl_tab = false;

	// create command bar window
	m_MenuBar.SetAlphaImages(true);
	HWND hWndCmdBar = m_MenuBar.Create(m_hWnd, rcDefault, NULL, ATL_SIMPLE_CMDBAR_PANE_STYLE);
	// attach menu
	m_MenuBar.AttachMenu(GetMenu());
	// remove old menu
	SetMenu(NULL);
	// load command bar images
	m_MenuBar.LoadImages(IDR_MAINFRAME_SMALL);

	m_CmdToolbar = CreateSimpleToolBarCtrl(m_hWnd, IDR_MAINFRAME, FALSE, ATL_SIMPLE_TOOLBAR_PANE_STYLE | TBSTYLE_LIST | CCS_ADJUSTABLE);
	m_CmdToolbar.SetExtendedStyle(TBSTYLE_EX_MIXEDBUTTONS);
	InitToolBar(m_CmdToolbar, IDR_MAINFRAME);
	// Restore commands toolbar layout and position
	//m_CmdToolbar.RestoreState(HKEY_CURRENT_USER, L"SOFTWARE\\FBETeam\\FictionBook Editor\\Toolbars", L"CommandToolbar");
	UIAddToolBar(m_CmdToolbar);

	m_ScriptsToolbar = CreateSimpleToolBarCtrl(m_hWnd, IDR_SCRIPTS, FALSE, ATL_SIMPLE_TOOLBAR_PANE_STYLE | TBSTYLE_LIST | CCS_ADJUSTABLE);
	m_ScriptsToolbar.SetExtendedStyle(TBSTYLE_EX_MIXEDBUTTONS);
	InitToolBar(m_ScriptsToolbar, IDR_SCRIPTS);
	UIAddToolBar(m_ScriptsToolbar);

	HWND hWndLinksBar = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, ATL_SIMPLE_TOOLBAR_PANE_STYLE | TBSTYLE_LIST, 0, 0, 100, 100,
		m_hWnd, NULL, _Module.GetModuleInstance(), NULL);

	HWND hWndTableBar = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, ATL_SIMPLE_TOOLBAR_PANE_STYLE | TBSTYLE_LIST, 0, 0, 100, 100,
		m_hWnd, NULL, _Module.GetModuleInstance(), NULL);
	HWND hWndTableBar2 = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, ATL_SIMPLE_TOOLBAR_PANE_STYLE | TBSTYLE_LIST, 0, 0, 100, 100,
		m_hWnd, NULL, _Module.GetModuleInstance(), NULL);

	wchar_t buf[MAX_LOAD_STRING + 1];
	HFONT hFont = (HFONT)::SendMessage(hWndLinksBar, WM_GETFONT, 0, 0);

	// Links toolbar preparation
	::SendMessage(hWndLinksBar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
	// Next line provides empty drawing of text
	::SendMessage(hWndLinksBar, TB_SETDRAWTEXTFLAGS, (WPARAM)DT_CALCRECT, (LPARAM)DT_CALCRECT);

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_ID, buf, MAX_LOAD_STRING);
	AddTbButton(hWndLinksBar, buf);
	AddStaticText(m_id_caption, hWndLinksBar, 0, buf, hFont);
	AddTbButton(hWndLinksBar, L"123456789012345678901234567890");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_HREF, buf, MAX_LOAD_STRING);
	AddTbButton(hWndLinksBar, buf);
	AddStaticText(m_href_caption, hWndLinksBar, 2, buf, hFont);
	AddTbButton(hWndLinksBar, L"123456789012345678901234567890");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_SECTION_ID, buf, MAX_LOAD_STRING);
	AddTbButton(hWndLinksBar, buf);
	AddStaticText(m_section_id_caption, hWndLinksBar, 4, buf, hFont);
	AddTbButton(hWndLinksBar, L"123456789012345678901234567890");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_IMAGE_TITLE, buf, MAX_LOAD_STRING);
	AddTbButton(hWndLinksBar, buf);
	AddStaticText(m_image_title_caption, hWndLinksBar, 6, buf, hFont);
	AddTbButton(hWndLinksBar, L"123456789012345678901234567890");

	// Table's first toolbar preparation
	::SendMessage(hWndTableBar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
	::SendMessage(hWndTableBar, TB_SETDRAWTEXTFLAGS, (WPARAM)DT_CALCRECT, (LPARAM)DT_CALCRECT);

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_TABLE_ID, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar, buf);
	AddStaticText(m_table_id_caption, hWndTableBar, 0, buf, hFont);
	AddTbButton(hWndTableBar, L"12345678901234567890");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_TABLE_STYLE, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar, buf);
	AddStaticText(m_table_style_caption, hWndTableBar, 2, buf, hFont);
	AddTbButton(hWndTableBar, L"123456789012345");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_ID, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar, buf);
	AddStaticText(m_id_table_caption, hWndTableBar, 4, buf, hFont);
	AddTbButton(hWndTableBar, L"12345678901234567890");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_STYLE, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar, buf);
	AddStaticText(m_style_caption, hWndTableBar, 6, buf, hFont);
	AddTbButton(hWndTableBar, L"123456789012345");

	// Table's second toolbar preparation
	::SendMessage(hWndTableBar2, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
	::SendMessage(hWndTableBar2, TB_SETDRAWTEXTFLAGS, (WPARAM)DT_CALCRECT, (LPARAM)DT_CALCRECT);

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_COLSPAN, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar2, buf);
	AddStaticText(m_colspan_caption, hWndTableBar2, 0, buf, hFont);
	AddTbButton(hWndTableBar2, L"12345");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_ROWSPAN, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar2, buf);
	AddStaticText(m_rowspan_caption, hWndTableBar2, 2, buf, hFont);
	AddTbButton(hWndTableBar2, L"12345");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_TR_ALIGN, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar2, buf);
	AddStaticText(m_tr_allign_caption, hWndTableBar2, 4, buf, hFont);
	AddTbButton(hWndTableBar2, L"12345678");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_TD_ALIGN, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar2, buf);
	AddStaticText(m_th_allign_caption, hWndTableBar2, 6, buf, hFont);
	AddTbButton(hWndTableBar2, L"12345678");

	::LoadString(_Module.GetResourceInstance(), IDS_TB_CAPT_TD_VALIGN, buf, MAX_LOAD_STRING);
	AddTbButton(hWndTableBar2, buf);
	AddStaticText(m_valign_caption, hWndTableBar2, 8, buf, hFont);
	AddTbButton(hWndTableBar2, L"12345678");

	CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);

	AddSimpleReBarBand(hWndCmdBar, 0, TRUE, 0);
	AddSimpleReBarBand(m_CmdToolbar, 0, TRUE, 0, FALSE);
	AddSimpleReBarBand(m_ScriptsToolbar, 0, TRUE, 0, FALSE);
	AddSimpleReBarBand(hWndLinksBar, 0, TRUE, 0, TRUE);
	AddSimpleReBarBand(hWndTableBar, 0, TRUE, 0, TRUE);
	AddSimpleReBarBand(hWndTableBar2, 0, TRUE, 0, TRUE);
	m_rebar = m_hWndToolBar;

	// add editor controls
	RECT rc;

	HDC hdc1 = ::GetDC(m_id_caption);
	SetBkColor(hdc1, RGB(0, 0, 0));
	//ReleaseDC(hdc);
	ReleaseDC(hdc1);

	DWORD CBS_COMMON_STYLE = WS_CHILD | WS_VISIBLE | CBS_AUTOHSCROLL;

	SubclassBox(hWndLinksBar, rc, 1, m_id_box, CBS_COMMON_STYLE, m_id, IDC_ID, hFont);
	SubclassBox(hWndLinksBar, rc, 3, m_href_box, CBS_COMMON_STYLE | WS_VSCROLL | CBS_DROPDOWN | CBS_SORT, m_href, IDC_HREF, hFont);
	SubclassBox(hWndLinksBar, rc, 5, m_section_box, CBS_COMMON_STYLE, m_section, IDC_SECTION, hFont);
	SubclassBox(hWndLinksBar, rc, 7, m_image_title_box, CBS_COMMON_STYLE, m_image_title, IDC_IMAGE_TITLE, hFont);

	// add editor-table controls
	HFONT hFontT = (HFONT)::SendMessage(hWndTableBar, WM_GETFONT, 0, 0);
	RECT rcT;

	SubclassBox(hWndTableBar, rcT, 1, m_id_table_id_box, CBS_COMMON_STYLE, m_id_table_id, IDC_IDT, hFontT);
	SubclassBox(hWndTableBar, rcT, 3, m_styleT_table_box, CBS_COMMON_STYLE, m_styleT_table, IDC_STYLET, hFontT);
	SubclassBox(hWndTableBar, rcT, 5, m_id_table_box, CBS_COMMON_STYLE, m_id_table, IDC_ID, hFontT);
	SubclassBox(hWndTableBar, rcT, 7, m_style_table_box, CBS_COMMON_STYLE, m_style_table, IDC_STYLE, hFontT);

	SubclassBox(hWndTableBar2, rcT, 1, m_colspan_table_box, CBS_COMMON_STYLE, m_colspan_table, IDC_COLSPAN, hFontT);
	SubclassBox(hWndTableBar2, rcT, 3, m_rowspan_table_box, CBS_COMMON_STYLE, m_rowspan_table, IDC_ROWSPAN, hFontT);
	SubclassBox(hWndTableBar2, rcT, 5, m_alignTR_table_box, CBS_COMMON_STYLE | WS_VSCROLL | CBS_DROPDOWNLIST, m_alignTR_table, IDC_ALIGNTR, hFontT);
	SubclassBox(hWndTableBar2, rcT, 7, m_align_table_box, CBS_COMMON_STYLE | WS_VSCROLL | CBS_DROPDOWNLIST, m_align_table, IDC_ALIGN, hFontT);
	SubclassBox(hWndTableBar2, rcT, 9, m_valign_table_box, CBS_COMMON_STYLE | WS_VSCROLL | CBS_DROPDOWNLIST, m_valign_table, IDC_VALIGN, hFontT);

	m_align_table_box.InsertString(0, _T(""));
	m_align_table_box.InsertString(1, _T("left"));
	m_align_table_box.InsertString(2, _T("right"));
	m_align_table_box.InsertString(3, _T("center"));

	m_alignTR_table_box.InsertString(0, _T(""));
	m_alignTR_table_box.InsertString(1, _T("left"));
	m_alignTR_table_box.InsertString(2, _T("right"));
	m_alignTR_table_box.InsertString(3, _T("center"));

	m_valign_table_box.InsertString(0, _T(""));
	m_valign_table_box.InsertString(1, _T("top"));
	m_valign_table_box.InsertString(2, _T("middle"));
	m_valign_table_box.InsertString(3, _T("bottom"));

	// create status bar
	CreateSimpleStatusBar();
	m_status.SubclassWindow(m_hWndStatusBar);
	int panes[] =
	{
		ID_DEFAULT_PANE,
		ID_PANE_CHAR,
		399,
		ID_PANE_INS };
	m_status.SetPanes(panes, sizeof(panes) / sizeof(panes[0]));
	m_status.SetPaneWidth(ID_PANE_CHAR, 60);
	m_status.SetPaneText(ID_PANE_CHAR, L"");
	m_status.SetPaneWidth(399, 20);
	m_status.SetPaneWidth(ID_PANE_INS, 30);

	// load insert/overwrite abbreviations
	::LoadString(_Module.GetResourceInstance(), IDS_PANE_INS, strINS, MAX_LOAD_STRING);
	::LoadString(_Module.GetResourceInstance(), IDS_PANE_OVR, strOVR, MAX_LOAD_STRING);

	// create splitter
	m_hWndClient = m_splitter.Create(m_hWnd, rcDefault, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
	m_splitter.SetSplitterExtendedStyle(0);

	// create splitter contents
	m_view.Create(m_splitter, rcDefault, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN);

	// create a tree
	/*m_dummy_pane.Create(m_document_tree,rcDefault,NULL,WS_CHILD|WS_VISIBLE|WS_CLIPSIBLINGS|WS_CLIPCHILDREN,WS_EX_CLIENTEDGE);
	m_document_tree.SetClient(m_dummy_pane);
	m_document_tree.Create(m_dummy_pane, rcDefault);
	m_document_tree.SetBkColor(::GetSysColor(COLOR_WINDOW));
	m_dummy_pane.SetSplitterPane(0,m_document_tree);
	m_dummy_pane.SetSinglePaneMode(SPLIT_PANE_LEFT);*/

	// create a source view (Scintilla editor)
	m_source.Create(_T("Scintilla"), m_view, rcDefault, NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0);
	m_view.AttachWnd(m_source);
	SetupSci();
	SetSciStyles();

	// initialize a new blank document
	m_doc = new FB::Doc(*this);
	FB::Doc::m_active_doc = m_doc;

	// load a command line arg if it was provided
	bool start_with_params = false;
	if (_ARGV.GetSize() > 0 && !_ARGV[0].IsEmpty())
	{
		if (m_doc->Load(m_view, _ARGV[0]))
		{
			start_with_params = true;
			m_file_age = FileAge(_ARGV[0]);
		}
		else
		{
			delete m_doc;
			m_doc = new FB::Doc(*this);
			FB::Doc::m_active_doc = m_doc;
			m_doc->CreateBlank(m_view);
			m_file_age = unsigned __int64(~0);
		}
	}
	else
	{
		m_doc->CreateBlank(m_view);
		m_file_age = unsigned __int64(~0);
	}
	UISetCheck(ID_VIEW_BODY, 1);

	m_document_tree.Create(m_splitter);

	if (U::_ARGS.start_in_desc_mode)
		ShowView(DESC);

	// init plugins&MRU list
	InitPlugins();

	// setup splitter
	m_splitter.SetSplitterPanes(m_document_tree, m_view);

	// hide elements
	if (_Settings.ViewStatusBar())
	{
		UISetCheck(ID_VIEW_STATUS_BAR, 1);
	}
	else
	{
		m_status.ShowWindow(SW_HIDE);
		UISetCheck(ID_VIEW_STATUS_BAR, FALSE);
	}

	if (_Settings.ViewDocumentTree())
	{
		UISetCheck(ID_VIEW_TREE, 1);
	}
	else
	{
		m_document_tree.ShowWindow(SW_HIDE);
		UISetCheck(ID_VIEW_TREE, FALSE);
		m_splitter.SetSinglePaneMode(SPLIT_PANE_RIGHT);
	}

	// load toolbar settings
	for (int j = ATL_IDW_BAND_FIRST; j < ATL_IDW_BAND_FIRST + 5; ++j)
		UISetCheck(j, TRUE);
	REBARBANDINFO rbi;
	memset(&rbi, 0, sizeof(rbi));
	rbi.cbSize = sizeof(rbi);
	rbi.fMask = RBBIM_SIZE | RBBIM_STYLE;
	CString tbs(_Settings.GetToolbarsSettings());
	const TCHAR *cp = tbs;
	for (int bn = 0;; ++bn)
	{
		const TCHAR *ce = _tcschr(cp, _T(';'));
		if (!ce)
			break;
		int id, style, cx;
		if (_stscanf(cp, _T("%d,%d,%d;"), &id, &style, &cx) != 3)
			break;
		cp = ce + 1;
		int idx = m_rebar.IdToIndex(id);
		m_rebar.GetBandInfo(idx, &rbi);
		rbi.fStyle &= ~(RBBS_BREAK | RBBS_HIDDEN);
		style &= RBBS_BREAK | RBBS_HIDDEN;
		rbi.fStyle |= style;
		rbi.cx = cx;
		m_rebar.SetBandInfo(idx, &rbi);
		if (idx != bn)
			m_rebar.MoveBand(idx, bn);
		UISetCheck(id, style & RBBS_HIDDEN ? FALSE : TRUE);
	}

	// delay resizing
	PostMessage(U::WM_POSTCREATE);

	// register object for message filtering and idle updates
	CMessageLoop *pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	// accept dropped files
	::DragAcceptFiles(*this, TRUE);

	// Modification by Pilgrim
	BOOL bVisible = _Settings.ViewDocumentTree();
	m_document_tree.ShowWindow(bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_TREE, bVisible);
	m_splitter.SetSinglePaneMode(bVisible ? SPLIT_PANE_NONE : SPLIT_PANE_RIGHT);

	if (start_with_params)
	{
		m_mru.AddToList(_ARGV[0]);
		if (_Settings.m_restore_file_position)
		{
			m_restore_pos_cmdline = true;
		}
	}

	// Change keyboard layout
	if (_Settings.m_change_kbd_layout_check)
	{
		CString layout;
		layout.Format(L"%08x", _Settings.GetKeybLayout());
		LoadKeyboardLayout(layout, KLF_ACTIVATE);
	}

	// added by SeNS: create blank document, and load incorrect XML to Scintilla
	/*  if (m_bad_xml)
	  if (!LoadToScintilla(_ARGV[0])) return -1;*/

	AttachDocument(m_doc);

	// Added by SeNS
	if (_Settings.m_usespell_check)
	{
		UIEnable(ID_TOOLS_SPELLCHECK, true, true);
		CString custDictName = _Settings.m_custom_dict;
		if (!custDictName.IsEmpty())
			m_Speller->SetCustomDictionary(U::GetProgDir() + custDictName);

		m_Speller->Enable(true);
	}
	else
		UIEnable(ID_TOOLS_SPELLCHECK, false, true);

	// Restore scripts toolbar layout and position
	m_ScriptsToolbar.RestoreState(HKEY_CURRENT_USER, L"SOFTWARE\\FBETeam\\FictionBook Editor\\Toolbars", L"ScriptsToolbar");

	return S_OK;
}

/// <summary>WM Handler(WM_DESTROY)</summary>
/// <returns>S_OK</returns>
LRESULT CMainFrame::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	DestroyAcceleratorTable(m_hAccel);
	bHandled = FALSE;
	return S_OK;
}
/// <summary>WM Handler(WM_CLOSE)
/// If discard accepted save settings</summary>
/// <returns>S_OK if can't close</returns>
LRESULT CMainFrame::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL & /*bHandled*/)
{
	if (DiscardChanges())
	{
		m_Speller->EndDocumentCheck();
		m_Speller->Enable(false);

		// Save settings
		_Settings.SetViewStatusBar(m_status.IsWindowVisible() != 0);
		_Settings.SetSplitterPos(m_splitter.GetSplitterPos());
		//_Settings.SetViewDocumentTree(IsSourceActive() ? m_document_tree.IsWindowVisible()==0 : !m_save_sp_mode);

		// save window position
		WINDOWPLACEMENT wpl;
		wpl.length = sizeof(wpl);
		GetWindowPlacement(&wpl);
		_Settings.SetWindowPosition(wpl);

		// save MRU list
		m_mru.WriteToRegistry(_Settings.GetKeyPath());

		// save toolbars state
		m_CmdToolbar.SaveState(HKEY_CURRENT_USER, L"SOFTWARE\\FBETeam\\FictionBook Editor\\Toolbars", L"CommandToolbar");
		m_ScriptsToolbar.SaveState(HKEY_CURRENT_USER, L"SOFTWARE\\FBETeam\\FictionBook Editor\\Toolbars", L"ScriptsToolbar");

		CString tbs;
		REBARBANDINFO rbi;
		memset(&rbi, 0, sizeof(rbi));
		rbi.cbSize = sizeof(rbi);
		rbi.fMask = RBBIM_ID | RBBIM_SIZE | RBBIM_STYLE;
		int num_bands = m_rebar.GetBandCount();
		for (int i = 0; i < num_bands; ++i)
		{
			m_rebar.GetBandInfo(i, &rbi);
			CString bi;
			bi.Format(_T("%d,%d,%d;"), rbi.wID, rbi.fStyle, rbi.cx);
			tbs += bi;
		}
		_Settings.SetToolbarsSettings(tbs);

		// save hotkey
		_Settings.SaveHotkeyGroups();

		_Settings.Save();
		_Settings.SaveWords();
		_Settings.Close();

		DefWindowProc(WM_CLOSE, 0, 0);
		return 1;
	}
	return S_OK;
}
/// <summary>WM Handler(WM_POSTCREATE) Some initialization after create
/// set Splitter and menu accelerators</summary>
/// <returns>0</returns>
LRESULT CMainFrame::OnPostCreate(UINT, WPARAM, LPARAM, BOOL &)
{
	//SetSplitterPos works best after the default WM_CREATE has been handled
	m_splitter.SetSplitterPos(_Settings.GetSplitterPos());

	// Init Menu Accelerators
	_Settings.LoadHotkeyGroups();
	DestroyAcceleratorTable(m_hAccel);

	LPACCEL lpaccelNew = new ACCEL[_Settings.keycodes];
	int HKentries = _Settings.keycodes;
	for (unsigned int i = 0; i < _Settings.m_hotkey_groups.size(); ++i)
	{
		for (unsigned int j = 0; j < _Settings.m_hotkey_groups.at(i).m_hotkeys.size(); ++j)
		{
			ACCEL accel = _Settings.m_hotkey_groups.at(i).m_hotkeys.at(j).m_accel;
			if (accel.fVirt != NULL && accel.key != NULL && accel.cmd != NULL)
				lpaccelNew[--HKentries] = accel;
		}
	}

	m_hAccel = CreateAcceleratorTable(lpaccelNew, _Settings.keycodes);
	delete[] lpaccelNew;

	FillMenuWithHkeys(m_MenuBar.GetMenu());
	return S_OK;
}

/// <summary>WM Handler(WM_SIZE) Update size</summary>
/// <returns>0</returns>
LRESULT CMainFrame::OnSize(UINT, WPARAM, LPARAM, BOOL &bHandled)
{
	UpdateViewSizeInfo();
	bHandled = FALSE;
	return S_OK;
}

/// <summary>WM Handler(WM_SETTINGCHANGED) Reapply settings</summary>
/// <returns>0</returns>
LRESULT CMainFrame::OnSettingChange(UINT, WPARAM, LPARAM, BOOL &)
{
	ApplyConfChanges();
	return S_OK;
}

// Fill current menu with accelerators' text
void CMainFrame::FillMenuWithHkeys(HMENU menu)
{
	for (unsigned int i = 0; i < _Settings.m_hotkey_groups.size(); ++i)
	{
		std::vector<CHotkey>::iterator begin = _Settings.m_hotkey_groups.at(i).m_hotkeys.begin();
		//skip scripts&plugins
		if (_Settings.m_hotkey_groups.at(i).m_reg_name == L"Scripts" || _Settings.m_hotkey_groups.at(i).m_reg_name == L"Plugins")
			begin++;

		std::sort(begin, _Settings.m_hotkey_groups.at(i).m_hotkeys.end());
		for (unsigned int j = 0; j < _Settings.m_hotkey_groups.at(i).m_hotkeys.size(); ++j)
		{
			CString text;
			WORD cmd = _Settings.m_hotkey_groups.at(i).m_hotkeys.at(j).m_accel.cmd;

			if (::GetMenuString(menu, cmd, text.GetBufferSetLength(MAX_LOAD_STRING + 1), MAX_LOAD_STRING + 1,
				MF_BYCOMMAND))
			{
				text.ReleaseBuffer();
				text += L"\t";
				text += U::AccelToString(_Settings.m_hotkey_groups.at(i).m_hotkeys.at(j).m_accel);

				MENUITEMINFO miim;
				ZeroMemory(&miim, sizeof(MENUITEMINFO));
				miim.cbSize = sizeof(MENUITEMINFO);
				miim.fMask = MIIM_STRING;
				miim.dwTypeData = text.GetBuffer();
				miim.cch = text.GetLength();
				::SetMenuItemInfo(menu, cmd, FALSE, &miim);
			}
		}
	}
}

CFBEView &CMainFrame::ActiveView()
{
	return m_doc->m_body;
}

/// <summary>WM Handler(WM_CONTEXT_MENU) Handle toolbar popup menu</summary>
LRESULT CMainFrame::OnContextMenu(UINT, WPARAM, LPARAM lParam, BOOL &)
{
	HMENU menu, popup;
	RECT rect;
	CPoint ptMousePos = (CPoint)lParam;
	ScreenToClient(&ptMousePos);
	// find clicked toolbar
	REBARBANDINFO rbi;
	ZeroMemory((void *)&rbi, sizeof(rbi));
	rbi.cbSize = sizeof(REBARBANDINFO);
	rbi.fMask = RBBIM_ID;
	m_selBandID = -1;
	for (unsigned int i = 0; i < m_rebar.GetBandCount(); i++)
	{
		m_rebar.GetRect(i, &rect);
		if (PtInRect(&rect, ptMousePos))
		{
			m_rebar.GetBandInfo(i, &rbi);
			m_selBandID = rbi.wID;
			break;
		}
	}
	// display context menu for command & script toolbars only
	if ((m_selBandID == ATL_IDW_BAND_FIRST + 1) || (m_selBandID == ATL_IDW_BAND_FIRST + 2))
	{
		menu = ::LoadMenu(_Module.GetResourceInstance(), MAKEINTRESOURCEW(IDR_TOOLBAR_MENU));
		popup = ::GetSubMenu(menu, 0);
		ClientToScreen(&ptMousePos);
		::TrackPopupMenu(popup, TPM_LEFTALIGN, ptMousePos.x, ptMousePos.y, 0, *this, 0);
	}
	return S_OK;
}

LRESULT CMainFrame::OnUnhandledCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	HWND hFocus = ::GetFocus();
	UINT idCtl = HIWORD(wParam);

	// only pass messages to the editors
	if (idCtl == 0 || idCtl == 1)
	{
		if (
			hFocus == m_id || hFocus == m_href || hFocus == m_section || ::IsChild(m_id, hFocus) || ::IsChild(m_href, hFocus) || ::IsChild(m_section, hFocus) || hFocus == m_styleT_table || hFocus == m_id_table_id || hFocus == m_id_table || hFocus == m_style_table || hFocus == m_colspan_table || hFocus == m_rowspan_table || hFocus == m_align_table || hFocus == m_valign_table || hFocus == m_alignTR_table || hFocus == m_image_title || ::IsChild(m_id_table_id, hFocus) || ::IsChild(m_id_table, hFocus) || ::IsChild(m_style_table, hFocus) || ::IsChild(m_styleT_table, hFocus) || ::IsChild(m_colspan_table, hFocus) || ::IsChild(m_rowspan_table, hFocus) || ::IsChild(m_alignTR_table, hFocus) || ::IsChild(m_align_table, hFocus) || ::IsChild(m_valign_table, hFocus) || ::IsChild(m_image_title, hFocus))
			return ::SendMessage(hFocus, WM_COMMAND, wParam, lParam);

		// We need to check that the focused window is a web browser indeed
		if (hFocus == m_view.GetActiveWnd() || ::IsChild(m_view.GetActiveWnd(), hFocus))
		{
			if (m_current_view == SOURCE)
			{
				switch (LOWORD(wParam))
				{
					/*case ID_EDIT_UNDO:
						m_source.SendMessage(SCI_UNDO);
						break;*/
				case ID_EDIT_REDO:
					m_source.SendMessage(SCI_REDO);
					break;
					/*case ID_EDIT_CUT:
						m_source.SendMessage(SCI_CUT);
						break;
					case ID_EDIT_COPY:
						m_source.SendMessage(SCI_COPY);
						break;
					case ID_EDIT_PASTE:
						m_source.SendMessage(SCI_PASTE);
						break;*/
				case ID_EDIT_FIND:
				{
					if (!m_sci_find_dlg)
						m_sci_find_dlg = new CSciFindDlg(&m_doc->m_body, m_source);

					if (m_sci_find_dlg->IsValid())
						break;

					m_sci_find_dlg->UpdatePattern();

					m_sci_find_dlg->ShowDialog();
					break;
				}
				case ID_EDIT_FINDNEXT:
					m_doc->m_body.SciFindNext(m_source, false, true);
					break;
				case ID_EDIT_REPLACE:
				{
					if (!m_sci_replace_dlg)
						m_sci_replace_dlg = new CSciReplaceDlg(&m_doc->m_body, m_source);

					if (m_sci_replace_dlg->IsValid())
						break;

					m_sci_replace_dlg->UpdatePattern();

					m_sci_replace_dlg->ShowDialog();
					break;
				}
				}
			}
			else
				return ActiveView().SendMessage(WM_COMMAND, wParam, 0);
		}

		if (hFocus == m_document_tree.m_hWnd || ::IsChild(m_document_tree.m_hWnd, hFocus))
			return m_doc->m_body.SendMessage(WM_COMMAND, wParam, 0);
	}

	// Last chance to send common commands to any focused window
	switch (LOWORD(wParam))
	{
	case ID_EDIT_UNDO:
		::SendMessage(hFocus, WM_UNDO, 0, 0);
		break;
	case ID_EDIT_REDO:
		::SendMessage(hFocus, EM_REDO, 0, 0);
		break;
	case ID_EDIT_CUT:
		::SendMessage(hFocus, WM_CUT, 0, 0);
		break;
	case ID_EDIT_COPY:
		::SendMessage(hFocus, WM_COPY, 0, 0);
		break;
	case ID_EDIT_PASTE:
		::SendMessage(hFocus, WM_PASTE, 0, 0);
		break;
	case ID_EDIT_INS_SYMBOL:
		::SendMessage(hFocus, WM_CHAR, wParam, 0);
		break;
	}

	return S_OK;
}

LRESULT CMainFrame::OnSetFocus(UINT, WPARAM, LPARAM, BOOL &)
{
	m_view.SetFocus();
	UpdateViewSizeInfo();
	return S_OK;
}

LRESULT CMainFrame::OnSetStatusText(UINT, WPARAM, LPARAM lParam, BOOL &)
{
	m_status_msg = (const TCHAR *)lParam;
	return S_OK;
}

/// <summary>WM Handler(WM_TRACKPOPUPMENU) Prepare and display popup menu in editor</summary>
/// <param name="lParam">Contain popup menu structure</param>
/// <returns>0</returns>
LRESULT CMainFrame::OnTrackPopupMenu(UINT, WPARAM, LPARAM lParam, BOOL &)
{
	U::TRACKPARAMS *tp = (U::TRACKPARAMS *)lParam;
	// add speller menu
	if (m_Speller)
	{
		//get suggestions for current selected word
		m_strSuggestions = m_Speller->GetSuggestions();
		if (m_strSuggestions)
		{
			int numSuggestions = m_strSuggestions->GetSize();
			if (numSuggestions > 0)
			{
				// limit up to MAX_POPUP_SUGGESTION suggestion
				numSuggestions = min(MAX_POPUP_SUGGESTION, numSuggestions);
				::AppendMenu(tp->hMenu, MF_SEPARATOR, 0, NULL);

				for (int i = 0; i < numSuggestions; i++)
					::AppendMenu(tp->hMenu, MF_STRING, IDC_SPELL_REPLACE + i, (*m_strSuggestions)[i]);

				// add standard spellcheck items
				if (numSuggestions > 0)
					::AppendMenu(tp->hMenu, MF_SEPARATOR, 0, NULL);
				CString itemName;
				itemName.LoadString(IDC_SPELL_IGNOREALL);
				::AppendMenu(tp->hMenu, MF_STRING, IDC_SPELL_IGNOREALL, itemName);

				itemName.LoadString(IDC_SPELL_ADD2DICT);
				::AppendMenu(tp->hMenu, MF_STRING, IDC_SPELL_ADD2DICT, itemName);
			}
		}
	}
	m_MenuBar.TrackPopupMenu(tp->hMenu, tp->uFlags, tp->x, tp->y);
	return S_OK;
}

LRESULT CMainFrame::OnPreCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	bHandled = FALSE;
	if ((HIWORD(wParam) == 0 || HIWORD(wParam) == 1) && LOWORD(wParam) != ID_EDIT_INCSEARCH)
		StopIncSearch(true);
	return S_OK;
}

LRESULT CMainFrame::OnDropFiles(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	HDROP hDrop = (HDROP)wParam;
	UINT nf = ::DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
	CString buf, ext;
	if (nf > 0)
	{
		UINT len = ::DragQueryFile(hDrop, 0, NULL, 0);
		TCHAR *cp = buf.GetBuffer(len + 1);
		len = ::DragQueryFile(hDrop, 0, cp, len + 1);
		buf.ReleaseBuffer(len);
	}
	::DragFinish(hDrop);
	if (!buf.IsEmpty())
	{
		ext.SetString(ATLPath::FindExtension(buf));
		if (ext.CompareNoCase(L".FB2") == 0)
		{
			if (LoadFile(buf) == OK)
			{
				m_mru.AddToList(m_doc->m_filename);
				PIDLIST_ABSOLUTE pidl = { 0 };
				HRESULT hr = SHParseDisplayName(m_doc->m_filename, NULL, &pidl, 0, NULL);
				if (SUCCEEDED(hr))
				{
					m_spShellItem.Release();
					SHCreateShellItem(NULL, NULL, pidl, &m_spShellItem);
					ILFree(pidl);
				}
			}
		}
		else if ((ext.CompareNoCase(L".JPG") == 0) || (ext.CompareNoCase(L".JPEG") == 0) || (ext.CompareNoCase(L".PNG") == 0))
		{
			m_doc->m_body.SetFocus();
			m_doc->m_body.AddImage(buf);
		}
	}
	return S_OK;
}

// drag & drop to the BODY window
LRESULT CMainFrame::OnNavigate(WORD, WORD, HWND, BOOL &)
{
	CString url(m_doc->m_body.NavURL());
	if (!url.IsEmpty())
	{
		CString ext(ATLPath::FindExtension(url));
		if (ext.CompareNoCase(L".FB2") == 0)
		{
			if (LoadFile(url) == OK)
				m_mru.AddToList(m_doc->m_filename);
		}
		else if ((ext.CompareNoCase(L".JPG") == 0) || (ext.CompareNoCase(L".JPEG") == 0) || (ext.CompareNoCase(L".PNG") == 0))
		{
			m_doc->m_body.AddImage(url);
		}
	}
	return S_OK;
}

///<summary>Recursivelly add all IDs to combobox.Internal function</summary>
/// <param name="box">XML document to validate</param>
/// <param name="node">XML node (on start=BODY)document to validate</param>
static void GrabIDs(CComboBox &box, MSHTML::IHTMLDOMNode *node)
{
	if (node->nodeType != NODE_ELEMENT)
		return;

	CString name = node->nodeName;

	if (name != L"P" && name != L"DIV" && name != L"BODY")
		return;

	// Add id to combobox
	MSHTML::IHTMLElementPtr elem(node);
	CString id = elem->id;
	if (!id.IsEmpty())
		box.AddString(L"#" + id);
	//			box.AddString(AddHash(tmp, id));

		// Add children's id to combobox
	MSHTML::IHTMLDOMNodePtr cn(node->firstChild);
	while (cn)
	{
		GrabIDs(box, cn);
		cn = cn->nextSibling;
	}
}

///<summary>Recursivelly add all IDs to combobox</summary>
/// <param name="box">combobox</param>
/// <param name="node">XML node (on start=BODY)document to validate</param>
void CMainFrame::ParaIDsToComboBox(CComboBox &box)
{
	try
	{
		//MSHTML::IHTMLDOMNodePtr body(m_body.Document()->body);
		GrabIDs(box, MSHTML::IHTMLDOMNodePtr(m_doc->m_body.Document()->body));
	}
	catch (_com_error &)
	{
	}
}

///<summary>Add all IDs to combobox</summary>
/// <param name="box">combobox</param>
void CMainFrame::BinIDsToComboBox(CComboBox &box)
{
	try
	{
		IDispatchPtr bo(m_doc->m_body.Document()->all->item(L"id"));
		if (!bo)
			return;
		// CString tmp;
		MSHTML::IHTMLElementCollectionPtr sbo(bo);
		if (sbo)
		{
			long l = sbo->length;
			for (long i = 0; i < l; ++i)
			{
				MSHTML::IHTMLElementPtr elem = sbo->item(i);
				CString value = elem->getAttribute(L"value", 0); //0=not case-sensitive
				if (!value.IsEmpty())
				{
					//const wchar_t* hash = AddHash(tmp, value.AllocSysString());
//						box.AddString(AddHash(tmp, value.AllocSysString()));
					box.AddString(L"#" + value);
				}
			}
		}
		else
		{
			MSHTML::IHTMLInputTextElementPtr ebo(bo);
			if (ebo)
				//					box.AddString(AddHash(tmp, ebo->value));
				box.AddString(L"#" + ebo->value);
		}
	}
	catch (_com_error &)
	{
	}
}

LRESULT CMainFrame::OnCbSetFocus(WORD, WORD, HWND, BOOL &)
{
	if (!m_cb_updated)
	{
		m_ignore_cb_changes = true;

		CString str(U::GetWindowText(m_href));

		m_href_box.ResetContent();
		m_href.SetWindowText(str);
		m_href.SetSel(0, str.GetLength() + 1);
		m_ignore_cb_changes = false;

		if (m_cb_last_images)
			BinIDsToComboBox(m_href_box);
		else
			ParaIDsToComboBox(m_href_box);
		m_cb_updated = true;
	}

	return S_OK;
}

///<summary>Display XML-error message dialog</summary>
/// <param name="Msg">Error message</param>
void CMainFrame::DisplayXMLError(const CString& Msg)
{
	wchar_t buf[MAX_LOAD_STRING + 1];
	::LoadString(_Module.GetResourceInstance(), IDS_VALIDATION_FAIL_CPT, buf, MAX_LOAD_STRING);
	CString ErrMessage(buf);
	ErrMessage += L"\n" + Msg;
	SendMessage(m_hWnd, U::WM_SETSTATUSTEXT, 0, (LPARAM)ErrMessage.GetString());
	MessageBeep(MB_ICONERROR);
	U::ReportError(Msg);
}

// commands
//-----------------File Commands -------------------------------
LRESULT CMainFrame::OnFileNew(WORD, WORD, HWND, BOOL &)
{
	if (!DiscardChanges())
		return S_OK;

	FB::Doc *doc = new FB::Doc(*this);
	FB::Doc::m_active_doc = doc;
	doc->CreateBlank(m_view);
	m_file_age = unsigned __int64(~0);
	if (m_spShellItem)
		m_spShellItem.Release();
	delete m_doc;
	m_doc = doc;
	AttachDocument(m_doc);

	return S_OK;
}

LRESULT CMainFrame::OnFileOpen(WORD, WORD, HWND, BOOL &bHandled)
{
	if (LoadFile() == OK)
		m_mru.AddToList(m_doc->m_filename);

	return S_OK;
}

LRESULT CMainFrame::OnFileOpenMRU(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	CString filename;
	m_mru.GetFromList(wID, filename);

	switch (LoadFile(filename))
	{
	case OK:
		m_mru.MoveToTop(wID);
		break;
	case FAIL:
		m_mru.RemoveFromList(wID);
		break;
	case CANCELLED:
		break;
	}

	return S_OK;
}

LRESULT CMainFrame::OnFileSave(WORD, WORD, HWND, BOOL &)
{
	SaveFile(false); //no SaveFile dialog
	return S_OK;
}

LRESULT CMainFrame::OnFileExit(WORD, WORD, HWND, BOOL &)
{
	// close (possible) opened in script modeless dialogs
	PostMessage(WM_CLOSEDIALOG);
	PostMessage(WM_CLOSE);
	return S_OK;
}

LRESULT CMainFrame::OnFileSaveAs(WORD, WORD, HWND, BOOL &)
{
	SaveFile(true); //show SaveFile dialog
	return S_OK;
}

LRESULT CMainFrame::OnFileValidate(WORD, WORD, HWND, BOOL &)
{
	int col = 0;
	int line = 0;
	CString ErrMsg(L"");
	bool fv = true; // validation flag true=valid
	if (!m_doc->m_body.IsSynched())
	{
		if (m_current_view == SOURCE)
		{
			// SOURCE view
			// convert text from Scintilla into xml
			int textlen = ::SendMessage(m_source, SCI_GETLENGTH, 0, 0);
			char *buffer = (char *)malloc(textlen + 1);
			if (!buffer)
			{
				U::ReportError(IDS_OUT_OF_MEM_MSG);
				return S_OK;
			}
			::SendMessage(m_source, SCI_GETTEXT, textlen + 1, (LPARAM)buffer);

			// convert to unicode
			DWORD ulen = ::MultiByteToWideChar(CP_UTF8, 0, buffer, textlen, NULL, 0);
			BSTR ustr = ::SysAllocStringLen(NULL, ulen);
			if (!ustr)
			{
				free(buffer);
				U::ReportError(IDS_OUT_OF_MEM_MSG);
				return S_OK;
			}
			::MultiByteToWideChar(CP_UTF8, 0, buffer, textlen, ustr, ulen);
			free(buffer);

			// check wellformed
			if (U::IsXMLWellFormed(ustr, col, line, ErrMsg))
				m_doc->LoadXMLFromString(ustr);
			else
				fv = false;
			::SysFreeString(ustr);
		}
		else
		{
			// BODY view
			m_doc->CreateDOM();
		}
	}

	if (fv)
		fv = m_doc->Validate(col, line, ErrMsg);

	if (!fv)
	{
		if (m_current_view != SOURCE)
			ShowView(SOURCE);
		SourceGoTo(line, col);
		DisplayXMLError(ErrMsg);
	}
	else
	{
		wchar_t buf[MAX_LOAD_STRING + 1];
		LoadString(_Module.GetResourceInstance(), IDS_SB_NO_ERR, buf, MAX_LOAD_STRING);
		SendMessage(m_hWnd, U::WM_SETSTATUSTEXT, 0, (LPARAM)buf);
		MessageBeep(MB_OK);
	}
	return S_OK;
}

LRESULT CMainFrame::OnToolsImport(WORD, WORD wID, HWND, BOOL &)
{
	/*	wID -= ID_IMPORT_BASE;
		if (wID < m_import_plugins.GetSize())
		{
			try
			{
				IUnknownPtr unk;
				CheckError(unk.CreateInstance(m_import_plugins[wID]));

				CComQIPtr<IFBEImportPlugin> ipl(unk);

				IDispatchPtr obj;
				_bstr_t filename;
				if (ipl)
				{
					m_last_plugin = wID + ID_EXPORT_BASE;
					BSTR bs = NULL;
					HRESULT hr = ipl->Import((long)m_hWnd, &bs, &obj);
					CheckError(hr);
					filename.Assign(bs);
					if (hr != S_OK)
						return S_OK;
				}
				else
				{
					AtlTaskDialog(*this, IDS_IMPORT_ERR_CPT, IDS_IMPORT_ERR_MSG, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
					return S_OK;
				}

				MSXML2::IXMLDOMDocument2Ptr dom(obj);
				if (!dom)
				{
					AtlTaskDialog(*this, IDS_ERRMSGBOX_CAPTION, IDS_IMPORT_XML_ERR_MSG, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
				}
				else if (DiscardChanges())
				{
					//FB::Doc *doc=new FB::Doc(*this);
					//FB::Doc::m_active_doc = doc;

					//if (doc->LoadFromDOM(m_view,dom)) {
					CComDispatchDriver body(m_doc->m_body.Script());
					CComVariant args[2];
					CComVariant res;
					args[1] = dom.GetInterfacePtr();
					args[0] = _Settings.GetInterfaceLanguageName();
					CheckError(body.InvokeN(L"LoadFromDOM", args, 2, &res));
					if (res.boolVal)
					//if (doc->LoadFromHTML(m_view,(const wchar_t* )filename))
					{
						if (filename.length() > 0)
						{
							m_doc->m_filename = (const TCHAR *)filename;
							wchar_t str[MAX_PATH];
							wcscpy(str, (const wchar_t *)filename);
							PathRemoveFileSpec(str);
							SetCurrentDirectory(str);
							if (m_doc->m_filename.GetLength() < 4 || m_doc->m_filename.Right(4).CompareNoCase(_T(".fb2")) != 0)
								m_doc->m_filename += _T(".fb2");
						}
						AttachDocument(doc);
						delete m_doc;
						m_doc=doc;
						m_doc->m_body.Init();
						m_doc->MarkSavePoint();
					} // else
					  //FB::Doc::m_active_doc = m_doc;
					  //delete doc;
				}
			}
			catch (_com_error &e)
			{
				U::ReportError(e);
			}
		}*/
	return S_OK;
}

LRESULT CMainFrame::OnToolsExport(WORD, WORD wID, HWND, BOOL &)
{
	/*	wID -= ID_EXPORT_BASE;
		if (wID < m_export_plugins.GetSize())
		{
			try
			{
				IUnknownPtr unk;
				CheckError(unk.CreateInstance(m_export_plugins[wID]));

				CComQIPtr<IFBEExportPlugin> epl(unk);

				if (epl)
				{
					m_last_plugin = wID + ID_EXPORT_BASE;
					m_doc->CreateDOM();
					_bstr_t filename;
					CString tmp(m_doc->m_filename);
					if (tmp.GetLength() >= 4 && tmp.Right(4).CompareNoCase(_T(".fb2")) == 0)
						tmp.Delete(tmp.GetLength() - 4, 4);
					filename = (const TCHAR *)tmp;
					if (m_doc->)
						CheckError(epl->Export((long)m_hWnd, filename, IDispatchPtr(dom)));
				}
				else
				{
					AtlTaskDialog(*this, IDS_EXPORT_ERR_CPT, IDS_EXPORT_ERR_MSG, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
					return S_OK;
				}
			}
			catch (_com_error &e)
			{
				U::ReportError(e);
			}
		}*/
	return S_OK;
}

//-----------------View Commands -------------------------------
LRESULT CMainFrame::OnViewToolBar(WORD, WORD wID, HWND, BOOL &)
{
	int nBandIndex = m_rebar.IdToIndex(wID);
	BOOL bVisible = !IsBandVisible(wID);
	m_rebar.ShowBand(nBandIndex, bVisible);
	UISetCheck(wID, bVisible);

	if (wID == 60164 || wID == 60165)
	{
		if (wID == 60164)
			wID++;
		else
			wID--;
		nBandIndex = m_rebar.IdToIndex(wID);
		m_rebar.ShowBand(nBandIndex, bVisible);
		UISetCheck(wID, bVisible);
	}
	UpdateLayout();
	return S_OK;
}

LRESULT CMainFrame::OnViewStatusBar(WORD, WORD, HWND, BOOL &)
{
	BOOL bVisible = !m_status.IsWindowVisible();
	::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
	UpdateLayout();
	return S_OK;
}

LRESULT CMainFrame::OnViewDesc(WORD, WORD, HWND, BOOL &)
{
	ShowView(DESC);
	return S_OK;
}

LRESULT CMainFrame::OnViewBody(WORD, WORD, HWND, BOOL &)
{
	ShowView(BODY);
	return S_OK;
}

LRESULT CMainFrame::OnViewSource(WORD, WORD, HWND, BOOL &)
{
	ShowView(SOURCE);
	return S_OK;
}

LRESULT CMainFrame::OnViewTree(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == SOURCE)
		return S_OK;

	BOOL bVisible = !_Settings.ViewDocumentTree();
	m_document_tree.ShowWindow(bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	if (bVisible)
		GetDocumentStructure();
	UISetCheck(ID_VIEW_TREE, bVisible);
	m_splitter.SetSinglePaneMode(bVisible ? SPLIT_PANE_NONE : SPLIT_PANE_RIGHT);
	_Settings.SetViewDocumentTree(bVisible != 0, TRUE);

	return S_OK;
}

LRESULT CMainFrame::OnViewOptions(WORD, WORD, HWND, BOOL &)
{
	int find_repl = TempCloseDialogs();

	if (ShowSettingsDialog(m_hWnd))
		ApplyConfChanges();

	RestoreDialogs(find_repl);

	return S_OK;
}

//-----------------Service Commands -------------------------------
// close find& replace dialog
//return flag, which dialog was closed
int CMainFrame::TempCloseDialogs()
{
	bool bFind = m_doc->m_body.CloseFindDialog(m_doc->m_body.m_find_dlg);
	bool bReplace = m_doc->m_body.CloseFindDialog(m_doc->m_body.m_replace_dlg);

	bool bSciFind = m_doc->m_body.CloseFindDialog(m_sci_find_dlg);
	bool bSciRepl = m_doc->m_body.CloseFindDialog(m_sci_replace_dlg);

	return bFind || bSciFind ? 1 : 0 + bReplace || bSciRepl ? 2 : 0;
}

// restore temporary closed dialog
// flag, which dialog was closed before
void CMainFrame::RestoreDialogs(int dialog_type)
{
	switch (dialog_type)
	{
	case 1: //find dialog
		SendMessage(WM_COMMAND, ID_EDIT_FIND, NULL);
		break;
	case 2: //replace dialog
		SendMessage(WM_COMMAND, ID_EDIT_REPLACE, NULL);
		break;
	}
}

LRESULT CMainFrame::OnToolsWords(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == SOURCE)
		return S_OK;

	int find_repl = TempCloseDialogs();

	ShowWordsDialog(*m_doc, m_hWnd);

	RestoreDialogs(find_repl);

	return S_OK;
}

LRESULT CMainFrame::OnToolsOptions(WORD, WORD, HWND, BOOL &)
{
	int find_repl = TempCloseDialogs();

	if (ShowSettingsDialog(m_hWnd))
		ApplyConfChanges();

	RestoreDialogs(find_repl);

	return S_OK;
}

LRESULT CMainFrame::OnToolsScript(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	wID -= ID_SCRIPT_BASE;

	if (m_current_view == SOURCE)
		return S_OK;

	// скрипты от FBE и от FBW запускаются по разному. В FBE скрипты исполняются через Active Scripting
	// и документ в него передаетяся через параметры.
	// В FBW скрипты выполняются в самом HTML документе
	for (int i = 0; i < m_scripts.GetSize(); ++i)
	{
		if (m_scripts[i].wID == -1)
			continue;

		if (m_scripts[i].Type == 2 && m_scripts[i].wID == wID)
		{
			m_doc->RunScript(m_scripts[i].path.GetBuffer());
			m_last_script = &m_scripts[i];
			break;
		}
	}

	// TODO тут должен быть else

	/*if (wID < m_scripts.GetSize()) {
	if (StartScript(this) >= 0) {
		  if (SUCCEEDED(ScriptLoad(m_scripts[wID].name))){
			  if(m_scripts[wID].Type == 0)
			  {
				  MSXML2::IXMLDOMDocument2Ptr dom(m_doc->CreateDOM(m_doc->m_encoding));
				  if (dom)
				  {
					  CComVariant arg;
					  V_VT(&arg) = VT_DISPATCH;
					  V_DISPATCH(&arg) = dom;
					  dom.AddRef();
					  if (SUCCEEDED(ScriptCall(L"Run",&arg,1,NULL)))
					  {
						  m_doc->SetXML(dom);
					  }
				  }
			  }
			  else if(m_scripts[wID].Type == 1)
			  {
				  SHD::IWebBrowser2Ptr HTMLdomBody = m_doc->m_body.Browser();
				  SHD::IWebBrowser2Ptr HTMLdomDesc = m_doc->m_body.Browser();
				  CComVariant* arg = new CComVariant[2];
				  V_VT(&arg[0]) = VT_DISPATCH;
				  V_DISPATCH(&arg[0]) = HTMLdomBody;
				  HTMLdomBody.AddRef();
				  V_VT(&arg[1]) = VT_DISPATCH;
				  V_DISPATCH(&arg[1]) = HTMLdomDesc;
				  HTMLdomDesc.AddRef();

				  CComVariant vt;
				  if (SUCCEEDED(ScriptCall(L"Run",arg,2,&vt)))
				  {
					  //m_doc->SetXML(dom);
				  }
			  }
			  else if(m_scripts[wID].Type == 2)
			  {
				  ScriptCall(L"Run",0,0,0);
			  }
		}
		StopScript();
	  }
	}*/

	return S_OK;
}

LRESULT CMainFrame::OnHideToolbar(WORD wNotifyCode, WORD /*wID*/, HWND hWndCtl, BOOL &bHandled)
{
	return OnViewToolBar(wNotifyCode, WORD(m_selBandID), hWndCtl, bHandled);
}

LRESULT CMainFrame::OnToolCustomize(WORD /*wNotifyCode*/, WORD /*wID*/, HWND hWndCtl, BOOL & /*bHandled*/)
{
	if (m_selBandID == ATL_IDW_BAND_FIRST + 1)
		m_CmdToolbar.Customize();
	else if (m_selBandID == ATL_IDW_BAND_FIRST + 2)
		m_ScriptsToolbar.Customize();

	return S_OK;
}

LRESULT CMainFrame::OnLastScript(WORD, WORD, HWND, BOOL &)
{
	if (m_last_script != 0 && m_current_view != SOURCE)
	{
		m_doc->RunScript((*m_last_script).path.GetBuffer());
	}
	return S_OK;
}

LRESULT CMainFrame::OnEditFind(WORD, WORD, HWND, BOOL &bHandled)
{
	if (m_current_view == DESC)
		ShowView(BODY);

	bHandled = FALSE;
	return S_OK;
}

LRESULT CMainFrame::OnEditInsSymbol(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	static int symHkGroup = -1;
	if (symHkGroup == -1)
	{
		for (unsigned int i = 0; i < _Settings.m_hotkey_groups.size(); ++i)
		{
			if (_Settings.m_hotkey_groups[i].m_reg_name == L"Symbols")
			{
				symHkGroup = i;
				break;
			}
		}
	}
	std::vector<CHotkey> &symHotkeys = _Settings.m_hotkey_groups[symHkGroup].m_hotkeys;

	wchar_t c = NULL;
	for (unsigned int i = 0; i < symHotkeys.size(); ++i)
	{
		if (symHotkeys[i].m_accel.cmd == wID)
		{
			c = symHotkeys[i].m_char_val;
			break;
		}
	}

	if (c)
	{
		//HWND aw = ::GetFocus();
		::SendMessage(::GetFocus(), WM_CHAR, c, NULL);

		/*IServiceProviderPtr ServiceProvider;
		ServiceProvider = m_doc->m_body.Browser();
		if(ServiceProvider)
		{
			IOleWindowPtr Window = NULL;
			if(SUCCEEDED(ServiceProvider->QueryService(SID_SShellBrowser, IID_IOleWindow, (void**)&Window)))
			{
				HWND hwndBrowser = NULL;
				if (SUCCEEDED(Window->GetWindow(&hwndBrowser)))
				{
					while(::GetWindow(hwndBrowser, GW_CHILD))
						hwndBrowser = ::GetWindow(hwndBrowser, GW_CHILD);
					::SendMessage(hwndBrowser, WM_CHAR, c, 0);
				}
			}
		}*/
	}

	return S_OK;
}

// Spellcheck commands

/// <summary>WM Handler(ID_TOOLS_SPELLCHECK) Start spellcheck</summary>
LRESULT CMainFrame::OnSpellCheck(WORD, WORD, HWND, BOOL &b)
{
	if (m_current_view == BODY)
		m_Speller->StartDocumentCheck(m_doc->m_body.m_mk_srv);
	return S_OK;
}

/// <summary>WM Handler(ID_TOOLS_SPELLCHECK_HIGHLIGHT) Toggle misspell highlights</summary>
LRESULT CMainFrame::OnToggleHighlight(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == BODY)
	{
		_Settings.SetHighlightMisspells(!_Settings.m_highlght_check);
	}
	return S_OK;
}

/// <summary>WM Handler(IDC_SPELL_REPLACE+x) Replace spell from menu item</summary>
/// <param name="wID">Menu identifier (IDM_*)</param>
/// <returns>0</returns>
LRESULT CMainFrame::OnSpellReplace(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	m_doc->m_body.BeginUndoUnit(L"replace word");
	CString replace = (*m_strSuggestions)[wID - IDC_SPELL_REPLACE]; //index=menu position
	m_Speller->Replace(replace);
	m_doc->m_body.EndUndoUnit();
	return S_OK;
}

LRESULT CMainFrame::OnSpellIgnoreAll(WORD, WORD, HWND, BOOL &)
{
	m_Speller->IgnoreAll();
	return S_OK;
}

LRESULT CMainFrame::OnSpellAddToDict(WORD, WORD, HWND, BOOL &)
{
	m_Speller->AddToDictionary();
	return S_OK;
}

LRESULT CMainFrame::OnDocChanged(WORD, WORD, HWND, BOOL &)
{
	m_doc->Changed(true);
	m_doc->m_body.Synched(false);
	return S_OK;
}

LRESULT CMainFrame::OnAppAbout(WORD, WORD, HWND, BOOL &)
{
	int find_repl = TempCloseDialogs();

	CAboutDlg dlg;
	dlg.DoModal();

	RestoreDialogs(find_repl);

	return S_OK;
}

// Navigation
LRESULT CMainFrame::OnSelectCtl(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	switch (wID)
	{
	case ID_SELECT_TREE:
		if (!m_document_tree.IsWindowVisible())
			OnViewTree(0, 0, 0, bHandled);
		m_document_tree.m_tree.m_tree.SetFocus();
		break;
	case ID_SELECT_ID:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_id.SetFocus();
		break;
	case ID_SELECT_HREF:
	{
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_href.SetFocus();
		CString href(U::GetWindowText(m_href));
		m_href.SetSel(0, href.GetLength(), FALSE);
		break;
	}
	case ID_SELECT_IMAGE:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_image_title.SetFocus();
		break;
	case ID_SELECT_TEXT:
		m_view.SetFocus();
		break;
	case ID_SELECT_SECTION:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_section.SetFocus();
		break;
	case ID_SELECT_IDT:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_id_table_id.SetFocus();
		break;
	case ID_SELECT_STYLET:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_styleT_table.SetFocus();
		break;
	case ID_SELECT_STYLE:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 3))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 3, NULL, bHandled);
		m_style_table.SetFocus();
		break;
	case ID_SELECT_COLSPAN:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 4))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 4, NULL, bHandled);
		m_colspan_table.SetFocus();
		break;
	case ID_SELECT_ROWSPAN:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 4))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 4, NULL, bHandled);
		m_rowspan_table.SetFocus();
		break;
	case ID_SELECT_ALIGNTR:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 4))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 4, NULL, bHandled);
		m_alignTR_table.SetFocus();
		break;
	case ID_SELECT_ALIGN:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 4))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 4, NULL, bHandled);
		m_align_table.SetFocus();
		break;
	case ID_SELECT_VALIGN:
		if (!IsBandVisible(ATL_IDW_BAND_FIRST + 4))
			OnViewToolBar(0, ATL_IDW_BAND_FIRST + 4, NULL, bHandled);
		m_valign_table.SetFocus();
		break;
	}

	return S_OK;
}

LRESULT CMainFrame::OnNextItem(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	ShowView(NEXT);
	return 1;
}

LRESULT CMainFrame::OnEdSelChange(WORD, WORD, HWND hWndCtl, BOOL &)
{
	m_sel_changed = true;
	StopIncSearch(true);
	DisplayCharCode();
	return S_OK;
}

LRESULT CMainFrame::OnEdStatusText(WORD, WORD, HWND hWndCtl, BOOL &)
{
	StopIncSearch(true);
	m_status.SetText(ID_DEFAULT_PANE, (const TCHAR *)hWndCtl);
	return S_OK;
}
LRESULT CMainFrame::OnEdWantFocus(WORD, WORD wID, HWND, BOOL &)
{
	m_want_focus = wID;
	return S_OK;
}
LRESULT CMainFrame::OnEdReturn(WORD, WORD, HWND, BOOL &)
{
	m_view.SetFocus();
	return S_OK;
}

// combobox editor notifications

LRESULT CMainFrame::OnCbEdChange(WORD code, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	if (m_ignore_cb_changes)
		return S_OK;

	try
	{
		if (wID == IDC_HREF)
		{
			MSHTML::IHTMLElementPtr an(m_doc->m_body.SelectionsStruct(L"anchor"));
			_variant_t href;
			if (an)
				href = an->getAttribute(L"href", 2);
			if ((bool)an && V_VT(&href) == VT_BSTR)
			{
				CString newhref(U::GetWindowText(m_href));

				// changed by SeNS: href's fix - by default internal hrefs begins from '#'
				// otherwise set http protocol (if no other protocols specified)
				if (!newhref.IsEmpty() && (newhref[0] != L'#'))
				{
					if (newhref.Find(L"://") < 0)
						newhref = L"http://" + newhref;
				}

				if ((U::scmp(an->tagName, L"DIV") == 0) || (U::scmp(an->tagName, L"SPAN") == 0)) // must be an image
				{
					U::ChangeAttribute(an, L"href", newhref);
					MSHTML::IHTMLElementPtr img = MSHTML::IHTMLDOMNodePtr(an)->firstChild;
					//					m_doc->m_body.ImgSetURL(img, newhref);
					U::ChangeAttribute(img, L"src", L"fbw-internal:" + newhref);
					img->removeAttribute(L"width", 0L);

					MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_doc->m_body.Document())->getSelection();
					MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(m_doc->m_body.Document())->createRange();
					range->selectNodeContents(an);
					sel->removeAllRanges();
					sel->addRange(range);
					/*IHTMLControlRangePtr r(((MSHTML::IHTMLElement2Ptr)(m_doc->m_body.Document()->body))->createControlRange());
					r->add((IHTMLControlElementPtr)img->parentElement);
					r->select();*/
				}
				else
				{
					U::ChangeAttribute(an, L"href", newhref);
					MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_doc->m_body.Document())->getSelection();
					MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(m_doc->m_body.Document())->createRange();
					range->selectNodeContents(an);
					sel->removeAllRanges();
					sel->addRange(range);

					/*					MSHTML::IHTMLTxtRangePtr r = m_doc->m_body.Document()->selection->createRange();
										r->moveToElementText(an);
										r->select();*/
				}
			}
			else
			{
				m_href_box.SetWindowText(_T(""));
				m_href_box.EnableWindow(FALSE);
				m_href_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_ID)
		{
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStruct(L""));
			if (sc)
				sc->id = (const wchar_t *)U::GetWindowText(m_id);
			else
			{
				m_id_box.EnableWindow(FALSE);
				m_id_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_SECTION)
		{
			MSHTML::IHTMLElementPtr scs(m_doc->m_body.SelectionsStruct(L"section"));
			if (scs)
				scs->id = (const wchar_t *)U::GetWindowText(m_section);
			else
				m_section.EnableWindow(FALSE);
		}

		if (wID == IDC_IMAGE_TITLE)
		{
			MSHTML::IHTMLElementPtr scs(m_doc->m_body.SelectionsStruct(L"image"));
			if (scs)
			{
				//scs->title=(const wchar_t *)U::GetWindowText(m_image_title);
				U::ChangeAttribute(scs, L"title", (const wchar_t *)U::GetWindowText(m_image_title));

				IHTMLControlRangePtr r(((MSHTML::IHTMLElement2Ptr)(m_doc->m_body.Document()->body))->createControlRange());
				r->add((IHTMLControlElementPtr)scs);
				r->select();
			}
			else
			{
				m_image_title_box.EnableWindow(FALSE);
				m_image_title_caption.SetEnabled(false);
			}
		}

		if (wID == IDC_IDT)
		{
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStruct(L"table"));
			if (sc)
				sc->id = (const wchar_t *)U::GetWindowText(m_id_table_id);
			else
			{
				m_id_table_id_box.EnableWindow(FALSE);
				m_table_id_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_ID)
		{
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStruct(L"th"));
			if (sc)
				sc->id = (const wchar_t *)U::GetWindowText(m_id_table);
			else
			{
				m_id_table_box.EnableWindow(FALSE);
				m_id_table_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_STYLET)
		{
			_bstr_t style("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"table", style));
			if (sc)
			{
				CString newsSyleT(U::GetWindowText(m_styleT_table));
				sc->setAttribute(L"fbstyle", _variant_t((const wchar_t *)newsSyleT), 0);
			}
			else
			{
				m_style_table_box.EnableWindow(FALSE);
				m_style_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_STYLE)
		{
			_bstr_t style("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"th", style));
			if (sc)
			{
				CString newsSyle(U::GetWindowText(m_style_table));
				sc->setAttribute(L"fbstyle", _variant_t((const wchar_t *)newsSyle), 0);
			}
			else
			{
				m_style_table_box.EnableWindow(FALSE);
				m_style_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_COLSPAN)
		{
			_bstr_t colspan("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"colspan", colspan));
			if (sc)
			{
				CString newsColspan(U::GetWindowText(m_colspan_table));
				sc->setAttribute(L"fbcolspan", _variant_t((const wchar_t *)newsColspan), 0);
			}
			else
			{
				m_colspan_table_box.EnableWindow(FALSE);
				m_colspan_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_ROWSPAN)
		{
			_bstr_t rowspan("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"rowspan", rowspan));
			if (sc)
			{
				CString newsRowspan(U::GetWindowText(m_rowspan_table));
				sc->setAttribute(L"fbrowspan", _variant_t((const wchar_t *)newsRowspan), 0);
			}
			else
			{
				m_rowspan_table_box.EnableWindow(FALSE);
				m_rowspan_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_ALIGNTR)
		{
			_bstr_t alignTR("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"tralign", alignTR));
			if (sc)
			{
				CString newsAlignTR(U::GetWindowText(m_alignTR_table));
				sc->setAttribute(L"fbalign", _variant_t((const wchar_t *)newsAlignTR), 0);
			}
			else
			{
				m_alignTR_table_box.EnableWindow(FALSE);
				m_tr_allign_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_ALIGN)
		{
			_bstr_t align("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"align", align));
			if (sc)
			{
				CString newsAlign(U::GetWindowText(m_align_table));
				sc->setAttribute(L"fbalign", _variant_t((const wchar_t *)newsAlign), 0);
			}
			else
			{
				m_align_table_box.EnableWindow(FALSE);
				m_th_allign_caption.SetEnabled(false);
			}
		}
		if (wID == IDC_VALIGN)
		{
			_bstr_t valign("");
			MSHTML::IHTMLElementPtr sc(m_doc->m_body.SelectionsStructB(L"valign", valign));
			if (sc)
			{
				CString newsVAlign(U::GetWindowText(m_valign_table));
				sc->setAttribute(L"fbvalign", _variant_t((const wchar_t *)newsVAlign), 0);
			}
			else
			{
				m_valign_table_box.EnableWindow(FALSE);
				m_valign_caption.SetEnabled(false);
			}
		}
	}
	catch (_com_error &)
	{
	}

	return S_OK;
}

LRESULT CMainFrame::OnCbSelEndOk(WORD code, WORD wID, HWND hWnd, BOOL &)
{
	PostMessage(WM_COMMAND, MAKELONG(wID, CBN_EDITCHANGE), (LPARAM)hWnd);
	return S_OK;
}

LRESULT CMainFrame::OnEdKillFocus(WORD, WORD, HWND, BOOL &)
{
	StopIncSearch(true);
	return S_OK;
}

// tree view notifications

LRESULT CMainFrame::OnTreeReturn(WORD, WORD, HWND, BOOL &)
{
	GoToSelectedTreeItem();
	return S_OK;
}

LRESULT CMainFrame::OnTreeUpdate(WORD, WORD, HWND, BOOL &)
{
	GetDocumentStructure();
	return S_OK;
}

LRESULT CMainFrame::OnTreeRestore(WORD, WORD, HWND, BOOL &b)
{
	GetDocumentStructure();
	return S_OK;
}

LRESULT CMainFrame::OnTreeMoveElement(WORD, WORD, HWND, BOOL &)
{
	m_doc->m_body.BeginUndoUnit(L"structure editing");
	CTreeItem from = m_document_tree.m_tree.m_tree.GetMoveElementFrom();
	CTreeItem to = m_document_tree.m_tree.m_tree.GetMoveElementTo();

	MSHTML::IHTMLElementPtr elemFrom = (MSHTML::IHTMLElement *)from.GetData();
	MSHTML::IHTMLElementPtr elemTo;

	MSHTML::IHTMLDOMNodePtr nodeFrom = (MSHTML::IHTMLDOMNodePtr)elemFrom;
	MSHTML::IHTMLDOMNodePtr nodeTo;
	MSHTML::IHTMLDOMNodePtr nodeInsertBefore;

	switch (m_document_tree.m_tree.m_tree.m_insert_type)
	{
	case CTreeView::child:
	{
		elemTo = (MSHTML::IHTMLElement *)to.GetData();
		nodeTo = (MSHTML::IHTMLDOMNodePtr)elemTo;
		nodeInsertBefore = nodeTo->firstChild;
		break;
	}

	case CTreeView::sibling:
	{
		elemTo = (MSHTML::IHTMLElement *)to.GetData();
		nodeTo = (MSHTML::IHTMLDOMNodePtr)elemTo;
		nodeInsertBefore = nodeTo->nextSibling;
		nodeTo = nodeTo->parentNode;
		break;
	}

	case CTreeView::none:
	{
		m_doc->m_body.EndUndoUnit();
		return S_OK;
	}
	}

	if (!IsNodeSection(nodeFrom) || !IsNodeSection(nodeTo))
	{
		m_doc->m_body.EndUndoUnit();
		return S_OK;
	}

	if (!U::IsEmptySection(nodeTo))
	{
		MSHTML::IHTMLDOMNodePtr new_section = CreateNestedSection(nodeTo);
		if (!bool(new_section))
		{
			m_doc->m_body.EndUndoUnit();
			return S_OK;
		}

		nodeInsertBefore = new_section->nextSibling;
	}
	m_doc->MoveNode(nodeFrom, nodeTo, nodeInsertBefore);
	m_document_tree.UpdateDocumentStructure(m_doc->m_body.Document(), nodeTo);
	m_doc->m_body.EndUndoUnit();
	return S_OK;
}

LRESULT CMainFrame::OnTreeMoveElementOne(WORD, WORD, HWND, BOOL &)
{
	m_doc->m_body.BeginUndoUnit(L"structure editing");
	CTreeItem item = m_document_tree.m_tree.m_tree.GetFirstSelectedItem();
	MSHTML::IHTMLElementPtr elem = 0;
	MSHTML::IHTMLDOMNodePtr ret_node = 0;

	do
	{
		if (item.IsNull())
			break;

		if (!item.GetData() || !(bool)(elem = (IHTMLElement *)item.GetData()))
			continue;

		MSHTML::IHTMLDOMNodePtr node = (MSHTML::IHTMLDOMNodePtr)elem;
		if (!(bool)node)
			continue;

		ret_node = MoveRightElementWithoutChildren(node);
	} while (item = m_document_tree.m_tree.m_tree.GetNextSelectedItem(item));

	GetDocumentStructure();
	if ((bool)ret_node)
	{
		MSHTML::IHTMLElementPtr elem1(ret_node);
		m_document_tree.m_tree.m_tree.SelectElement(elem1);
		GoTo(elem1);
	}

	m_doc->m_body.EndUndoUnit();
	return S_OK;
}

LRESULT CMainFrame::OnTreeMoveLeftElement(WORD, WORD, HWND, BOOL &)
{
	m_doc->m_body.BeginUndoUnit(L"structure editing");
	CTreeItem item = m_document_tree.m_tree.m_tree.GetLastSelectedItem();
	MSHTML::IHTMLElementPtr elem = 0;
	MSHTML::IHTMLDOMNodePtr ret_node;

	do
	{
		if (item.IsNull())
			break;

		if (!item.GetData() || !(bool)(elem = (IHTMLElement *)item.GetData()))
			continue;

		MSHTML::IHTMLDOMNodePtr node = (MSHTML::IHTMLDOMNodePtr)elem;
		if (!(bool)node)
			continue;

		ret_node = MoveLeftElement(node);
	} while (item = m_document_tree.m_tree.m_tree.GetPrevSelectedItem(item));

	GetDocumentStructure();
	if (ret_node)
	{
		MSHTML::IHTMLElementPtr elem1(ret_node);
		m_document_tree.m_tree.m_tree.SelectElement(elem1);
		GoTo(elem1);
	}

	m_doc->m_body.EndUndoUnit();
	return S_OK;
}

LRESULT CMainFrame::OnTreeMoveElementSmart(WORD, WORD, HWND, BOOL &)
{
	// если выделен только один элемент, то двигаем его вправо
	// если несколько, то проверяем братья они или нет
	// если братья, то делаем такую фичу
	// ----------
	// ----------
	//   ----------    первый выделенный элемент
	//   ----------
	// ----------
	//   ----------    второй выделенный элемент
	//   ----------

	// для небратьев делаем то же что и для одного

	m_doc->m_body.BeginUndoUnit(L"structure editing");
	CTreeItem item = m_document_tree.m_tree.m_tree.GetFirstSelectedItem();

	MSHTML::IHTMLDOMNodePtr node = RecursiveMoveRightElement(item);
	GetDocumentStructure();
	if ((bool)node)
	{
		MSHTML::IHTMLElementPtr elem(node);
		m_document_tree.m_tree.m_tree.SelectElement(elem);
		GoTo(elem);
	}

	m_doc->m_body.EndUndoUnit();

	return S_OK;
}

LRESULT CMainFrame::OnTreeViewElement(WORD, WORD, HWND, BOOL &)
{
	GoToSelectedTreeItem();
	return S_OK;
}

LRESULT CMainFrame::OnTreeViewElementSource(WORD, WORD, HWND, BOOL &)
{
	CTreeItem item = m_document_tree.GetSelectedItem();
	if (!item.IsNull() && item.GetData())
	{
		try
		{
			MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(m_doc->m_body.Document())->createRange();
			MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElement *)(item.GetData());
			range->selectNodeContents(elem);
			range->collapse(VARIANT_TRUE);

			MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_doc->m_body.Document())->getSelection();
			sel->removeAllRanges();
			sel->addRange(range);
		}
		catch (_com_error &e)
		{
			U::ReportError(e);
		}
		/*	MSHTML::IHTMLBodyElementPtr body = (MSHTML::IHTMLBodyElementPtr)m_doc->m_body.Document()->body;
			MSHTML::IHTMLTxtRangePtr rng = body->createTextRange();
			MSHTML::IHTMLElement *elem = (MSHTML::IHTMLElement *)item.GetData();
			rng->moveToElementText(elem);
			rng->select();*/
		ShowView(SOURCE);
	}

	return S_OK;
}

LRESULT CMainFrame::OnTreeDeleteElement(WORD, WORD, HWND, BOOL &)
{
	CString message;
	message.LoadString(ID_DT_DELETE);
	message += L"?";

	if (AtlTaskDialog(::GetActiveWindow(), IDS_DOCUMENT_TREE_CAPTION, (LPCTSTR)message, (LPCTSTR)NULL, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, TD_WARNING_ICON) == IDYES)
	{
		CTreeItem item = m_document_tree.m_tree.m_tree.GetLastSelectedItem();
		m_doc->m_body.BeginUndoUnit(L"delete elements");
		do
		{
			if (!item.IsNull() && item.GetData())
			{
				MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElement *)item.GetData();
				if (!elem)
					return S_OK;

				MSHTML::IHTMLDOMNodePtr(elem)->removeNode(VARIANT_TRUE);
			}
			else
				break;

			item = m_document_tree.m_tree.m_tree.GetPrevSelectedItem(item);
		} while (!item.IsNull());
		m_doc->m_body.EndUndoUnit();
	}
	return S_OK;
}

LRESULT CMainFrame::OnTreeMerge(WORD, WORD, HWND, BOOL &)
{
	CTreeItem item = m_document_tree.GetSelectedItem();
	if (item.IsNull())
		return S_OK;

	MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElement *)item.GetData();
	if (!elem)
		return S_OK;

	bool merged = m_doc->m_body.bCall(L"MergeContainers", elem);
	m_doc->m_body.Call(L"MergeContainers", elem);
	// Move cursor to selected element
	if (merged)
		GoTo(elem);

	return S_OK;
}

LRESULT CMainFrame::OnTreeClick(WORD, WORD, HWND hWndCtl, BOOL &)
{
	GoToSelectedTreeItem();
	return S_OK;
}

// binary objects
LRESULT CMainFrame::OnEditAddBinary(WORD, WORD, HWND, BOOL &)
{
	if (!m_doc)
		return S_OK;

	// Modification by Pilgrim
	CMultiFileDialog dlg(
		_T("*"),
		NULL,
		OFN_ALLOWMULTISELECT | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR,
		L"All files (*.*)\0*.*\0FBE supported (*.jpg;*.jpeg;*.png)\0*.jpg;*.jpeg;*.png\0JPEG (*.jpg)\0*.jpg\0PNG (*.png)"
		L"\0*.png\0Bitmap (*.bmp)\0*.bmp\0GIF (*.gif)\0*.gif\0TIFF (*.tif)\0*.tif\0\0");
	wchar_t dlgTitle[MAX_LOAD_STRING + 1];
	::LoadString(_Module.GetResourceInstance(), IDS_ADD_BINARIES_FILEDLG, dlgTitle, MAX_LOAD_STRING);
	dlg.m_ofn.lpstrTitle = dlgTitle;
	dlg.m_ofn.nFilterIndex = 2;
	dlg.m_ofn.Flags &= ~OFN_ENABLEHOOK;
	dlg.m_ofn.lpfnHook = NULL;

	if (dlg.DoModal(*this) == IDOK)
	{
		CString strPath;
		if (dlg.GetFirstPathName(strPath))
		{
			if (strPath.IsEmpty())
				return S_OK;
			do
			{
				m_doc->m_body.AttachImage(strPath);
			} while (dlg.GetNextPathName(strPath));
		}
	}

	return S_OK;
}

// incremental search
LRESULT CMainFrame::OnEditIncSearch(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == SOURCE)
		return S_OK;

	if (m_incsearch == 0)
	{
		ShowView(BODY);
		m_doc->m_body.StartIncSearch();
		m_is_str.Empty();
		m_is_prev = m_doc->m_body.LastSearchPattern();
		m_incsearch = 1;
		m_is_fail = false;
		SetIsText();
	}
	else if (m_incsearch == 1 && m_is_str.IsEmpty() && !m_is_prev.IsEmpty())
	{
		m_incsearch = 2;
		m_is_str.Empty();
		for (int i = 0; i < m_is_prev.GetLength(); ++i)
			PostMessage(WM_CHAR, m_is_prev[i], 0x20000000);
	}
	else if (!m_is_fail)
		m_doc->m_body.DoIncSearch(m_is_str, true);
	return S_OK;
}

LRESULT CMainFrame::OnChar(UINT, WPARAM wParam, LPARAM lParam, BOOL &)
{
	if (!m_incsearch)
		return S_OK;
	// only a few keys are supported
	if (wParam == 8)
	{ // backspace
		if (!m_is_str.IsEmpty())
			m_is_str.Delete(m_is_str.GetLength() - 1);
		if (!m_doc->m_body.DoIncSearch(m_is_str, false))
		{
			m_is_fail = true;
			::MessageBeep(MB_ICONEXCLAMATION);
		}
		else
			m_is_fail = false;
	}
	else if (wParam == 13)
	{ // enter
		StopIncSearch(false);
		return S_OK;
	}
	else if (wParam >= 32 && wParam != 127)
	{ // printable char
		if (m_is_fail)
		{
			::MessageBeep(MB_ICONEXCLAMATION);
			if (!(lParam & 0x20000000))
				return S_OK;
		}
		m_is_str += (TCHAR)wParam;
		if (!m_doc->m_body.DoIncSearch(m_is_str, false))
		{
			if (!m_is_fail)
				::MessageBeep(MB_ICONEXCLAMATION);
			m_is_fail = true;
		}
		else
			m_is_fail = false;
	}
	SetIsText();
	return S_OK;
}

/// <summary>Transfer Source content to HTML</summary>
/// <returns>true on SUCCESS</returns>
bool CMainFrame::SourceToHTML()
{
	// get text from Scintilla and convert it to UTF-16
	int textlen = 0;
	char *buffer = nullptr;
	textlen = m_source.SendMessage(SCI_GETLENGTH);
	buffer = new char[textlen + 1];
	m_source.SendMessage(SCI_GETTEXT, textlen + 1, (LPARAM)buffer);

	DWORD ulen = ::MultiByteToWideChar(CP_UTF8, 0, buffer, textlen, NULL, 0);
	BSTR ustr = ::SysAllocStringLen(NULL, ulen);

	::MultiByteToWideChar(CP_UTF8, 0, buffer, textlen, ustr, ulen);

	//	get selection position
	int selectedPosBegin = m_source.SendMessage(SCI_GETSELECTIONSTART);
	int selectedPosEnd = m_source.SendMessage(SCI_GETSELECTIONEND);
	if (selectedPosEnd == selectedPosBegin)
	{
		selectedPosEnd = selectedPosBegin = MultiByteToWideChar(CP_UTF8, 0, buffer, selectedPosBegin, NULL, 0);
	}
	else
	{
		selectedPosBegin = MultiByteToWideChar(CP_UTF8, 0, buffer, selectedPosBegin, NULL, 0);
		selectedPosEnd = MultiByteToWideChar(CP_UTF8, 0, buffer, selectedPosEnd, NULL, 0);
	}
	delete buffer;

	// Rebuild html if text changed
	LRESULT isTextChanged = m_source.SendMessage(SCI_GETMODIFY);
	if (isTextChanged)
	{
		// check document is well formed (not validate!)
		int err_col = 0;
		int err_line = 0;
		CString ErrMsg(L"");
		if (!U::IsXMLWellFormed(ustr, err_col, err_line, ErrMsg))
		{
			::SysFreeString(ustr);
			DisplayXMLError(ErrMsg);
			SourceGoTo(err_col, err_line);
			return false;
		}
		if (!m_doc->LoadXMLFromString(ustr))
		{
			::SysFreeString(ustr);
			return false;
		}
		if (!m_doc->BuildHTML())
		{
			::SysFreeString(ustr);
			return false;
		}
	}

	// transfer selection
	if (m_doc->SetHTMLSelectionFromXML(selectedPosBegin, selectedPosEnd, ustr))
	{
		::SysFreeString(ustr);
	}
	else
	{
		::SysFreeString(ustr);
		return false;
	}

	if (_Settings.ViewDocumentTree())
	{
		GetDocumentStructure();
		m_document_tree.HighlightItemAtPos(m_doc->m_body.SelectionContainer());
	}
	m_doc->m_body.Synched(true);
	m_doc->m_body.DescFormChanged(false);
	m_source.SendMessage(SCI_SETSAVEPOINT);

	return true;
}
/// <summary>Transfer HTML to SOURCE</summary>
/// <returns>true on SUCCESS</returns>
bool CMainFrame::HTMLToSource()
{
	int savedPosBegin = 0;
	int savedPosEnd = 0;
	_bstr_t src;

	if (!m_doc->GetXMLSelectionFromHTML(savedPosBegin, savedPosEnd, src))
		return false;

	// recalculate position for UTF-8
	savedPosBegin = ::WideCharToMultiByte(CP_UTF8, 0, src, savedPosBegin, NULL, 0, NULL, NULL);
	savedPosEnd = ::WideCharToMultiByte(CP_UTF8, 0, src, savedPosEnd, NULL, 0, NULL, NULL);

	//Put text to Scintilla editor
	if (!m_doc->m_body.IsSynched())
	{
		// Convert to UTF-8
		DWORD nch = ::WideCharToMultiByte(CP_UTF8, 0, src, src.length(), NULL, 0, NULL, NULL);
		m_source.SendMessage(SCI_CLEARALL);
		char *buffer = (char *)malloc(nch);
		if (buffer)
		{
			::WideCharToMultiByte(CP_UTF8, 0, src, src.length(), buffer, nch, NULL, NULL);
			m_source.SendMessage(SCI_APPENDTEXT, nch, (LPARAM)buffer);
			free(buffer);
		}
		m_source.SendMessage(SCI_EMPTYUNDOBUFFER);
		m_sci_painted = false; // set flag to wait SCN_PAINTED event
	}
	//	set selection range
	m_source.SendMessage(SCI_SETSEL, savedPosBegin, savedPosEnd);
	m_doc->m_body.Synched(true);
	return true;
}

/// <summary>Show different views<summary>
void CMainFrame::ShowView(VIEW_TYPE vt)
{
	// if same view - return
	if (m_current_view == vt)
		return;

	// Define next view if using Ctr+Tab
	if (vt == NEXT)
	{
		if (!m_ctrl_tab)
		{
			if (m_current_view != m_last_ctrl_tab_view)
				vt = m_last_ctrl_tab_view;
			else
			{
				if ((m_last_view == BODY && m_current_view == DESC) ||
					(m_last_view == DESC && m_current_view == BODY))
					vt = SOURCE;
				if ((m_last_view == BODY && m_current_view == SOURCE) ||
					(m_last_view == SOURCE && m_current_view == BODY))
					vt = DESC;
				if ((m_last_view == SOURCE && m_current_view == DESC) ||
					(m_last_view == DESC && m_current_view == SOURCE))
					vt = BODY;
			}
			m_last_ctrl_tab_view = m_current_view;
			m_ctrl_tab = true;
		}
		else
		{
			if ((m_last_view == BODY && m_current_view == DESC) ||
				(m_last_view == DESC && m_current_view == BODY))
				vt = SOURCE;
			if ((m_last_view == BODY && m_current_view == SOURCE) ||
				(m_last_view == SOURCE && m_current_view == BODY))
				vt = DESC;
			if ((m_last_view == SOURCE && m_current_view == DESC) ||
				(m_last_view == DESC && m_current_view == SOURCE))
				vt = BODY;
		}
	}

	// if view changed, close all dialogs
	if (m_current_view != vt)
	{
		m_doc->m_body.CloseFindDialog(m_doc->m_body.m_find_dlg);
		m_doc->m_body.CloseFindDialog(m_doc->m_body.m_replace_dlg);
		m_doc->m_body.CloseFindDialog(m_sci_find_dlg);
		m_doc->m_body.CloseFindDialog(m_sci_replace_dlg);
	}

	if (!m_ctrl_tab && m_current_view != vt)
	{
		m_last_ctrl_tab_view = m_current_view;
	}

	WORD body_only_commands[] =
	{
		ID_STYLE_NORMAL,
		ID_ADD_SUBTITLE,
		ID_ADD_TEXTAUTHOR,
		ID_STYLE_CODE,
		ID_STYLE_LINK,
		ID_STYLE_NOTE,
		ID_STYLE_NOLINK,
		ID_STYLE_BOLD,
		ID_STYLE_ITALIC,
		ID_STYLE_SUPERSCRIPT,
		ID_STYLE_SUBSCRIPT,
		ID_STYLE_STRIKETROUGH,
		ID_EDIT_CLONE,
		ID_EDIT_SPLIT,
		ID_EDIT_MERGE,
		ID_EDIT_REMOVE_OUTER,
		ID_GOTO_FOOTNOTE,
		ID_ADD_BODY,
		ID_ADD_BODY_NOTES,
		ID_ADD_ANN,
		ID_ADD_SECTION,
		ID_ADD_TITLE,
		ID_ADD_EPIGRAPH,
		ID_ADD_POEM,
		ID_ADD_STANZA,
		ID_ADD_CITE,
		ID_ADD_IMAGE,
		ID_INS_TABLE,
		ID_ATTACH_BINARY,
		ID_TOOLS_WORDS,
		ID_TOOLS_SPELLCHECK,
		ID_LAST_SCRIPT
	};
	// if move not to BODY, close spell dialogs
	if (m_current_view == BODY)
	{
		// disable non BODY commands
		for (int i = 0; i < sizeof(body_only_commands) / sizeof(body_only_commands[0]); ++i)
			UIEnable(body_only_commands[i], FALSE);

		// disable scripts menu
		//get menu handle
		HMENU scripts = GetSubMenu(m_MenuBar.GetMenu(), SCRIPT_MENU_ORDER);

		for (int i = 0; i < m_scripts.GetSize(); ++i)
		{
			if (!m_scripts[i].isFolder)
			{
				::EnableMenuItem(scripts, ID_SCRIPT_BASE + m_scripts[i].wID, MF_BYCOMMAND | MF_GRAYED);
			}
		}

		// clear & disable toolbar elements
		m_href_box.SetWindowText(_T(""));
		m_href_box.EnableWindow(FALSE);

		m_id_box.SetWindowText(_T(""));
		m_id_box.EnableWindow(FALSE);

		m_image_title_box.SetWindowText(_T(""));
		m_image_title_box.EnableWindow(FALSE);

		m_section_box.SetWindowText(_T(""));
		m_section_box.EnableWindow(FALSE);

		m_id_table_id_box.SetWindowText(_T(""));
		m_id_table_id_box.EnableWindow(FALSE);

		m_id_table_box.SetWindowText(_T(""));
		m_id_table_box.EnableWindow(FALSE);

		m_styleT_table_box.SetWindowText(_T(""));
		m_styleT_table_box.EnableWindow(FALSE);

		m_style_table_box.SetWindowText(_T(""));
		m_style_table_box.EnableWindow(FALSE);

		m_colspan_table_box.SetWindowText(_T(""));
		m_colspan_table_box.EnableWindow(FALSE);

		m_rowspan_table_box.SetWindowText(_T(""));
		m_rowspan_table_box.EnableWindow(FALSE);

		m_alignTR_table_box.SetWindowText(_T(""));
		m_alignTR_table_box.EnableWindow(FALSE);

		m_align_table_box.SetWindowText(_T(""));
		m_align_table_box.EnableWindow(FALSE);

		m_valign_table_box.SetWindowText(_T(""));
		m_valign_table_box.EnableWindow(FALSE);

		m_id_caption.SetEnabled(false);
		m_href_caption.SetEnabled(false);
		m_section_id_caption.SetEnabled(false);
		m_image_title_caption.SetEnabled(false);
		m_table_id_caption.SetEnabled(false);
		m_table_style_caption.SetEnabled(false);
		m_id_table_caption.SetEnabled(false);
		m_style_caption.SetEnabled(false);
		m_colspan_caption.SetEnabled(false);
		m_rowspan_caption.SetEnabled(false);
		m_tr_allign_caption.SetEnabled(false);
		m_th_allign_caption.SetEnabled(false);
		m_valign_caption.SetEnabled(false);
	}

	// Move from Source view
	if (m_current_view == SOURCE)
	{
		if (SourceToHTML())
		{
			// Show Tree & splitter
			UIEnable(ID_VIEW_TREE, 1);
			/*m_save_sp_mode=true;// Modification by Pilgrim - иначе только на ХР(!)при выборе DESC слитает ID_VIEW_TREE и переход на BODY не восстанавливает. Но, если после запуска сразу перейти на SOURCE, то переходы на DESC и BODY не сносят ID_VIEW_TREE. Надо разобраться, а потом удалить m_save_sp_mode=true;
			UISetCheck(ID_VIEW_TREE, m_save_sp_mode);*/
			m_splitter.SetSinglePaneMode(_Settings.ViewDocumentTree() ? SPLIT_PANE_NONE : SPLIT_PANE_RIGHT);

			// disable Source view commands
			UIEnable(ID_GOTO_MATCHTAG, FALSE);
			m_doc->RestoreHTMLSelection();
		}
		else
			return;
	}

	// Clear checked on all buttons
	UISetCheck(ID_VIEW_BODY, 0);
	UISetCheck(ID_VIEW_DESC, 0);
	UISetCheck(ID_VIEW_SOURCE, 0);
	MSHTML::IHTMLSelectElementPtr elem = nullptr;


	if (vt == BODY)
	{
		UISetCheck(ID_VIEW_BODY, 1); //set view button pressed
		m_status.SetPaneText(ID_PANE_INS, m_last_ie_ovr ? strOVR : strINS);

		for (int i = 0; i < sizeof(body_only_commands) / sizeof(body_only_commands[0]); ++i)
			UIEnable(body_only_commands[i], TRUE);

		UIEnable(ID_VIEW_TREE, TRUE);
		UIEnable(ID_EDIT_FIND, TRUE);
		UIEnable(ID_EDIT_REPLACE, TRUE);
		UIEnable(ID_EDIT_FINDNEXT, TRUE);


		// disable Source view commands
		UIEnable(ID_GOTO_MATCHTAG, FALSE);

		// enable scripts menu
		HMENU scripts = GetSubMenu(m_MenuBar.GetMenu(), SCRIPT_MENU_ORDER);
		for (int i = 0; i < m_scripts.GetSize(); ++i)
		{
			if (!m_scripts[i].isFolder)
			{
				::EnableMenuItem(scripts, ID_SCRIPT_BASE + m_scripts[i].wID, MF_BYCOMMAND | MF_ENABLED);
			}
		}

		m_sel_changed = true;

		//hide description, show body
		m_doc->m_body.TraceChanges(false);
		{
			CComDispatchDriver body(m_doc->m_body.Script());
			CComVariant args[1];
			args[0] = false;
			CheckError(body.Invoke1(L"apiShowDesc", &args[0]));
		}
		m_doc->m_body.TraceChanges(true);

		// detect document language (based on FB2 document settings)
		elem = MSHTML::IHTMLDocument3Ptr(m_doc->m_body.Document())->getElementById(L"tiLang");
		if (elem)
			m_Speller->SetDocumentLanguage((LPWSTR)elem->value);
		m_view.ActivateWnd(m_doc->m_body);
		m_doc->m_body.ScrollToSelection();
		m_doc->m_body.SetFocus();
	}
	if (vt == DESC)
	{
		UISetCheck(ID_VIEW_DESC, 1);
		m_view.ActivateWnd(m_doc->m_body);
		m_status.SetPaneText(ID_DEFAULT_PANE, _T(""));
		UIEnable(ID_VIEW_TREE, TRUE);
		UIEnable(ID_EDIT_FIND, FALSE);
		UIEnable(ID_EDIT_REPLACE, FALSE);
		UIEnable(ID_EDIT_FINDNEXT, FALSE);

		//show description, hide body
		m_doc->m_body.TraceChanges(false);
		{
			CComDispatchDriver body(m_doc->m_body.Script());
			CComVariant arg = true;
			body.Invoke1(L"apiShowDesc", &arg);
		}
		m_doc->m_body.SetFocus();
		m_doc->m_body.TraceChanges(true);
	}
	if (vt == SOURCE)
	{
		//important for jump from DESC view to raise focusOut event
		m_source.SetFocus();
		m_doc->m_body.TraceChanges(false);
		if (!HTMLToSource())
		{
			m_doc->m_body.TraceChanges(true);
			return;
		}
		/*		// added by SeNS: display line numbers
		if (_Settings.m_show_line_numbers)
			m_source.SendMessage(SCI_SETMARGINWIDTHN, 0, 64);
		else
			m_source.SendMessage(SCI_SETMARGINWIDTHN, 0, 0);
			*/
		UISetCheck(ID_VIEW_SOURCE, 1);
		
		// enable Source view commands
		UIEnable(ID_GOTO_MATCHTAG, TRUE);
		UIEnable(ID_VIEW_TREE, FALSE);
		UIEnable(ID_EDIT_FIND, TRUE);
		UIEnable(ID_EDIT_REPLACE, TRUE);
		UIEnable(ID_EDIT_FINDNEXT, TRUE);

		m_status.SetPaneText(ID_PANE_CHAR, L"");
		m_status.SetPaneText(ID_PANE_INS, m_last_sci_ovr ? strOVR : strINS);
		m_view.ActivateWnd(m_source);

		// wait for complete Scintilla view to scroll carret
		MSG msg;
		while (!m_sci_painted && ::GetMessage(&msg, NULL, 0, 0))
		{
			::TranslateMessage(&msg);
			::DispatchMessage(&msg);
		}
		m_splitter.SetSinglePaneMode(SPLIT_PANE_RIGHT);
		m_source.SendMessage(SCI_SCROLLCARET);
		m_doc->m_body.TraceChanges(true);
		/*if (m_current_view == BODY)
		{
			CComDispatchDriver body(m_doc->m_body.Script());
			CheckError(body.InvokeN(L"SaveBodyScroll", 0, 0));
		}*/
	}

	m_last_view = m_current_view;
	m_current_view = vt;
}

LRESULT CMainFrame::OnGoToFootnote(WORD wNotifyCode, WORD wID, HWND hWndCtl)
{
	if (!m_doc->m_body.GoToFootnote(false))
		m_doc->m_body.GoToReference(false);
	return S_OK;
}

LRESULT CMainFrame::OnGoToReference(WORD wNotifyCode, WORD wID, HWND hWndCtl)
{
	m_doc->m_body.GoToReference(false);
	return S_OK;
}

/// <summary>Setup Scintilla acccording to settings<summary>
void CMainFrame::SetupSci()
{
	m_source.SendMessage(SCI_SETCODEPAGE, SC_CP_UTF8);
	m_source.SendMessage(SCI_SETEOLMODE, SC_EOL_CRLF);
	m_source.SendMessage(SCI_SETVIEWEOL, _Settings.m_xml_src_showEOL);
	m_source.SendMessage(SCI_SETVIEWWS, _Settings.m_xml_src_showSpace);
	m_source.SendMessage(SCI_SETWRAPMODE, _Settings.m_xml_src_wrap ? SC_WRAP_WORD : SC_WRAP_NONE);

	// added by SeNS: try to speed-up wrap mode
	m_source.SendMessage(SCI_SETLAYOUTCACHE, SC_CACHE_DOCUMENT);
	m_source.SendMessage(SCI_SETXCARETPOLICY, CARET_SLOP | CARET_EVEN, 50);
	m_source.SendMessage(SCI_SETYCARETPOLICY, CARET_SLOP | CARET_EVEN, 50);

	// added by SeNS: display line numbers
	if (_Settings.m_show_line_numbers)
	{
		int margin_line_w = m_source.SendMessage(SCI_TEXTWIDTH, STYLE_LINENUMBER, (WPARAM) "_99999");
		m_source.SendMessage(SCI_SETMARGINWIDTHN, 0, margin_line_w);
	}
	else
	{
		m_source.SendMessage(SCI_SETMARGINWIDTHN, 0, 0);
	}

	// added by SeNS: disable Scintilla's control characters
	char sciCtrlChars[] = { 'Q', 'E', 'R', 'S', 'K', ':' };
	for (int i = 0; i < sizeof(sciCtrlChars); i++)
		m_source.SendMessage(SCI_ASSIGNCMDKEY, sciCtrlChars[i] + (SCMOD_CTRL << 16), SCI_NULL);

	char sciCtrlShiftChars[] = { 'Q', 'W', 'E', 'R', 'Y', 'O', 'P', 'A', 'S', 'D', 'F', 'G', 'H', 'K', 'Z', 'X', 'C', 'V', 'B', 'N', ':' };
	for (int i = 0; i < sizeof(sciCtrlShiftChars); i++)
		m_source.SendMessage(SCI_ASSIGNCMDKEY, sciCtrlShiftChars[i] + ((SCMOD_CTRL + SCMOD_SHIFT) << 16), SCI_NULL);
	///
	if (_Settings.m_xml_src_syntaxHL)
	{
		m_source.SendMessage(SCI_SETLEXER, SCLEX_XML);
		m_source.SendMessage(SCI_SETMARGINTYPEN, 2, SC_MARGIN_SYMBOL);
		m_source.SendMessage(SCI_SETMARGINWIDTHN, 2, 16);
		m_source.SendMessage(SCI_SETMARGINMASKN, 2, SC_MASK_FOLDERS);
		m_source.SendMessage(SCI_SETMARGINSENSITIVEN, 2, 1);

		// Define markers
		DefineMarker(SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));
		DefineMarker(SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER, RGB(0xff, 0xff, 0xff), RGB(0, 0, 0));

		// indicator for tag match
		m_source.SendMessage(SCI_INDICSETSTYLE, SCE_UNIVERSAL_TAGMATCH, INDIC_ROUNDBOX);
		m_source.SendMessage(SCI_INDICSETALPHA, SCE_UNIVERSAL_TAGMATCH, 100);
		m_source.SendMessage(SCI_INDICSETUNDER, SCE_UNIVERSAL_TAGMATCH, TRUE);
		m_source.SendMessage(SCI_INDICSETFORE, SCE_UNIVERSAL_TAGMATCH, RGB(128, 128, 255));

		m_source.SendMessage(SCI_INDICSETSTYLE, SCE_UNIVERSAL_TAGATTR, INDIC_ROUNDBOX);
		m_source.SendMessage(SCI_INDICSETALPHA, SCE_UNIVERSAL_TAGATTR, 100);
		m_source.SendMessage(SCI_INDICSETUNDER, SCE_UNIVERSAL_TAGATTR, TRUE);
		m_source.SendMessage(SCI_INDICSETFORE, SCE_UNIVERSAL_TAGATTR, RGB(128, 128, 255));

		// turn on folding
		m_source.SendMessage(SCI_SETFOLDFLAGS, SC_FOLDFLAG_LINEAFTER_CONTRACTED);
		m_source.SendMessage(SCI_SETPROPERTY, (WPARAM) "fold", (WPARAM) "1");
		m_source.SendMessage(SCI_SETPROPERTY, (WPARAM) "fold.html", (WPARAM) "1");
		m_source.SendMessage(SCI_SETPROPERTY, (WPARAM) "fold.compact", (WPARAM) "1");
		m_source.SendMessage(SCI_SETPROPERTY, (WPARAM) "fold.flags", (WPARAM) "16");

		m_source.SendMessage(SCI_COLOURISE, 0, -1);
	}
	else
	{
		m_source.SendMessage(SCI_SETLEXER, SCLEX_NULL);
		m_source.SendMessage(SCI_SETMARGINWIDTHN, 2, 0);
		FoldAll();
		m_source.SendMessage(SCI_COLOURISE, 0, -1);
	}
}

/// <summary>Setup Scintilla styles acccording to settings<summary>
void CMainFrame::SetSciStyles()
{
	m_source.SendMessage(SCI_STYLERESETDEFAULT);

	/// Set source font
	CT2A srcFont(_Settings.GetSrcFont());
	m_source.SendMessage(SCI_STYLESETFONT, STYLE_DEFAULT, (LPARAM)srcFont.m_psz);
	m_source.SendMessage(SCI_STYLESETSIZE, STYLE_DEFAULT, _Settings.GetFontSize());

	/*  DWORD fs = _Settings.GetColorFG();
	  if (fs!=CLR_DEFAULT)
		m_source.SendMessage(SCI_STYLESETFORE,STYLE_DEFAULT,fs);

	  fs = _Settings.GetColorBG();
	  if (fs!=CLR_DEFAULT)
		m_source.SendMessage(SCI_STYLESETBACK,STYLE_DEFAULT,fs);*/

	m_source.SendMessage(SCI_STYLECLEARALL);

	// set XML specific styles
	static struct
	{
		char style;
		int color;
	} styles[] = {
		{0, RGB(0, 0, 0)},		// default text
		{1, RGB(128, 0, 0)},	// tags
		{2, RGB(128, 0, 0)},	// unknown tags
		{3, RGB(128, 128, 0)},  // attributes
		{4, RGB(255, 0, 0)},	// unknown attributes
		{5, RGB(0, 128, 96)},   // numbers
		{6, RGB(0, 128, 0)},	// double quoted strings
		{7, RGB(0, 128, 0)},	// single quoted strings
		{8, RGB(128, 0, 128)},  // other inside tag
		{9, RGB(0, 128, 128)},  // comments
		{10, RGB(128, 0, 128)}, // entities
		{11, RGB(128, 0, 0)},   // tag ends
		{12, RGB(128, 0, 128)}, // xml decl start
		{13, RGB(128, 0, 128)}, // xml decl end
		{17, RGB(128, 0, 0)},   // cdata
		{18, RGB(128, 0, 0)},   // question
		{19, RGB(96, 128, 96)}, // unquoted value
	};
	if (_Settings.m_xml_src_syntaxHL)
		for (int i = 0; i < sizeof(styles) / sizeof(styles[0]); ++i)
			m_source.SendMessage(SCI_STYLESETFORE, styles[i].style, styles[i].color);
}

// zoom in
LRESULT CMainFrame::OnZoomIn(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == BODY || m_current_view == DESC)
	{
		SHD::IWebBrowser2Ptr brws = m_doc->m_body.Browser();
		// get current zoom value
		CComVariant ret;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, NULL, &ret);

		// set new zoom value
		CComVariant zoom_value = ret.intVal <= 975 ? ret.intVal + 25 : 1000;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, &zoom_value, NULL);
	}
	else
	{
		m_source.SendMessage(SCI_ZOOMIN);
	}

	return 0;
}

// zoom out
LRESULT CMainFrame::OnZoomOut(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == BODY || m_current_view == DESC)
	{
		SHD::IWebBrowser2Ptr brws = m_doc->m_body.Browser();
		// get current zoom value
		CComVariant ret;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, NULL, &ret);
		// set new zoom value
		CComVariant zoom_value = ret.intVal >= 50 ? ret.intVal - 25 : 25;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, &zoom_value, NULL);
	}
	else 
	{
		m_source.SendMessage(SCI_ZOOMOUT);
	}

	return 0;
}

// reset zoom to 100%
LRESULT CMainFrame::OnZoom100(WORD, WORD, HWND, BOOL &)
{
	if (m_current_view == BODY || m_current_view == DESC)
	{
		SHD::IWebBrowser2Ptr brws = m_doc->m_body.Browser();
		// set new zoom value
		CComVariant zoom_value = 100;
		brws->ExecWB(SHD::OLECMDID_OPTICAL_ZOOM, SHD::OLECMDEXECOPT_DONTPROMPTUSER, &zoom_value, NULL);
	}
	else
	{
		m_source.SendMessage(SCI_SETZOOM, 0);
	}

	return 0;
}


LRESULT CMainFrame::OnSciModified(int id, NMHDR *hdr, BOOL &bHandled)
{
	if (hdr->hwndFrom != m_source)
	{
		bHandled = FALSE;
		return S_OK;
	}
	SciModified(*(SCNotification *)hdr);
	return S_OK;
}

LRESULT CMainFrame::OnSciMarginClick(int id, NMHDR *hdr, BOOL &bHandled)
{
	if (hdr->hwndFrom != m_source)
	{
		bHandled = FALSE;
		return S_OK;
	}
	SciMarginClicked(*(SCNotification *)hdr);
	return S_OK;
}

LRESULT CMainFrame::OnSciUpdateUI(int id, NMHDR *hdr, BOOL &bHandled)
{
	if (m_current_view == SOURCE)
		SciUpdateUI(false);
	return S_OK;
}

//set flag for right scrolling
LRESULT CMainFrame::OnSciPainted(int id, NMHDR *hdr, BOOL &bHandled)
{
	m_sci_painted = true;
	return S_OK;
}

LRESULT CMainFrame::OnSciCollapse(WORD cose, WORD wID, HWND, BOOL &)
{
	if (m_current_view == SOURCE)
		SciCollapse(wID - ID_SCI_COLLAPSE_BASE, false);

	if (m_document_tree.IsWindowVisible())
		m_document_tree.m_tree.m_tree.Collapse(0, wID - ID_SCI_COLLAPSE_BASE, false);

	return S_OK;
}

LRESULT CMainFrame::OnSciExpand(WORD cose, WORD wID, HWND, BOOL &)
{
	if (m_current_view == SOURCE)
		SciCollapse(wID - ID_SCI_EXPAND_BASE, true);

	if (m_document_tree.IsWindowVisible())
		m_document_tree.m_tree.m_tree.Collapse(0, wID - ID_SCI_EXPAND_BASE, true);

	return S_OK;
}

LRESULT CMainFrame::OnGoToMatchTag(WORD wNotifyCode, WORD wID, HWND hWndCtl)
{
	if (m_current_view == SOURCE)
		SciUpdateUI(true);
	return S_OK;
}

LRESULT CMainFrame::OnGoToWrongTag(WORD wNotifyCode, WORD wID, HWND hWndCtl)
{
	//	if (m_current_view == SOURCE)
	//		SciGotoWrongTag();
	return S_OK;
}

void CMainFrame::FoldAll()
{
	m_source.SendMessage(SCI_COLOURISE, 0, -1);
	int maxLine = m_source.SendMessage(SCI_GETLINECOUNT);
	bool expanding = true;
	for (int lineSeek = 0; lineSeek < maxLine; lineSeek++)
	{
		if (m_source.SendMessage(SCI_GETFOLDLEVEL, lineSeek) & SC_FOLDLEVELHEADERFLAG)
		{
			expanding = !m_source.SendMessage(SCI_GETFOLDEXPANDED, lineSeek);
			break;
		}
	}
	for (int line = 0; line < maxLine; line++)
	{
		int level = m_source.SendMessage(SCI_GETFOLDLEVEL, line);
		if ((level & SC_FOLDLEVELHEADERFLAG) &&
			(SC_FOLDLEVELBASE == (level & SC_FOLDLEVELNUMBERMASK)))
		{
			if (expanding)
			{
				m_source.SendMessage(SCI_SETFOLDEXPANDED, line, 1);
				ExpandFold(line, true, false, 0, level);
				line--;
			}
			else
			{
				int lineMaxSubord = m_source.SendMessage(SCI_GETLASTCHILD, line, -1);
				m_source.SendMessage(SCI_SETFOLDEXPANDED, line, 0);
				if (lineMaxSubord > line)
					m_source.SendMessage(SCI_HIDELINES, line + 1, lineMaxSubord);
			}
		}
	}
}

void CMainFrame::ExpandFold(int &line, bool doExpand, bool force, int visLevels, int level)
{
	int lineMaxSubord = m_source.SendMessage(SCI_GETLASTCHILD, line, level & SC_FOLDLEVELNUMBERMASK);
	line++;
	while (line <= lineMaxSubord)
	{
		if (force)
		{
			if (visLevels > 0)
				m_source.SendMessage(SCI_SHOWLINES, line, line);
			else
				m_source.SendMessage(SCI_HIDELINES, line, line);
		}
		else
		{
			if (doExpand)
				m_source.SendMessage(SCI_SHOWLINES, line, line);
		}
		int levelLine = level;
		if (levelLine == -1)
			levelLine = m_source.SendMessage(SCI_GETFOLDLEVEL, line);
		if (levelLine & SC_FOLDLEVELHEADERFLAG)
		{
			if (force)
			{
				if (visLevels > 1)
					m_source.SendMessage(SCI_SETFOLDEXPANDED, line, 1);
				else
					m_source.SendMessage(SCI_SETFOLDEXPANDED, line, 0);
				ExpandFold(line, doExpand, force, visLevels - 1);
			}
			else
			{
				if (doExpand)
				{
					if (!m_source.SendMessage(SCI_GETFOLDEXPANDED, line))
						m_source.SendMessage(SCI_SETFOLDEXPANDED, line, 1);
					ExpandFold(line, true, force, visLevels - 1);
				}
				else
				{
					ExpandFold(line, false, force, visLevels - 1);
				}
			}
		}
		else
		{
			line++;
		}
	}
}

void CMainFrame::SciCollapse(int level2Collapse, bool mode)
{
	m_source.SendMessage(SCI_COLOURISE, 0, -1);
	int maxLine = m_source.SendMessage(SCI_GETLINECOUNT);

	for (int line = 0; line < maxLine; line++)
	{
		int level = m_source.SendMessage(SCI_GETFOLDLEVEL, line);
		if (level & SC_FOLDLEVELHEADERFLAG)
		{
			level -= SC_FOLDLEVELBASE;
			if (level2Collapse == (level & SC_FOLDLEVELNUMBERMASK))
				if ((m_source.SendMessage(SCI_GETFOLDEXPANDED, line) != 0) != mode)
					m_source.SendMessage(SCI_TOGGLEFOLD, line);
		}
	}
}

/// <summary>Internal function for defining Scintilla marker<summary>
void CMainFrame::DefineMarker(int marker, int markerType, COLORREF fore, COLORREF back)
{
	m_source.SendMessage(SCI_MARKERDEFINE, marker, markerType);
	m_source.SendMessage(SCI_MARKERSETFORE, marker, fore);
	m_source.SendMessage(SCI_MARKERSETBACK, marker, back);
}

void CMainFrame::SciModified(const SCNotification &scn)
{
	if (scn.modificationType & SC_MOD_CHANGEFOLD)
	{
		if (scn.foldLevelNow & SC_FOLDLEVELHEADERFLAG)
		{
			if (!(scn.foldLevelPrev & SC_FOLDLEVELHEADERFLAG))
				m_source.SendMessage(SCI_SETFOLDEXPANDED, scn.line, 1);
		}
		else if (scn.foldLevelPrev & SC_FOLDLEVELHEADERFLAG)
		{
			if (!m_source.SendMessage(SCI_GETFOLDEXPANDED, scn.line))
			{
				// Removing the fold from one that has been contracted so should expand
				// otherwise lines are left invisible with no way to make them visible
				int tmpline = scn.line;
				ExpandFold(tmpline, true, false, 0, scn.foldLevelPrev);
			}
		}
	}
}

bool CMainFrame::SciUpdateUI(bool gotoTag)
{
	if (_Settings.m_xml_src_tagHL || gotoTag)
	{
		XmlMatchedTagsHighlighter xmlTagMatchHiliter(&m_source);
		UIEnable(ID_GOTO_MATCHTAG, xmlTagMatchHiliter.tagMatch(_Settings.m_xml_src_tagHL, false, gotoTag));
		return true;
	}
	return false;
}

void CMainFrame::SciMarginClicked(const SCNotification &scn)
{
	int lineClick = m_source.SendMessage(SCI_LINEFROMPOSITION, scn.position);
	if ((scn.modifiers & SCMOD_SHIFT) && (scn.modifiers & SCMOD_CTRL))
	{
		FoldAll();
	}
	else
	{
		int levelClick = m_source.SendMessage(SCI_GETFOLDLEVEL, lineClick);
		if (levelClick & SC_FOLDLEVELHEADERFLAG)
		{
			if (scn.modifiers & SCMOD_SHIFT)
			{
				// Ensure all children visible
				m_source.SendMessage(SCI_SETFOLDEXPANDED, lineClick, 1);
				ExpandFold(lineClick, true, true, 100, levelClick);
			}
			else if (scn.modifiers & SCMOD_CTRL)
			{
				if (m_source.SendMessage(SCI_GETFOLDEXPANDED, lineClick))
				{
					// Contract this line and all children
					m_source.SendMessage(SCI_SETFOLDEXPANDED, lineClick, 0);
					ExpandFold(lineClick, false, true, 0, levelClick);
				}
				else
				{
					// Expand this line and all children
					m_source.SendMessage(SCI_SETFOLDEXPANDED, lineClick, 1);
					ExpandFold(lineClick, true, true, 100, levelClick);
				}
			}
			else
			{
				// Toggle this line
				m_source.SendMessage(SCI_TOGGLEFOLD, lineClick);
			}
		}
	}
}

void CMainFrame::GoToSelectedTreeItem()
{
	CTreeItem ii(m_document_tree.GetSelectedItem());
	if (!ii.IsNull() && ii.GetData())
	{
		if (m_current_view != BODY)
			ShowView(BODY);
		GoTo((MSHTML::IHTMLElement *)ii.GetData());
		m_doc->m_body.SetFocus();
	}
}

MSHTML::IHTMLDOMNodePtr CMainFrame::RecursiveMoveRightElement(CTreeItem item)
{
	MSHTML::IHTMLDOMNodePtr ret;
	if (item.IsNull() || !item.GetData())
		return false;

	CTreeItem next_selected_sibling = m_document_tree.m_tree.m_tree.GetNextSelectedSibling(item);
	bool smart_selection = (!next_selected_sibling.IsNull()) && (item.GetNextSibling() != next_selected_sibling);

	if (smart_selection)
	{
		CTreeItem next_sibling = item.GetNextSibling();
		CTreeItem cur_selected = next_selected_sibling;
		while (!item.IsNull())
		{
			if (!item.GetData())
				return 0;
			MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElement *)item.GetData();
			if (!(bool)elem)
				return 0;

			MSHTML::IHTMLDOMNodePtr node = MSHTML::IHTMLDOMNodePtr(elem);

			if (!(bool)node)
				return 0;

			MoveRightElement(node);
			if (next_sibling.IsNull())
				break;

			item = next_sibling;
			next_sibling = next_sibling.GetNextSibling();

			if (!next_selected_sibling.IsNull() && next_sibling == next_selected_sibling)
			{
				item = next_selected_sibling;
				next_sibling = next_selected_sibling.GetNextSibling();
				cur_selected = next_selected_sibling;
				next_selected_sibling = m_document_tree.m_tree.m_tree.GetNextSelectedSibling(next_selected_sibling);
				continue;
			}
		}
		RecursiveMoveRightElement(m_document_tree.m_tree.m_tree.GetNextSelectedItem(cur_selected));
	}
	else
	{
		while (!item.IsNull())
		{
			MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElement *)item.GetData();
			if (!(bool)elem)
				return 0;

			MSHTML::IHTMLDOMNodePtr node = MSHTML::IHTMLDOMNodePtr(elem);
			if (!(bool)node)
				return 0;

			ret = MoveRightElement(node);

			item = m_document_tree.m_tree.m_tree.GetNextSelectedItem(item);
		}
	}
	return ret;
}

MSHTML::IHTMLDOMNodePtr CMainFrame::MoveRightElementWithoutChildren(MSHTML::IHTMLDOMNodePtr node)
{
	MSHTML::IHTMLDOMNodePtr move_from;
	MSHTML::IHTMLDOMNodePtr move_to;
	MSHTML::IHTMLDOMNodePtr insert_before;
	MSHTML::IHTMLDOMNodePtr ret;
	// делаем себя ребенком своего предыдущего брата
	// потом всех своих детей делаем своими братьями

	if (!(bool)(ret = MoveRightElement(node)))
		return 0;

	MSHTML::IHTMLDOMNodePtr parent = node->parentNode;
	MSHTML::IHTMLDOMNodePtr nextSibling = GetNextSiblingSection(node);

	MSHTML::IHTMLDOMNodePtr child = GetFirstChildSection(node);
	if ((bool)child)
	{
		move_to = parent;
		insert_before = 0;
		MSHTML::IHTMLDOMNodePtr nextChild;
		do
		{
			move_from = child;
			nextChild = GetNextSiblingSection(child);
			m_doc->MoveNode(move_from, move_to, insert_before);
			child = nextChild;
		} while (nextChild);
	}

	return ret;
}

MSHTML::IHTMLDOMNodePtr CMainFrame::MoveRightElement(MSHTML::IHTMLDOMNodePtr node)
{
	MSHTML::IHTMLDOMNodePtr move_from;
	MSHTML::IHTMLDOMNodePtr move_to;
	MSHTML::IHTMLDOMNodePtr insert_before;
	// делаем себя ребенком своего предыдущего брата

	if (!(bool)node)
		return 0;

	// пока будем таскать только секции
	if (!IsNodeSection(node))
		return 0;

	// если не можем переместить себя, то не дклаем ничего
	MSHTML::IHTMLDOMNodePtr prev_sibling = GetPrevSiblingSection(node);

	if (!(bool)prev_sibling)
		return 0;

	MSHTML::IHTMLDOMNodePtr child = GetLastChildSection(prev_sibling);

	// делаем себя последним ребенком своего предыдущего брата
	move_to = prev_sibling;
	insert_before = 0;
	move_from = node;

	if (!U::IsEmptySection(move_to))
	{
		CreateNestedSection(move_to);
	}

	return m_doc->MoveNode(move_from, move_to, insert_before);
}

MSHTML::IHTMLDOMNodePtr CMainFrame::MoveLeftElement(MSHTML::IHTMLDOMNodePtr node)
{
	MSHTML::IHTMLDOMNodePtr ret;
	// делаем себя  ближайшим братом своего отца
	// а своих следующих братьев своими детьми

	if (!(bool)node)
		return 0;

	// пока будем таскать только секции
	if (!IsNodeSection(node))
		return 0;

	// если не можем переместить себя, то не делаем ничего
	MSHTML::IHTMLDOMNodePtr parent = node->parentNode;
	if (!(bool)parent || !IsNodeSection(parent->parentNode))
		return 0;

	MSHTML::IHTMLDOMNodePtr sibling = node->nextSibling;

	while ((bool)sibling)
	{
		MSHTML::IHTMLDOMNodePtr next_sibling = sibling->nextSibling;
		m_doc->MoveNode(sibling, node, 0);
		sibling = next_sibling;
	}
	// делаем себя  ближайшим братом своего отца
	ret = m_doc->MoveNode(node, parent->parentNode, parent->nextSibling);

	return ret;
}

bool CMainFrame::IsNodeSection(MSHTML::IHTMLDOMNodePtr node)
{
	if (!(bool)node)
	{
		return false;
	}

	MSHTML::IHTMLElementPtr elem = MSHTML::IHTMLElementPtr(node);
	if (node->nodeType != NODE_ELEMENT)
		return false;

	return (U::scmp(elem->tagName, L"DIV") == 0 && (U::scmp(elem->className, L"section") == 0 || U::scmp(elem->className, L"body") == 0));
}

MSHTML::IHTMLDOMNodePtr CMainFrame::GetFirstChildSection(MSHTML::IHTMLDOMNodePtr node)
{
	if (!(bool)node)
		return 0;

	MSHTML::IHTMLDOMNodePtr child = node->firstChild;

	if (!(bool)child)
		return 0;

	if (IsNodeSection(child))
		return child;

	return GetNextSiblingSection(child);
}

MSHTML::IHTMLDOMNodePtr CMainFrame::GetNextSiblingSection(MSHTML::IHTMLDOMNodePtr node)
{
	if (!(bool)node)
		return 0;

	node = node->nextSibling;

	while (1)
	{
		if (!(bool)node)
			return 0;

		if (IsNodeSection(node))
			return node;

		node = node->nextSibling;
	}

	return 0;
}

MSHTML::IHTMLDOMNodePtr CMainFrame::GetPrevSiblingSection(MSHTML::IHTMLDOMNodePtr node)
{
	if (!(bool)node)
		return 0;

	node = node->previousSibling;

	while (1)
	{
		if (!(bool)node)
			return 0;

		if (IsNodeSection(node))
			return node;

		node = node->previousSibling;
	}

	return 0;
}

MSHTML::IHTMLDOMNodePtr CMainFrame::GetLastChildSection(MSHTML::IHTMLDOMNodePtr node)
{
	if (!(bool)node)
		return 0;

	MSHTML::IHTMLDOMNodePtr child = node->lastChild;

	if (!(bool)child)
		return 0;

	if (IsNodeSection(child))
		return child;

	return GetPrevSiblingSection(child);
}


MSHTML::IHTMLDOMNodePtr CMainFrame::CreateNestedSection(MSHTML::IHTMLDOMNodePtr node)
{
	MSHTML::IHTMLDOMNodePtr section = node->firstChild;
	MSHTML::IHTMLDOMNodePtr new_node;
	if (!(bool)section)
		return 0;
	do
	{
		_bstr_t tag_name(section->nodeName);
		MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElementPtr)section;
		_bstr_t class_name(elem->className);

		if ((0 == U::scmp(tag_name, L"DIV")) &&
			((0 == U::scmp(class_name, L"section")) || (0 == U::scmp(class_name, L"title")) || (0 == U::scmp(class_name, L"epigraph")) || (0 == U::scmp(class_name, L"annotation")) || (0 == U::scmp(class_name, L"image"))))
		{
			continue;
		}

		MSHTML::IHTMLElementPtr new_elem = m_doc->m_body.Document()->createElement(L"DIV");
		new_elem->className = L"section";
		new_node = MSHTML::IHTMLDOMNodePtr(new_elem);
		MSHTML::IHTMLDOMNodePtr insert_before = section;
		m_doc->MoveNode(new_node, node, insert_before);
		do
		{
			MSHTML::IHTMLDOMNodePtr next_node = section->nextSibling;
			m_doc->MoveNode(section, new_node, 0);
			section = next_node;
		} while ((bool)section);
		break;

	} while ((bool)(section = section->nextSibling));

	return new_node;
}

///<summary>Set caret in SOURCE view</summary>
///<returns>FILE_OP_STATUS</summary>
void CMainFrame::SourceGoTo(int line, int col)
{
	int pos = m_source.SendMessage(SCI_POSITIONFROMLINE, line - 1);
	while (col--)
		pos = m_source.SendMessage(SCI_POSITIONAFTER, pos);
	m_source.SendMessage(SCI_SETSELECTIONSTART, pos);
	m_source.SendMessage(SCI_SETSELECTIONEND, pos);
	m_source.SendMessage(SCI_SCROLLCARET);
}

void CMainFrame::SetIsText()
{
	if (m_is_fail)
		m_status.SetText(ID_DEFAULT_PANE, L"Failing Incremental Search: " + m_is_str);
	else
		m_status.SetText(ID_DEFAULT_PANE, L"Incremental Search: " + m_is_str);
}

void CMainFrame::StopIncSearch(bool fCancel)
{
	if (!m_incsearch)
		return;
	m_incsearch = 0;
	m_sel_changed = true; // will cause status line update soon
	if (fCancel)
		m_doc->m_body.CancelIncSearch();
	else
		m_doc->m_body.StopIncSearch();
}

void CMainFrame::GoTo(MSHTML::IHTMLElement *e)
{
	m_doc->m_body.GoTo(e);
}

bool CMainFrame::ShowSettingsDialog(HWND parent)
{
	CSettingsDlg dlg(IDS_SETTINGS);
	return dlg.DoModal(parent) == IDOK;
}

void CMainFrame::ApplyConfChanges()
{
	CWaitCursor *hourglass = new CWaitCursor();
	LONG visible = false;

	wchar_t restartMsg[MAX_LOAD_STRING + 1];
	::LoadString(_Module.GetResourceInstance(), IDS_SETTINGS_NEED_RESTART, restartMsg, MAX_LOAD_STRING);
	if (m_doc)
		m_doc->m_body.ApplyConfChanges();
	SetupSci();
	SetSciStyles();

	// added by SeNS: display line numbers
	/*if (_Settings.m_show_line_numbers)
		m_source.SendMessage(SCI_SETMARGINWIDTHN,0,64);
	else
		m_source.SendMessage(SCI_SETMARGINWIDTHN,0,0);*/

		/*XmlMatchedTagsHighlighter xmlTagMatchHiliter(&m_source);
			xmlTagMatchHiliter.tagMatch(_Settings.m_xml_src_tagHL, false, false);
			UIEnable(ID_GOTO_MATCHTAG, _Settings.m_xml_src_tagHL);
			*/
			// added by SeNS
	if (_Settings.m_usespell_check)
	{
		CString custDictName = _Settings.m_custom_dict;
		if (!custDictName.IsEmpty())
			m_Speller->SetCustomDictionary(U::GetProgDir() + custDictName);
		if (!m_Speller->Enabled())
			m_Speller->Enable(true);
	}
	// don't use spellchecker
	else if (m_Speller)
		m_Speller->Enable(false);

	// added by SeNS: issue 17: process nbsp settings change
	if (_Settings.GetOldNBSPChar().Compare(_Settings.GetNBSPChar()) != 0)
	{
		//int numChanges = 0;
		// save caret position
		MSHTML::IDisplayServicesPtr ids(MSHTML::IDisplayServicesPtr(m_doc->m_body.Document()));
		MSHTML::IHTMLCaretPtr caret = nullptr;
		MSHTML::tagPOINT *point = new MSHTML::tagPOINT();
		if (ids)
		{
			ids->GetCaret(&caret);
			if (caret)
			{
				caret->IsVisible(&visible);
				if (visible)
					caret->GetLocation(point, true);
			}
		}

		MSHTML::IHTMLElementPtr fbwBody = MSHTML::IHTMLDocument3Ptr(m_doc->m_body.Document())->getElementById(L"fbw_body");
		MSHTML::IHTMLDOMNodePtr el = MSHTML::IHTMLDOMNodePtr(fbwBody)->firstChild;
		CString nbsp = _Settings.GetNBSPChar();
		while (el && el != fbwBody)
		{
			if (el->nodeType == NODE_TEXT)
			{
				CString s = el->nodeValue;
				int n = s.Replace(_Settings.GetOldNBSPChar(), _Settings.GetNBSPChar());
				if (n)
				{
					//numChanges += n;
					el->nodeValue = s.GetString();
				}
			}
			if (el->firstChild)
				el = el->firstChild;
			else
			{
				while (el && el != fbwBody && el->nextSibling == NULL)
					el = el->parentNode;
				if (el && el != fbwBody)
					el = el->nextSibling;
			}
		}

		// restore caret position
		if (caret && visible)
		{
			MSHTML::IDisplayPointerPtr disptr;
			ids->CreateDisplayPointer(&disptr);
			disptr->moveToPoint(*point, MSHTML::COORD_SYSTEM_GLOBAL, fbwBody, 0, 0);
			caret->MoveCaretToPointer(disptr, true, MSHTML::CARET_DIRECTION_SAME);
		}
	}

	_Settings.SaveHotkeyGroups();
	_Settings.Save();
	_Settings.SaveWords();

	delete hourglass;

	if (_Settings.NeedRestart() && AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)restartMsg, (LPCTSTR)NULL, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, TD_WARNING_ICON) == IDYES)
	{
		return RestartProgram();
	}
}

void CMainFrame::RestartProgram()
{
	BOOL b = false;
	if (OnClose(0, 0, 0, b))
	{
		wchar_t filename[MAX_PATH];
		::GetModuleFileName(_Module.GetModuleInstance(), filename, MAX_PATH);
		CString ofn = m_doc->GetDocFileName();
		ofn = L"\"" + ofn + L"\"";
		ShellExecute(0, L"open", filename, ofn, 0, SW_SHOW);
	}
}

void CMainFrame::CollectScripts(CString path, TCHAR *mask, int lastid, CString refid)
{
	if (U::HasFilesWithExt(path, mask))
	{
		lastid = GrabScripts(path, mask, refid);
	}

	if (U::HasSubFolders(path))
	{
		WIN32_FIND_DATA fd;
		HANDLE found = FindFirstFile(path + L"*.*", &fd);
		if (found)
		{
			do
			{
				if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					if (wcscmp(fd.cFileName, L".") && wcscmp(fd.cFileName, L"..") && U::HasScriptsEndpoint((path + fd.cFileName) + L"\\", mask))
					{
						ScrInfo folder;

						wchar_t *Name = new wchar_t[wcslen(fd.cFileName) + 1];
						wchar_t *pos = wcschr(fd.cFileName, L'_');

						folder.order = L"0_";

						if (!pos || !U::CheckScriptsVersion(fd.cFileName))
						{
							wcscpy(Name, fd.cFileName);
							folder.order += Name;
						}
						else
						{
							wcscpy(Name, pos + 1);
							folder.order = fd.cFileName;
						}

						folder.name = Name;

						CString temp;
						temp.Format(L"_%d", lastid);
						folder.id = refid + temp;
						folder.refid = refid;
						folder.isFolder = true;
						//folder.accel.key = 0;

						folder.picture = NULL;
						folder.pictType = CMainFrame::NO_PICT;

						WIN32_FIND_DATA picFd;
						wchar_t *picName = new wchar_t[wcslen(fd.cFileName) + 1];
						wcscpy(picName, fd.cFileName);

						picName[wcslen(picName)] = 0;
						CString picPathNoExt = (path + picName);
						HANDLE hPicture = FindFirstFile(picPathNoExt + L".bmp", &picFd);
						HANDLE hIcon = FindFirstFile(picPathNoExt + L".ico", &picFd);

						if (hPicture != INVALID_HANDLE_VALUE)
						{
							HBITMAP bitmap = (HBITMAP)LoadImage(NULL, (picPathNoExt + L".bmp").GetBuffer(),
								IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
							if (bitmap != NULL)
							{
								folder.picture = bitmap;
								folder.pictType = CMainFrame::BITMAP;
							}
						}

						else if (hIcon != INVALID_HANDLE_VALUE)
						{
							HICON icon = (HICON)LoadImage(NULL, (picPathNoExt + L".ico").GetBuffer(),
								IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
							if (icon != NULL)
							{
								folder.picture = icon;
								folder.pictType = CMainFrame::ICON;
							}
						}

						FindClose(hPicture);
						FindClose(hIcon);
						delete[] picName;

						m_scripts.Add(folder);
						delete[] Name;

						CollectScripts((path + fd.cFileName) + L"\\", mask, 1, folder.id);
						lastid++;
					}
				}
			} while (FindNextFile(found, &fd));

			FindClose(found);
		}
	}
}

int CMainFrame::GrabScripts(CString path, TCHAR *mask, CString refid)
{
	WIN32_FIND_DATA fd;
	HANDLE found = FindFirstFile(path + mask, &fd);
	int newid = 1;

	if (found)
	{
		do
		{
			if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				if (StartScript(this) == 0 && SUCCEEDED(ScriptLoad(path + fd.cFileName)) && ScriptFindFunc(L"Run"))
				{
					{
						ScrInfo script;
						wchar_t *Name = new wchar_t[wcslen(fd.cFileName) + 1];
						wchar_t *pos = wcschr(fd.cFileName, L'_');

						script.order = L"0_";

						if (!pos || !U::CheckScriptsVersion(fd.cFileName))
						{
							wcscpy(Name, fd.cFileName);
							script.order += Name;
						}
						else
						{
							wcscpy(Name, pos + 1);
							script.order = fd.cFileName;
						}

						Name[wcslen(Name) - 3] = 0;
						script.name = Name;
						script.path = path + fd.cFileName;

						script.picture = NULL;
						script.pictType = CMainFrame::NO_PICT;
						WIN32_FIND_DATA picFd;
						wchar_t *picName = new wchar_t[wcslen(fd.cFileName) + 1];
						wcscpy(picName, fd.cFileName);

						picName[wcslen(picName) - 3] = 0;
						CString picPathNoExt = (path + picName);
						HANDLE hPicture = FindFirstFile(picPathNoExt + L".bmp", &picFd);
						HANDLE hIcon = FindFirstFile(picPathNoExt + L".ico", &picFd);

						if (hPicture != INVALID_HANDLE_VALUE)
						{
							HBITMAP bitmap = (HBITMAP)LoadImage(NULL, (picPathNoExt + L".bmp").GetBuffer(), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
							if (bitmap != NULL)
							{
								script.picture = bitmap;
								script.pictType = CMainFrame::BITMAP;
							}
						}
						else if (hIcon != INVALID_HANDLE_VALUE)
						{
							HICON icon = (HICON)LoadImage(NULL, (picPathNoExt + L".ico").GetBuffer(), IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
							if (icon != NULL)
							{
								script.picture = icon;
								script.pictType = CMainFrame::ICON;
							}
						}

						FindClose(hPicture);
						FindClose(hIcon);
						delete[] picName;

						/*CComVariant accel;
						ZeroMemory(&script.accel, sizeof(script.accel));
						script.accel.key = 0;*/

						/*if (SUCCEEDED(ScriptCall(L"SetHotkey", NULL, 0, &accel)))
						{
							TCHAR errCaption[MAX_LOAD_STRING + 1];
							LoadString(_Module.GetResourceInstance(), IDS_ERRMSGBOX_CAPTION, errCaption, MAX_LOAD_STRING);

							int j;
							for(j = 0; j < m_scripts.GetSize(); ++j)
							{
								if(m_scripts[j].accel.key == accel.intVal && m_scripts[j].accel.key != 0 && !m_scripts[j].isFolder)
								{
									if(_Settings.GetScriptsHkErrNotify())
									{
										CString errDescr;
										errDescr.Format(IDS_SCRIPT_HOTKEY_CONFLICT, m_scripts[j].path, script.path);
										MessageBox(errDescr.GetBuffer(), errCaption, MB_OK|MB_ICONSTOP);
									}
									break;
								}
							}
							if(j == m_scripts.GetSize() && keycodes.FindKey(accel.intVal) != -1)
								script.accel.key = accel.intVal;
						}*/

						CString temp;
						temp.Format(L"_%d", newid);
						script.id = refid + temp;
						script.refid = refid;
						script.isFolder = false;
						script.Type = 2;
						m_scripts.Add(script);
						newid++;

						delete[] Name;
					}
					StopScript();
				}
			}
		} while (FindNextFile(found, &fd));

		FindClose(found);
	}

	return newid;
}

void CMainFrame::AddScriptsSubMenu(HMENU parentItem, CString refid, CSimpleArray<ScrInfo> &scripts)
{
	MENUITEMINFO mi;
	static int SCRIPT_COMMAND_ID = 1;
	int menupos = 0;

	for (int i = 0; i < scripts.GetSize(); ++i)
	{
		memset(&mi, NULL, sizeof(MENUITEMINFO));
		mi.cbSize = sizeof(MENUITEMINFO);
		mi.fMask = MIIM_TYPE | MIIM_STATE;
		mi.fType = MFT_STRING;

		for (int j = 0; j < scripts.GetSize(); j++)
		{
			if (scripts[j].refid == refid)
				menupos++;
		}

		if (scripts[i].refid == refid)
		{
			if (scripts[i].isFolder)
			{
				mi.fMask |= MIIM_SUBMENU | MIIM_ID;
				mi.hSubMenu = CreateMenu();
				mi.wID = SCRIPT_COMMAND_ID++;
				scripts[i].wID = -1;
				AddScriptsSubMenu(mi.hSubMenu, scripts[i].id, scripts);
			}
			else
			{
				mi.fMask |= MIIM_ID;
				mi.wID = ID_SCRIPT_BASE + SCRIPT_COMMAND_ID;
				scripts[i].wID = SCRIPT_COMMAND_ID;
				SCRIPT_COMMAND_ID++;

				InitScriptHotkey(scripts[i]);
			}

			mi.dwTypeData = scripts[i].name.GetBuffer();
			mi.cch = wcslen(scripts[i].name);

			if (scripts[i].isFolder)
				InsertMenuItem(parentItem, 0, true, &mi);
			else
			{
				InsertMenuItem(parentItem, menupos--, true, &mi);
				// added by SeNS: add scripts with icon to toolbar
				if (scripts[i].pictType == CMainFrame::ICON)
					AddTbButton(m_ScriptsToolbar, scripts[i].name, mi.wID, TBSTATE_ENABLED, (HICON)scripts[i].picture);
			}

			switch (scripts[i].pictType)
			{
			case CMainFrame::BITMAP:
				m_MenuBar.AddBitmap((HBITMAP)scripts[i].picture, mi.wID);
				break;
			case CMainFrame::ICON:
				m_MenuBar.AddIcon((HICON)scripts[i].picture, mi.wID);
				break;
			}
		}
	}
}

void CMainFrame::QuickScriptsSort(CSimpleArray<ScrInfo> &scripts, int min, int max)
{
	int i, j;
	ScrInfo mid, tmp;
	{
		if (min < max)
		{
			mid = scripts[min];
			i = min - 1;
			j = max + 1;
			while (i < j)
			{
				do
				{
					i++;

				} while (!(scripts[i].order.CompareNoCase(mid.order.GetBuffer()) >= 0));
				do
				{
					j--;
				} while (!(scripts[j].order.CompareNoCase(mid.order.GetBuffer()) <= 0));
				if (i < j)
				{
					tmp = scripts[i];
					scripts[i] = scripts[j];
					scripts[j] = tmp;
				}
			}

			QuickScriptsSort(scripts, min, j);
			QuickScriptsSort(scripts, j + 1, max);
		}
	}
}

void CMainFrame::UpScriptsFolders(CSimpleArray<ScrInfo> &scripts)
{
	for (int i = 0; i < scripts.GetSize(); ++i)
	{
		if (!scripts[i].isFolder)
		{
			for (int j = i; j < scripts.GetSize(); ++j)
			{
				if (scripts[j].isFolder)
				{
					for (int k = j; k > i; --k)
					{
						ScrInfo tmp = scripts[k - 1];
						scripts[k - 1] = scripts[k];
						scripts[k] = tmp;
					}
				}
			}
		}
	}
}

void CMainFrame::InitScriptHotkey(CMainFrame::ScrInfo &script)
{
	std::vector<CHotkeysGroup> &hotkey_groups = _Settings.m_hotkey_groups;
	for (unsigned int i = 0; i < hotkey_groups.size(); ++i)
	{
		if (hotkey_groups.at(i).m_reg_name == L"Scripts")
		{
			CHotkey ScriptsHotkey(script.path,
				script.name,
				NULL,
				WORD(ID_SCRIPT_BASE + script.wID),
				NULL,
				script.path);
			hotkey_groups.at(i).m_hotkeys.push_back(ScriptsHotkey);
		}
	}
}

void CMainFrame::InitPluginHotkey(CString guid, UINT cmd, CString name)
{
	std::vector<CHotkeysGroup> &hotkey_groups = _Settings.m_hotkey_groups;
	for (unsigned int i = 0; i < hotkey_groups.size(); ++i)
	{
		if (hotkey_groups.at(i).m_reg_name == L"Plugins")
		{
			CHotkey PluginsHotkey(guid,
				name,
				NULL,
				WORD(cmd),
				NULL);
			hotkey_groups.at(i).m_hotkeys.push_back(PluginsHotkey);
		}
	}
}

//
// Idea by Sclex
//
/// <summary>Replace NBSP in text content of node on settings character</summary>
/// <param name="node">HTML-node</param>
void CMainFrame::ChangeNBSP(MSHTML::IHTMLDOMNodePtr node)
{
	if (!node)
		return;
	MSHTML::IHTMLDOMNodePtr cur_node = nullptr;
	MSHTML::IHTMLDOMChildrenCollectionPtr children = node->childNodes;

	for (int i = 0; i < children->length; i++)
	{
		cur_node = children->item(i);
		if (cur_node->nodeType == NODE_TEXT)
		{
			CString s(cur_node->nodeValue);
			if (s.Replace(L"\u00A0", _Settings.GetNBSPChar()) > 0)
				cur_node->nodeValue = s.GetString();
		}
		else
			ChangeNBSP(cur_node);
	}
	/*MSHTML::IHTMLElementPtr fbwBody = MSHTML::IHTMLDocument3Ptr(m_doc->m_body.Document())->getElementById(L"fbw_body");
	if (fbwBody && elem && fbwBody->contains(elem))
	{
		// save caret position
		MSHTML::IHTMLTxtRangePtr tr1;
		int offset = 0;
		MSHTML::IHTMLTxtRangePtr sel(m_doc->m_body.Document()->selection->createRange());
		if (sel)
		{
			tr1 = sel->duplicate();
			if (tr1)
			{
				tr1->moveToElementText(elem);
				tr1->setEndPoint(L"EndToStart", sel);
				CString s = tr1->text;
				offset = s.GetLength();
				// special fix for strange MSHTML bug (inline image present in html code)
				CString s2 = tr1->htmlText;
				int l = s2.Replace(L"<IMG", L"<IMG");
				offset += (l * 3);
			}
		}

		MSHTML::IHTMLDOMNodePtr el = MSHTML::IHTMLDOMNodePtr(elem)->firstChild;
		CString nbsp = _Settings.GetNBSPChar();
		CString s;
		int numChanges = 0;

		while (el && el != elem)
		{
			if (el->nodeType == 3)
			{
				try
				{
					s = el->nodeValue;
				}
				catch (...)
				{
					break;
				}
				int n = s.Replace(L"\u00A0", _Settings.GetNBSPChar());
				int k = s.Replace(L"<p>\u00A0<p>", L"<p><p>");
				if (n || k)
				{
					numChanges += n + k;
					el->nodeValue = s.AllocSysString();
				}
			}
			if (el->firstChild)
				el = el->firstChild;
			else
			{
				while (el && el != elem && el->nextSibling == NULL)
					el = el->parentNode;
				if (el && el != elem)
					el = el->nextSibling;
			}
		}

		if (numChanges)
		{
			// restore caret position
			if (tr1)
			{
				tr1->moveToElementText(elem);
				tr1->collapse(true);
				if (offset == 0)
				{
					tr1->move(L"character", 1);
					tr1->move(L"character", -1);
				}
				else
					tr1->move(L"character", offset);
				tr1->select();
			}
		}
	}*/
}

void CMainFrame::RemoveLastUndo()
{
	// remove last undo operation
	IServiceProviderPtr serviceProvider = IServiceProviderPtr(m_doc->m_body.Document());
	CComPtr<IOleUndoManager> undoManager;
	CComPtr<IOleUndoUnit> undoUnit[10];
	CComPtr<IEnumOleUndoUnits> undoUnits;
	if (SUCCEEDED(serviceProvider->QueryService(SID_SOleUndoManager, IID_IOleUndoManager, (void **)&undoManager)))
	{
		undoManager->EnumUndoable(&undoUnits);
		if (undoUnits)
		{
			ULONG numUndos = 0;
			undoUnits->Next(10, &undoUnit[0], &numUndos);
			// delete whole stack
			undoManager->DiscardFrom(NULL);
			// restore all except previous
			if (numUndos)
				for (ULONG i = 0; i < numUndos - 1; i++)
					undoManager->Add(undoUnit[i]);
		}
	}
}

LRESULT CMainFrame::OnEdChange(WORD, WORD, HWND hWnd, BOOL &b)
{
	//CString ss = MSHTML::IHTMLElementPtr(m_doc->m_body.GetChangedNode())->outerHTML;
	/*MSHTML::IHTMLDOMNodePtr chp(m_doc->m_body.GetChangedNode());
	if (!chp)
		return S_OK;
	CString test = MSHTML::IHTMLElementPtr(chp)->outerHTML;*/
	StopIncSearch(true);

	m_doc->Changed(true);
	m_doc->m_body.Synched(false);
	//m_cb_updated = false;

	// update document tree
	if (m_current_view == BODY)
	{
		MSHTML::IHTMLDOMNodePtr chp(m_doc->m_body.GetChangedNode());
		if (chp && m_document_tree.IsWindowVisible())
		{
			m_document_tree.UpdateDocumentStructure(m_doc->m_body.Document(), chp);
			m_document_tree.HighlightItemAtPos(m_doc->m_body.SelectionContainer());
		}
		// added by SeNS: update
		UpdateViewSizeInfo();
		// process nbsp in changed element
		if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
			ChangeNBSP(m_doc->m_body.SelectionContainer());

		// added by SeNS: do spellcheck
		if (m_Speller->Enabled())
			m_Speller->CheckElement(m_doc->m_body.SelectionContainer(), m_doc->m_body.IsHTMLChanged());
	}

	return S_OK;
}

///<summary>Show on StatusBar hex-code of char on carret</summary>
void CMainFrame::DisplayCharCode()
{
	CString s(L"");
	if (m_current_view == SOURCE)
	{
		WCHAR dd = WCHAR(m_source.SendMessage(SCI_GETCHARAT, m_source.SendMessage(SCI_GETCURRENTPOS)));
		s.Format(L"  U+%.4X", dd);
	}
	else if (m_current_view == BODY)
	{
		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_doc->m_body.Document())->getSelection();
		if (!sel)
			return;
		if (sel->rangeCount == 0 || !sel->isCollapsed)
			return;
		MSHTML::IHTMLDOMRangePtr range = sel->getRangeAt(0)->cloneRange();
		if (!range)
			return;
		s.SetString(range->endContainer->nodeValue.bstrVal);
		if (range->endOffset < s.GetLength())
			s.Format(L"  U+%.4X", s.GetAt(range->endOffset));
		else
			s.SetString(L"U+0000");
	}
	m_status.SetPaneText(ID_PANE_CHAR, s);
}

// added by SeNS - paste pictures
bool CMainFrame::BitmapInClipboard()
{
	bool result = false;
	if (OpenClipboard())
	{
		if (IsClipboardFormatAvailable(CF_BITMAP))
			result = true;
		CloseClipboard();
	}
	return result;
}

void CMainFrame::UpdateViewSizeInfo()
{
	if (m_doc && m_doc->m_body)
		if (m_doc->m_body.Document())
		{
			MSHTML::IHTMLElement2Ptr m_scrollElement = MSHTML::IHTMLDocument3Ptr(m_doc->m_body.Document())->documentElement;
			if (m_scrollElement)
			{
				_Settings.SetViewWidth(m_scrollElement->clientWidth);
				_Settings.SetViewHeight(m_scrollElement->clientHeight);
				_Settings.SetMainWindow(m_hWnd);
			}
		}
}
