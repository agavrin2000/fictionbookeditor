// FBEView.h : interface of the CFBEView class
//
/////////////////////////////////////////////////////////////////////////////

#include <wtl/atlcrack.h>
#if !defined(AFX_FBEVIEW_H__E0C71279_419D_4273_93E3_57F6A57C7CFE__INCLUDED_)
#define AFX_FBEVIEW_H__E0C71279_419D_4273_93E3_57F6A57C7CFE__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#pragma warning(disable : 4996)
#endif // _MSC_VER >= 1000

#include "resource.h"
#include "Settings.h"

extern CSettings _Settings;

// void BubbleUp(MSHTML::IHTMLDOMNode *node, const wchar_t *name);

static void CenterChildWindow(CWindow parent, CWindow child);

template <class T, int chgID>
class ATL_NO_VTABLE CHTMLChangeSink : public MSHTML::IHTMLChangeSink
{
protected:
public:
  // IUnknown
  STDMETHOD(QueryInterface)
  (REFIID iid, void **ppvObject)
  {
    if (iid == IID_IUnknown || iid == IID_IHTMLChangeSink)
    {
      *ppvObject = this;
      return S_OK;
    }

    return E_NOINTERFACE;
  }
  STDMETHOD_(ULONG, AddRef)
  () { return 1; }
  STDMETHOD_(ULONG, Release)
  () { return 1; }

  // IHTMLChangeSink
  STDMETHOD(raw_Notify)
  ()
  {
    T *pT = static_cast<T *>(this);
    pT->EditorChanged(chgID);
    return S_OK;
  }
};

typedef CWinTraits<WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0>
    CFBEViewWinTraits;

enum
{
  FWD_SINK,
  BACK_SINK,
  RANGE_SINK
};

class CFindDlgBase;

class CFBEView : public CWindowImpl<CFBEView, CAxWindow, CFBEViewWinTraits>,
                 public IDispEventSimpleImpl<0, CFBEView, &DIID_DWebBrowserEvents2>,
                 public IDispEventSimpleImpl<0, CFBEView, &DIID_HTMLDocumentEvents2>,
                 public IDispEventSimpleImpl<0, CFBEView, &DIID_HTMLTextContainerEvents2>,
                 public CHTMLChangeSink<CFBEView, RANGE_SINK>
{
protected:
  typedef IDispEventSimpleImpl<0, CFBEView, &DIID_DWebBrowserEvents2> BrowserEvents;
  typedef IDispEventSimpleImpl<0, CFBEView, &DIID_HTMLDocumentEvents2> DocumentEvents;
  typedef IDispEventSimpleImpl<0, CFBEView, &DIID_HTMLTextContainerEvents2> TextEvents;
  typedef CHTMLChangeSink<CFBEView, FWD_SINK> ForwardSink;
  typedef CHTMLChangeSink<CFBEView, BACK_SINK> BackwardSink;
  typedef CHTMLChangeSink<CFBEView, RANGE_SINK> RangeSink;

  HWND m_frame;
public:
	CString m_file_name, m_file_path;
	// changed by SeNS
	DWORD m_dirtyRangeCookie;
	MSHTML::IMarkupServices2Ptr m_mk_srv;
	bool m_synched;
	bool IsDescFormChanged()
	{
		return m_decr_form_changed;
	}
	
	void DescFormChanged(bool state)
	{
		m_decr_form_changed = state;
	}
protected:
  SHD::IWebBrowser2Ptr m_browser;
  MSHTML::IHTMLDocument2Ptr m_hdoc;
  MSHTML::IMarkupContainer2Ptr m_mkc;

  int m_ignore_changes;
  int m_enable_paste;
  bool m_normalize;
  bool m_complete : 1;
  bool m_initialized;
  bool m_decr_form_changed;

  MSHTML::IHTMLElementPtr m_cur_sel;
  MSHTML::IHTMLElementPtr m_cur_input_element;
  CString m_cur_val;

  CString m_nav_url;

  static _ATL_FUNC_INFO DocumentCompleteInfo;
  static _ATL_FUNC_INFO BeforeNavigateInfo;
  static _ATL_FUNC_INFO EventInfo;
  static _ATL_FUNC_INFO VoidEventInfo;
  static _ATL_FUNC_INFO VoidInfo;

  enum
  {
    FRF_REVERSE = 1,
    FRF_WHOLE = 2,
    FRF_CASE = 4,
    FRF_REGEX = 8
  };

  struct FindReplaceOptions
  {
    CString pattern;
    CString replacement;
    U::ReMatch match;
    int flags; // IHTMLTxtRange::findText() flags
    bool fRegexp;

    int replNum;

    FindReplaceOptions() : fRegexp(false), flags(0) {}
  };

  FindReplaceOptions m_fo;
  MSHTML::IHTMLTxtRangePtr m_is_start;

  struct pElAdjacent
  {
    MSHTML::IHTMLElementPtr elem;
    _bstr_t innerText;

    pElAdjacent(MSHTML::IHTMLElementPtr pElem)
    {
      elem = pElem;
      innerText = pElem->innerText;
    }
  };

  friend class CFindDlgBase;
  friend class CViewFindDlg;
  friend class CReplaceDlgBase;
  friend class CViewReplaceDlg;
  friend class CSciFindDlg;
  friend class CSciReplaceDlg;
  friend class FRBase;

  int TextOffset(MSHTML::IHTMLTxtRange *rng, U::ReMatch rm, CString txt = L"", CString htmlTxt = L"");

  void SelMatch(MSHTML::IHTMLTxtRange *tr, U::ReMatch rm);
  MSHTML::IHTMLElementPtr SelectionContainerImp();

public:
  CFindDlgBase *m_find_dlg;
  CReplaceDlgBase *m_replace_dlg;

  SHD::IWebBrowser2Ptr Browser();
  MSHTML::IHTMLDocument2Ptr Document();
  void StopIncSearch();
  long GetVersionNumber();

  IDispatchPtr Script();
  CString NavURL();
  bool Loaded();
  void Init();


  CAtlList<CString> m_UndoStrings;
  void BeginUndoUnit(const wchar_t *name);
  void EndUndoUnit();

  DECLARE_WND_SUPERCLASS(NULL, CAxWindow::GetWndClassName());

  CFBEView(HWND frame, bool fNorm) : m_frame(frame), m_ignore_changes(0), m_enable_paste(0), m_dirtyRangeCookie(0), m_trace_changes(false),
                                     m_normalize(fNorm), m_complete(false), m_initialized(false), m_startMatch(0), m_endMatch(0),
                                     m_decr_form_changed(false), m_find_dlg(0), m_replace_dlg(0), m_file_path(L""), m_file_name(L"")
  {
  }
  ~CFBEView();

  BOOL PreTranslateMessage(MSG *pMsg);

  BEGIN_MSG_MAP(CFBEView)
      MESSAGE_HANDLER(WM_CREATE, OnCreate)
      MESSAGE_HANDLER(WM_SETFOCUS, OnFocus)

      // editing commands
      COMMAND_ID_HANDLER(ID_EDIT_UNDO, OnUndo)
      COMMAND_ID_HANDLER(ID_EDIT_REDO, OnRedo)
      COMMAND_ID_HANDLER(ID_EDIT_CUT, OnCut)
      COMMAND_ID_HANDLER(ID_EDIT_COPY, OnCopy)
      COMMAND_ID_HANDLER(ID_EDIT_PASTE, OnPaste)
	  COMMAND_ID_HANDLER(ID_EDIT_FIND, OnFind)
      COMMAND_ID_HANDLER(ID_EDIT_FINDNEXT, OnFindNext)
      COMMAND_ID_HANDLER(ID_EDIT_REPLACE, OnReplace)

      COMMAND_ID_HANDLER(ID_STYLE_LINK, OnStyleLink)
      COMMAND_ID_HANDLER(ID_STYLE_NOTE, OnStyleFootnote)
      COMMAND_ID_HANDLER(ID_STYLE_NOLINK, OnStyleNolink)
	  COMMAND_ID_HANDLER(ID_STYLE_BOLD, OnBold)
	  COMMAND_ID_HANDLER(ID_STYLE_ITALIC, OnItalic)
	  COMMAND_ID_HANDLER(ID_STYLE_STRIKETROUGH, OnStrik)
	  COMMAND_ID_HANDLER(ID_STYLE_SUPERSCRIPT, OnSup)
	  COMMAND_ID_HANDLER(ID_STYLE_SUBSCRIPT, OnSub)
	  COMMAND_ID_HANDLER(ID_STYLE_CODE, OnStyleCode)

      COMMAND_ID_HANDLER(ID_STYLE_NORMAL, OnStyleNormal)

	  COMMAND_ID_HANDLER(ID_ADD_BODY, OnEditAddBody)
	  COMMAND_ID_HANDLER(ID_ADD_BODY_NOTES, OnEditAddBodyNotes)
	  COMMAND_ID_HANDLER(ID_ADD_SECTION, OnEditAddSection)
	  COMMAND_ID_HANDLER(ID_ADD_ANN, OnEditAddAnn)
	  COMMAND_ID_HANDLER(ID_ADD_TITLE, OnEditAddTitle)
	  COMMAND_ID_HANDLER(ID_ADD_SUBTITLE, OnEditAddSubtitle)
	  COMMAND_ID_HANDLER(ID_ADD_POEM, OnEditAddPoem)
	  COMMAND_ID_HANDLER(ID_ADD_CITE, OnEditAddCite)
	  COMMAND_ID_HANDLER(ID_ADD_STANZA, OnEditAddStanza)
	  COMMAND_ID_HANDLER(ID_ADD_EPIGRAPH, OnEditAddEpigraph)
	  COMMAND_ID_HANDLER(ID_ADD_TEXTAUTHOR, OnEditAddTA)
	  COMMAND_ID_HANDLER(ID_ADD_IMAGE, OnEditInsImage)

	  COMMAND_ID_HANDLER(ID_EDIT_CLONE, OnEditClone)
      COMMAND_ID_HANDLER(ID_EDIT_SPLIT, OnEditSplit)
      COMMAND_ID_HANDLER(ID_EDIT_MERGE, OnEditMerge)
      COMMAND_ID_HANDLER(ID_EDIT_REMOVE_OUTER, OnEditRemoveOuter)

      COMMAND_ID_HANDLER_EX(ID_INS_TABLE, OnEditInsertTable)


      COMMAND_ID_HANDLER(ID_VIEW_HTML, OnViewHTML)
      COMMAND_ID_HANDLER(ID_SAVEIMG_AS, OnSaveImageAs)
      COMMAND_RANGE_HANDLER(ID_SEL_BASE, ID_SEL_BASE + 99, OnSelectElement)
  END_MSG_MAP()

  BEGIN_SINK_MAP(CFBEView)
      SINK_ENTRY_INFO(0, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete, &DocumentCompleteInfo)
      SINK_ENTRY_INFO(0, DIID_DWebBrowserEvents2, DISPID_BEFORENAVIGATE2, OnBeforeNavigate, &BeforeNavigateInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONSELECTIONCHANGE, OnSelChange, &VoidEventInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONCONTEXTMENU, OnContextMenu, &EventInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONCLICK, OnClick, &EventInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONKEYDOWN, OnKeyDown, &EventInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONFOCUSIN, OnFocusInOut, &VoidEventInfo)
	  SINK_ENTRY_INFO(0, DIID_HTMLDocumentEvents2, DISPID_HTMLDOCUMENTEVENTS2_ONFOCUSOUT, OnFocusInOut, &VoidEventInfo)
	  SINK_ENTRY_INFO(0, DIID_HTMLTextContainerEvents2, DISPID_HTMLELEMENTEVENTS2_ONPASTE, OnRealPaste, &EventInfo)
      SINK_ENTRY_INFO(0, DIID_HTMLTextContainerEvents2, DISPID_HTMLELEMENTEVENTS2_ONDRAGEND, OnDrop, &VoidEventInfo)
  END_SINK_MAP()

  LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
  LRESULT OnFocus(UINT, WPARAM, LPARAM, BOOL &);

  // editing commands
  LRESULT ExecCommand(int cmd);
  void QueryStatus(OLECMD *cmd, int ncmd);
  CString QueryCmdText(ULONG cmd);

  LRESULT OnUndo(WORD, WORD, HWND hWnd, BOOL &);
  LRESULT OnRedo(WORD, WORD, HWND, BOOL &);
  LRESULT OnCut(WORD, WORD, HWND, BOOL &);
  LRESULT OnCopy(WORD, WORD, HWND, BOOL &);
  LRESULT OnPaste(WORD, WORD, HWND, BOOL &);
  LRESULT OnBold(WORD, WORD, HWND, BOOL &);
  LRESULT OnItalic(WORD, WORD, HWND, BOOL &);
  LRESULT OnStrik(WORD, WORD, HWND, BOOL &);
  LRESULT OnSup(WORD, WORD, HWND, BOOL &);
  LRESULT OnSub(WORD, WORD, HWND, BOOL &);
  LRESULT OnCode(WORD, WORD, HWND, BOOL &);


  LRESULT OnFind(WORD, WORD, HWND, BOOL &);
  LRESULT OnFindNext(WORD, WORD, HWND, BOOL &);
  LRESULT OnReplace(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleLink(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleFootnote(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleNolink(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleNormal(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleTextAuthor(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleSubtitle(WORD, WORD, HWND, BOOL &);
  LRESULT OnStyleCode(WORD, WORD, HWND, BOOL &);
  LRESULT OnViewHTML(WORD, WORD, HWND, BOOL &);
  LRESULT OnSelectElement(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddBody(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddBodyNotes(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddSection(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddAnn(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddTitle(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddSubtitle(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddPoem(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddStanza(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddCite(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddEpigraph(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddTA(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditAddImage(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditInsImage(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditClone(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditMerge(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditSplit(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditRemoveOuter(WORD, WORD, HWND, BOOL &);
  LRESULT OnSaveImageAs(WORD, WORD, HWND, BOOL &);
  LRESULT OnEditInsertTable(WORD wNotifyCode, WORD wID, HWND hWndCtl);

  bool CheckCommand(WORD wID);
  bool CheckSetCommand(WORD wID);
  bool canApplyCommand(const CString name);

  // Searching
  bool CanFindNext();
  void CancelIncSearch();
  void StartIncSearch();
 bool DoIncSearch(const CString &str, bool fMore);
  bool DoSearch(bool fMore = true);
  bool DoSearchStd(bool fMore = true);
  bool DoSearchRegexp(bool fMore = true);
  void DoReplace();
  int GlobalReplace(MSHTML::IHTMLElementPtr elem = NULL, CString cntTag = L"P");
  CString LastSearchPattern();
  int ReplaceAllRe(const CString &re, const CString &str, MSHTML::IHTMLElementPtr elem = NULL, CString cntTag = L"P");

  int ToolWordsGlobalReplace(MSHTML::IHTMLElementPtr fbw_body, int *pIndex = NULL, int *globIndex = NULL, bool find = false, CString cntTag = L"P");



  int ReplaceToolWordsRe(const CString &re,
                         const CString &str,
                         MSHTML::IHTMLElementPtr fbw_body,
                         bool replace = false,
                         CString cntTag = L"P",
                         int *pIndex = NULL,
                         int *globIndex = NULL,
                         int replNum = 0);

  // searching in scintilla
  bool SciFindNext(HWND src, bool fFwdOnly, bool fBarf);

  void AddImage(const CString &filename);
  CString AttachImage(const CString &filename);

  // utilities
  CString SelPath();
  void GoTo(MSHTML::IHTMLElementPtr e, bool fScroll = true);
  void CFBEView::ScrollToSelection();
  void CFBEView::ScrollToElement(MSHTML::IHTMLElementPtr e);
  void CFBEView::ApplyConfChanges();
  MSHTML::IHTMLElementPtr SelectionContainer();

  MSHTML::IHTMLElementPtr SelectionsStruct(const CString search);
  MSHTML::IHTMLElementPtr SelectionsStructB(const CString search, _bstr_t &style);
  // bool SelectionHasTags(wchar_t *elem);
  MSHTML::IHTMLElementPtr SelectionAnchor(MSHTML::IHTMLElementPtr cur);
//  MSHTML::IHTMLElementPtr SelectionStructNearestCon();

  void Normalize();
  MSHTML::IHTMLDOMNodePtr GetChangedNode();
  void ImgSetURL(IDispatch *elem, const CString &url);

  bool SplitContainer(bool fCheck);
  //  MSHTML::IHTMLDOMNodePtr	  ChangeAttribute(MSHTML::IHTMLElementPtr elem, const wchar_t* attrib, const wchar_t* value);
  bool InsertTable(bool fCheck, bool bTitle = true, int nrows = 1);
  long InsertCode();
  bool GoToFootnote(bool fCheck);
  bool GoToReference(bool fCheck);
  MSHTML::IHTMLDOMRangePtr SetSelection(MSHTML::IHTMLElementPtr begin, MSHTML::IHTMLElementPtr end, int begin_pos, int end_pos);
  void GetRealCharPos1(MSHTML::IHTMLDOMNodePtr& node, int& pos);
  int CountNodeChars(MSHTML::IHTMLDOMNodePtr node);

  // script calls
  IDispatchPtr Call(const wchar_t *name);
  bool bCall(const wchar_t *name, int nParams, VARIANT *params);
  bool bCall(const wchar_t *name);
  IDispatchPtr Call(const wchar_t *name, IDispatch *pDisp);
  bool bCall(const wchar_t *name, IDispatch *pDisp);
  bool IsCommandPossible(const wchar_t *name);

  // change notifications
  void EditorChanged(int id);
  bool m_trace_changes;
  void TraceChanges(bool state);
  MSHTML::IHTMLChangeLogPtr m_pChangeLog;
  // external helper
  static IDispatchPtr CreateHelper();

  // DWebBrowserEvents2
  void __stdcall OnDocumentComplete(IDispatch *pDisp, VARIANT *vtUrl);
  void __stdcall OnBeforeNavigate(IDispatch *pDisp, VARIANT *vtUrl, VARIANT *vtFlags,
                                  VARIANT *vtTargetFrame, VARIANT *vtPostData,
                                  VARIANT *vtHeaders, VARIANT_BOOL *fCancel);

  // HTMLDocumentEvents2
  void __stdcall OnSelChange(IDispatch *evt);
  VARIANT_BOOL __stdcall OnContextMenu(IDispatch *evt);
  VARIANT_BOOL __stdcall OnClick(IDispatch *evt);
  VARIANT_BOOL __stdcall OnKeyDown(IDispatch *evt);
  void __stdcall OnFocusInOut(IDispatch *evt);

  // HTMLTextContainerEvents2
  VARIANT_BOOL __stdcall OnRealPaste(IDispatch *evt);
  void __stdcall OnDrop(IDispatch *);
  VARIANT_BOOL __stdcall OnDragDrop(IDispatch *);

  // form changes
  bool IsSynched();
  void Synched(bool state);

  // extract currently selected text
  CString GetSelection();
  bool CloseFindDialog(CFindDlgBase *dlg);
  bool CloseFindDialog(CReplaceDlgBase *dlg);
  bool CloseReplaceDialog(CReplaceDlgBase *dlg);

  // added by SeNS
  long m_elementsNum;
  bool IsHTMLChanged();

private:
  // added by SeNS
  int m_startMatch, m_endMatch;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FBEVIEW_H__E0C71279_419D_4273_93E3_57F6A57C7CFE__INCLUDED_)
