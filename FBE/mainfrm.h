// MainFrm.h : interface of the CMainFrame class

#if !defined(AFX_MAINFRM_H__38D356D4_C28B_47B0_A7AD_8C6B70F7F283__INCLUDED_)
#define AFX_MAINFRM_H__38D356D4_C28B_47B0_A7AD_8C6B70F7F283__INCLUDED_

#include "stdafx.h"
#include <dlgs.h>
#include "resource.h"
#include "res1.h"

#include "atlctrlsext.h"

#include "utils.h"

#include "FBEView.h"
#include "FBDoc.h"
#include "TreeView.h"
#include "SettingsViewPage.h"
#include "ContainerWnd.h"
#include <Scintilla.h>
#include <SciLexer.h>
#include "FBE.h"
#include "Words.h"
#include "SearchReplace.h"
#include "DocumentTree.h"
#include "Speller.h"

#if _MSC_VER >= 1000
#pragma once
#pragma warning(disable : 4996)
#endif // _MSC_VER >= 1000
/*
#define MSGFLT_ADD 1
#define MSGFLT_REMOVE 2
*/
typedef CWinTraits<WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_LEFT, WS_EX_CLIENTEDGE> CCustomEditWinTraits;

const int MAX_POPUP_SUGGESTION = 8;

class CCustomEdit : public CWindowImpl<CCustomEdit, CEdit, CCustomEditWinTraits>, public CEditCommands<CCustomEdit>
{
public:
	DECLARE_WND_SUPERCLASS(NULL, CEdit::GetWndClassName())

	CCustomEdit() {}

	BEGIN_MSG_MAP(CCustomEdit)
		MESSAGE_HANDLER(WM_CHAR, OnChar)
		CHAIN_MSG_MAP_ALT(CEditCommands<CCustomEdit>, 1)
	END_MSG_MAP()

	LRESULT OnChar(UINT, WPARAM wParam, LPARAM, BOOL &bHandled);
};

class CCustomStatic : public CWindowImpl<CCustomStatic, CStatic /*,CCustomStaticWinTraits*/>
{
private:
	HFONT m_font;
	bool m_enabled;

public:
	CCustomStatic() : m_font(0), m_enabled(0) {}

	LRESULT OnPaint(UINT, WPARAM wParam, LPARAM, BOOL &);
	void DoPaint(CDCHandle dc);

	void SetFont(HFONT pFont);
	void SetEnabled(bool Enabled);

	BEGIN_MSG_MAP(CCustomStatic)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
	END_MSG_MAP()
};

class CTableToolbarsWindow : public CFrameWindowImpl<CTableToolbarsWindow>,
	public CUpdateUI<CTableToolbarsWindow>
{
public:
	BEGIN_UPDATE_UI_MAP(CTableToolbarsWindow)
	END_UPDATE_UI_MAP()
};

class CCustomSaveDialog : public CFileDialogImpl<CCustomSaveDialog>
{
public:
	HWND m_hDlg;
	CString m_encoding;

	CCustomSaveDialog(BOOL bOpenFileDialog, // TRUE for FileOpen, FALSE for FileSaveAs
		const CString &encoding,
		LPCTSTR lpszDefExt = NULL,
		LPCTSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCTSTR lpszFilter = NULL,
		HWND hWndParent = NULL)
		: CFileDialogImpl<CCustomSaveDialog>(bOpenFileDialog, lpszDefExt, lpszFileName, dwFlags, lpszFilter, hWndParent),
		m_hDlg(NULL)
	{
		m_ofn.lpTemplateName = MAKEINTRESOURCE(IDD_CUSTOMSAVEDLG);
		m_encoding = encoding;
	}

	BEGIN_MSG_MAP(CCustomSaveDialog)
		if (uMsg == WM_INITDIALOG)
			return OnInitDialog(hWnd, uMsg, wParam, lParam);

	MESSAGE_HANDLER(WM_SIZE, OnSize);

	CHAIN_MSG_MAP(CFileDialogImpl<CCustomSaveDialog>)
	END_MSG_MAP()

	LRESULT OnInitDialog(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
	BOOL OnFileOK(LPOFNOTIFY on);
};

class CMainFrame : public CFrameWindowImpl<CMainFrame>,
	public CCustomizableToolBarCommands<CMainFrame>,
	public CUpdateUI<CMainFrame>,
	public CMessageFilter,
	public CIdleHandler
{
public:
	DECLARE_FRAME_WND_CLASS(L"FictionBookEditorFrame", IDR_MAINFRAME);

	CMainFrame();
	~CMainFrame();

	CSciFindDlg *m_sci_find_dlg;
	CSciReplaceDlg *m_sci_replace_dlg;

	// Child windows
	CSplitterWindow m_splitter; // doc tree and views
	CContainerWnd m_view;		// document, description and source
	CDocumentTree m_document_tree; //doc tree
	CWindow m_source; // source editor
	CMultiPaneStatusBarCtrl m_status; // status bar


	CCommandBarCtrl m_MenuBar;	 // menu bar
	CToolBarCtrl m_CmdToolbar;	 // commands toolbar
	CToolBarCtrl m_ScriptsToolbar; // commands toolbar
	CReBarCtrl m_rebar;			   // toolbars

	// links toolbar controls
	CCustomStatic m_id_caption;
	CComboBox m_id_box;
	CCustomEdit m_id; // paragraph ID
	CCustomStatic m_href_caption;
	CComboBox m_href_box;
	CCustomEdit m_href; // link's href
	CCustomStatic m_image_title_caption;
	CComboBox m_image_title_box;
	CCustomEdit m_image_title; // image title
	CCustomStatic m_section_id_caption;
	CComboBox m_section_box;
	CCustomEdit m_section; // section ID

	// table toolbar controls
	CCustomStatic m_table_id_caption;
	CComboBox m_id_table_id_box;
	CCustomEdit m_id_table_id; // Table ID
	CCustomStatic m_id_table_caption;
	CComboBox m_id_table_box;
	CCustomEdit m_id_table; // ID
	CCustomStatic m_table_style_caption;
	CComboBox m_styleT_table_box;
	CCustomEdit m_styleT_table; // style table
	CComboBox m_style_table_box;
	CCustomStatic m_style_caption;
	CCustomEdit m_style_table; // style cell
	CCustomStatic m_colspan_caption;
	CComboBox m_colspan_table_box;
	CCustomEdit m_colspan_table; // colspan
	CCustomStatic m_rowspan_caption;
	CComboBox m_rowspan_table_box;
	CCustomEdit m_rowspan_table; // rowspan
	CCustomStatic m_tr_allign_caption;
	CCustomEdit m_alignTR_table; // tr align
	CComboBox m_alignTR_table_box;
	CCustomStatic m_th_allign_caption;
	CComboBox m_align_table_box;
	CCustomEdit m_align_table; // th,td align
	CCustomStatic m_valign_caption;
	CComboBox m_valign_table_box;
	CCustomEdit m_valign_table; // th,td valign

	CRecentDocumentList m_mru; // MRU list

	FB::Doc *m_doc; // currently open document

	CComPtr<IShellItem> m_spShellItem;

	// IDs in combobox
	bool m_cb_updated : 1;
	bool m_cb_last_images : 1; // images or plain ids?
	bool m_ignore_cb_changes : 1;

	int m_want_focus; // focus this control when idle

	CString m_status_msg; // message to be posted to frame's status line

	bool m_restore_pos_cmdline;

	// incremental search helpers
	CString m_is_str;
	CString m_is_prev;
	int m_incsearch;
	bool m_is_fail;

	// Script structure (either for scripts files or for scripts folders)
	struct ScrInfo
	{
		CString name;
		CString path;
		CString order;
		HANDLE picture;
		int pictType;
		int Type;
		CString id;
		CString refid;
		bool isFolder;
		int wID;
		/*ACCEL accel;*/
	};

	// Script small menu icon (16x16) type
	enum ScrPictType
	{
		NO_PICT,
		BITMAP,
		ICON
	};

	CSimpleArray<ScrInfo> m_scripts;					//array of scripts
	CSimpleMap<unsigned int, HBITMAP> m_scripts_images; //array of icons

	void CollectScripts(CString path, TCHAR *mask, int lastid, CString refid);
	int GrabScripts(CString path, TCHAR *mask, CString refid);
	void AddScriptsSubMenu(HMENU, CString, CSimpleArray<ScrInfo> &);
	void QuickScriptsSort(CSimpleArray<ScrInfo> &, int, int);
	void UpScriptsFolders(CSimpleArray<ScrInfo> &);
	ScrInfo *m_last_script;
	void InitScriptHotkey(CMainFrame::ScrInfo &);

	// toolbars
	bool IsBandVisible(int id);

	// browser controls
	void AttachDocument(FB::Doc *doc);
	CFBEView &ActiveView();

	// document structure
	void GetDocumentStructure();
	void GoTo(MSHTML::IHTMLElement *elem);

	// loading/saving support
	enum FILE_OP_STATUS
	{
		FAIL,
		OK,
		CANCELLED
	};

	CString GetOpenFileName();
	CString GetSaveFileName(CString &encoding);
	bool SaveToFile(const CString &filename);
	bool DiscardChanges();
	FILE_OP_STATUS SaveFile(bool askname);
	FILE_OP_STATUS LoadFile(const wchar_t *initfilename = NULL);
	bool CheckFileTimeStamp();
	bool ReloadFile();
	void UpdateFileTimeStamp();
	unsigned __int64 FileAge(LPCTSTR FileName);
	unsigned __int64 m_file_age;
	CString m_bad_filename;

	// show a specific view
	enum VIEW_TYPE
	{
		BODY,
		DESC,
		SOURCE,
		NEXT
	};
	void ShowView(VIEW_TYPE vt = BODY);
	bool ShowSource(bool saveSelection = true);
	//bool IsSourceActive();

	VIEW_TYPE m_current_view;
	VIEW_TYPE m_last_view;
	VIEW_TYPE m_last_ctrl_tab_view;
	bool m_ctrl_tab;

	MSHTML::IHTMLTxtRangePtr m_body_selection;
	MSHTML::IHTMLTxtRangePtr m_desc_selection;

	MSHTML::IHTMLDOMRangePtr m_body_sel;
	MSHTML::IHTMLDOMRangePtr m_desc_sel;
	void DisplayXMLError(const CString& Msg);

	/*void SaveSelection();
	void RestoreSelection();
	void ClearSelection();*/

	// Plugins support
	CSimpleArray<CLSID> m_import_plugins;
	CSimpleArray<CLSID> m_export_plugins;
	void InitPlugins();
	void InitPluginsType(HMENU hMenu, const TCHAR *type, UINT cmdbase, CSimpleArray<CLSID> &plist);
	void InitPluginHotkey(CString guid, UINT cmd, CString name);
	UINT m_last_plugin;

	// utility function on creating
	void AddTbButton(HWND hWnd, const TCHAR *text, const int idCommand = 0, const BYTE bState = 0, const HICON icon = 0);
	void AddStaticText(CCustomStatic &st, HWND toolbarHwnd, int id, const TCHAR *text, HFONT hFont);
	void FillMenuWithHkeys(HMENU);

	// ui updating
	void UIUpdateViewCmd(CFBEView &view, WORD wID, OLECMD &oc, const TCHAR *hk);
	void UIUpdateViewCmd(CFBEView &view, WORD wID);
	void UISetCheckCmd(CFBEView &view, WORD wID);

	void StopIncSearch(bool fCancel);
	void SetIsText();

	// source editor
	void DefineMarker(int marker, int markerType, COLORREF fore, COLORREF back);
	void SetupSci();
	// source folding
	void FoldAll();
	void ExpandFold(int &line, bool doExpand, bool force = false, int visLevels = 0, int level = -1);
	void SciCollapse(int level2Collapse, bool mode);
	// source editor styles
	void SetSciStyles();
	void SciModified(const SCNotification &scn);
	void SciMarginClicked(const SCNotification &scn);
	bool SciUpdateUI(bool gotoTag);

	// void ChangeNBSP(MSHTML::IHTMLElementPtr elem);
	void CMainFrame::ChangeNBSP(MSHTML::IHTMLDOMNodePtr node);

	void RemoveLastUndo();

	// source<->html exchange
	int BuildXMLPath(MSXML2::IXMLDOMElementPtr selectedElement);
	bool SourceToHTML();
	bool HTMLToSource();

	int m_selBandID;

	// message handlers
	virtual BOOL PreTranslateMessage(MSG *pMsg);
	virtual BOOL OnIdle();

	BEGIN_UPDATE_UI_MAP(CMainFrame)
		// ui windows
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST + 1, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST + 2, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST + 3, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST + 4, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ATL_IDW_BAND_FIRST + 5, UPDUI_MENUPOPUP)

		// view commands
		UPDATE_ELEMENT(ID_VIEW_STATUS_BAR, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ID_VIEW_TREE, UPDUI_MENUPOPUP)
		UPDATE_ELEMENT(ID_VIEW_DESC, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_BODY, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_SOURCE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_ZOOMIN, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_ZOOMOUT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_ZOOM100, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)

		// editing commands
		UPDATE_ELEMENT(ID_EDIT_UNDO, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_REDO, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_CUT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_COPY, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_PASTE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_FIND, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_FINDNEXT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_REPLACE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_GOTO_FOOTNOTE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_GOTO_MATCHTAG, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_CLONE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_SPLIT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_MERGE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_EDIT_REMOVE_OUTER, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		
		//add command
		UPDATE_ELEMENT(ID_ADD_TITLE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_BODY, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_BODY_NOTES, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_ANN, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_SECTION, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_POEM, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_STANZA, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_EPIGRAPH, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_CITE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_INS_TABLE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_IMAGE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ATTACH_BINARY, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)

		// style commands
		UPDATE_ELEMENT(ID_STYLE_NORMAL, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_SUBTITLE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_ADD_TEXTAUTHOR, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_CODE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_LINK, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_NOTE, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_NOLINK, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_BOLD, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_ITALIC, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_STRIKETROUGH, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_SUPERSCRIPT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_STYLE_SUBSCRIPT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)

		UPDATE_ELEMENT(ID_TOOLS_WORDS, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_VIEW_OPTIONS, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_TOOLS_SPELLCHECK, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)
		UPDATE_ELEMENT(ID_LAST_SCRIPT, UPDUI_MENUPOPUP | UPDUI_TOOLBAR)

	END_UPDATE_UI_MAP()

	BEGIN_MSG_MAP(CMainFrame)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(U::WM_POSTCREATE, OnPostCreate)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
		MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)

		// added by SeNS: toolbar customization menu
		MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)

		MESSAGE_HANDLER(WM_THEMECHANGED, OnSettingChange)
		MESSAGE_HANDLER(WM_DROPFILES, OnDropFiles)
		MESSAGE_HANDLER(U::WM_SETSTATUSTEXT, OnSetStatusText)
		MESSAGE_HANDLER(U::WM_TRACKPOPUPMENU, OnTrackPopupMenu)

		// incremental search support
		MESSAGE_HANDLER(WM_CHAR, OnChar)
		MESSAGE_HANDLER(WM_COMMAND, OnPreCommand)
		MESSAGE_HANDLER(WM_SIZE, OnSize)

		// tree view notifications
		COMMAND_CODE_HANDLER(IDN_TREE_CLICK, OnTreeClick)
		COMMAND_CODE_HANDLER(IDN_TREE_RETURN, OnTreeReturn)
		COMMAND_CODE_HANDLER(IDN_TREE_MOVE_ELEMENT, OnTreeMoveElement)
		COMMAND_CODE_HANDLER(IDN_TREE_MOVE_ELEMENT_ONE, OnTreeMoveElementOne)
		COMMAND_CODE_HANDLER(IDN_TREE_MOVE_LEFT, OnTreeMoveLeftElement)
		COMMAND_CODE_HANDLER(IDN_TREE_MOVE_ELEMENT_SMART, OnTreeMoveElementSmart)
		COMMAND_CODE_HANDLER(IDN_TREE_VIEW_ELEMENT, OnTreeViewElement)
		COMMAND_CODE_HANDLER(IDN_TREE_VIEW_ELEMENT_SOURCE, OnTreeViewElementSource)
		COMMAND_CODE_HANDLER(IDN_TREE_DELETE_ELEMENT, OnTreeDeleteElement)
		COMMAND_CODE_HANDLER(IDN_TREE_MERGE, OnTreeMerge)
		COMMAND_CODE_HANDLER(IDN_TREE_UPDATE_ME, OnTreeUpdate)
		COMMAND_CODE_HANDLER(IDN_TREE_RESTORE, OnTreeRestore)

		// file menu
		COMMAND_ID_HANDLER(ID_APP_EXIT, OnFileExit)
		COMMAND_ID_HANDLER(ID_FILE_NEW, OnFileNew)
		COMMAND_ID_HANDLER(ID_FILE_OPEN, OnFileOpen)
		COMMAND_ID_HANDLER(ID_FILE_SAVE, OnFileSave)
		COMMAND_ID_HANDLER(ID_FILE_SAVE_AS, OnFileSaveAs)
		COMMAND_ID_HANDLER(ID_FILE_VALIDATE, OnFileValidate)
		COMMAND_RANGE_HANDLER(ID_FILE_MRU_FIRST, ID_FILE_MRU_LAST, OnFileOpenMRU)

		// Source editor
		COMMAND_RANGE_HANDLER(ID_SCI_COLLAPSE1, ID_SCI_COLLAPSE9, OnSciCollapse)
		COMMAND_RANGE_HANDLER(ID_SCI_EXPAND1, ID_SCI_EXPAND9, OnSciExpand)

		// edit menu
		COMMAND_ID_HANDLER(ID_EDIT_INCSEARCH, OnEditIncSearch)
		COMMAND_ID_HANDLER(ID_ATTACH_BINARY, OnEditAddBinary)
		COMMAND_ID_HANDLER(ID_EDIT_FIND, OnEditFind)
		COMMAND_ID_HANDLER(ID_EDIT_FINDNEXT, OnEditFind)
		COMMAND_ID_HANDLER(ID_EDIT_REPLACE, OnEditFind)
		COMMAND_RANGE_HANDLER(ID_EDIT_INS_SYMBOL, ID_EDIT_INS_SYMBOL + 100, OnEditInsSymbol)

		// zoom notifications
		COMMAND_ID_HANDLER(ID_VIEW_ZOOMIN, OnZoomIn)
		COMMAND_ID_HANDLER(ID_VIEW_ZOOMOUT, OnZoomOut)
		COMMAND_ID_HANDLER(ID_VIEW_ZOOM100, OnZoom100)



		// popup menu (speller addons)
		COMMAND_ID_HANDLER(IDC_SPELL_IGNOREALL, OnSpellIgnoreAll)
		COMMAND_ID_HANDLER(IDC_SPELL_ADD2DICT, OnSpellAddToDict)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 1, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 2, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 3, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 4, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 5, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 6, OnSpellReplace)
		COMMAND_ID_HANDLER(IDC_SPELL_REPLACE + 7, OnSpellReplace)

		COMMAND_ID_HANDLER(ID_DOC_CHANGED, OnDocChanged)

		// view menu
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 1, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 2, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 3, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 4, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 5, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 6, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 7, OnViewToolBar)
		COMMAND_ID_HANDLER(ATL_IDW_BAND_FIRST + 8, OnViewToolBar)
		COMMAND_ID_HANDLER(ID_VIEW_STATUS_BAR, OnViewStatusBar)
		COMMAND_ID_HANDLER(ID_VIEW_TREE, OnViewTree)
		COMMAND_ID_HANDLER(ID_VIEW_DESC, OnViewDesc)
		COMMAND_ID_HANDLER(ID_VIEW_BODY, OnViewBody)
		COMMAND_ID_HANDLER(ID_VIEW_SOURCE, OnViewSource)
		COMMAND_ID_HANDLER(ID_VIEW_OPTIONS, OnViewOptions)

		// tools menu
		COMMAND_ID_HANDLER(ID_TOOLS_WORDS, OnToolsWords)
		COMMAND_ID_HANDLER(ID_TOOLS_OPTIONS, OnToolsOptions)
		COMMAND_ID_HANDLER(ID_TOOLS_CUSTOMIZE, OnToolCustomize)

		//scripts
		COMMAND_RANGE_HANDLER(ID_SCRIPT_BASE, ID_SCRIPT_BASE + 999, OnToolsScript)
		COMMAND_ID_HANDLER(ID_LAST_SCRIPT, OnLastScript)

		//spellcheck
		COMMAND_ID_HANDLER(ID_TOOLS_SPELLCHECK, OnSpellCheck);
	COMMAND_ID_HANDLER(ID_TOOLS_SPELLCHECK_HIGHLIGHT, OnToggleHighlight);

	// help menu
	COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		// navigation commands
		COMMAND_ID_HANDLER(ID_SELECT_TREE, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_ID, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_HREF, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_IMAGE, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_TEXT, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_NEXT_ITEM, OnNextItem)
		COMMAND_ID_HANDLER(ID_SELECT_SECTION, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_IDT, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_STYLET, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_STYLE, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_COLSPAN, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_ROWSPAN, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_ALIGNTR, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_ALIGN, OnSelectCtl)
		COMMAND_ID_HANDLER(ID_SELECT_VALIGN, OnSelectCtl)

		// editor notifications
		COMMAND_CODE_HANDLER(IDN_SEL_CHANGE, OnEdSelChange)
		COMMAND_CODE_HANDLER(IDN_ED_CHANGED, OnEdChange)
		COMMAND_CODE_HANDLER(IDN_ED_TEXT, OnEdStatusText)
		COMMAND_CODE_HANDLER(IDN_WANTFOCUS, OnEdWantFocus)
		COMMAND_CODE_HANDLER(IDN_ED_RETURN, OnEdReturn)
		COMMAND_CODE_HANDLER(IDN_NAVIGATE, OnNavigate)
		COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnEdKillFocus)
		COMMAND_CODE_HANDLER(CBN_EDITCHANGE, OnCbEdChange)
		COMMAND_CODE_HANDLER(CBN_SELENDOK, OnCbSelEndOk)
		COMMAND_HANDLER(IDC_HREF, CBN_SETFOCUS, OnCbSetFocus)

		// source code editor notifications
		NOTIFY_CODE_HANDLER(SCN_MODIFIED, OnSciModified)
		NOTIFY_CODE_HANDLER(SCN_MARGINCLICK, OnSciMarginClick)
		NOTIFY_CODE_HANDLER(SCN_UPDATEUI, OnSciUpdateUI)
		NOTIFY_CODE_HANDLER(SCN_PAINTED, OnSciPainted)

		// tree pane
		COMMAND_ID_HANDLER(ID_PANE_CLOSE, OnViewTree)

		// FBEview calls to process messages without FBEview focused
		COMMAND_ID_HANDLER_EX(ID_GOTO_FOOTNOTE, OnGoToFootnote)
		COMMAND_ID_HANDLER_EX(ID_GOTO_REFERENCE, OnGoToReference)
		COMMAND_ID_HANDLER_EX(ID_GOTO_MATCHTAG, OnGoToMatchTag);
	COMMAND_ID_HANDLER_EX(ID_GOTO_WRONGTAG, OnGoToWrongTag);

	// chain commands to active view
	MESSAGE_HANDLER(WM_COMMAND, OnUnhandledCommand)

		CHAIN_MSG_MAP(CUpdateUI<CMainFrame>)
		CHAIN_MSG_MAP(CFrameWindowImpl<CMainFrame>)
		CHAIN_MSG_MAP(CCustomizableToolBarCommands<CMainFrame>)
	END_MSG_MAP()

	LRESULT OnCreate(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnPostCreate(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnSettingChange(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnSize(UINT, WPARAM, LPARAM, BOOL &bHandled);

	LRESULT OnContextMenu(UINT, WPARAM, LPARAM lParam, BOOL &);
	LRESULT OnUnhandledCommand(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnSetFocus(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnDropFiles(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnSetStatusText(UINT, WPARAM, LPARAM lParam, BOOL &);

	LRESULT OnTrackPopupMenu(UINT, WPARAM, LPARAM lParam, BOOL &);

	LRESULT OnChar(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnPreCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
	LRESULT OnFileExit(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileNew(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileOpen(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileOpenMRU(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileSave(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileSaveAs(WORD, WORD, HWND, BOOL &);
	LRESULT OnFileValidate(WORD, WORD, HWND, BOOL &);
	LRESULT OnToolsImport(WORD, WORD, HWND, BOOL &);
	LRESULT OnToolsExport(WORD, WORD, HWND, BOOL &);

	LRESULT OnEditIncSearch(WORD, WORD, HWND, BOOL &);
	LRESULT OnEditAddBinary(WORD, WORD, HWND, BOOL &);
	LRESULT OnEditFind(WORD, WORD, HWND, BOOL &);
	LRESULT OnEditInsSymbol(WORD, WORD, HWND, BOOL &);

	LRESULT OnSpellReplace(WORD, WORD, HWND, BOOL &);
	LRESULT OnSpellIgnoreAll(WORD, WORD, HWND, BOOL &);
	LRESULT OnSpellAddToDict(WORD, WORD, HWND, BOOL &);
	LRESULT OnDocChanged(WORD, WORD, HWND, BOOL &);

	LRESULT OnViewToolBar(WORD, WORD, HWND, BOOL &);
	LRESULT OnViewStatusBar(WORD, WORD, HWND, BOOL &);
	LRESULT OnViewTree(WORD, WORD, HWND, BOOL &);
	LRESULT OnViewDesc(WORD, WORD, HWND, BOOL &);
	LRESULT OnViewBody(WORD, WORD, HWND, BOOL &);
	LRESULT OnSciShowAllSymbols(WORD, WORD, HWND, BOOL &);

	LRESULT OnViewSource(WORD, WORD, HWND, BOOL &);
	LRESULT OnViewOptions(WORD, WORD, HWND, BOOL &);

	LRESULT OnToolsWords(WORD, WORD, HWND, BOOL &);
	LRESULT OnToolsOptions(WORD, WORD, HWND, BOOL &);
	LRESULT OnToolsScript(WORD, WORD, HWND, BOOL &);

	LRESULT OnHideToolbar(WORD, WORD, HWND, BOOL &);
	LRESULT OnToolCustomize(WORD, WORD, HWND, BOOL &);
	LRESULT OnLastScript(WORD, WORD, HWND, BOOL &);

	LRESULT OnAppAbout(WORD, WORD, HWND, BOOL &);

	LRESULT OnSelectCtl(WORD, WORD, HWND, BOOL &);
	LRESULT OnNextItem(WORD, WORD, HWND, BOOL &);

	LRESULT OnEdSelChange(WORD, WORD, HWND, BOOL &);
	LRESULT OnEdStatusText(WORD, WORD, HWND, BOOL &);
	LRESULT OnEdWantFocus(WORD, WORD, HWND, BOOL &);
	LRESULT OnEdReturn(WORD, WORD, HWND, BOOL &);
	LRESULT OnNavigate(WORD, WORD, HWND, BOOL &);

	LRESULT OnCbSetFocus(WORD, WORD, HWND, BOOL &);

	LRESULT OnEdChange(WORD, WORD, HWND hWnd, BOOL &b);
	LRESULT OnCbEdChange(WORD, WORD, HWND, BOOL &);
	LRESULT OnCbSelEndOk(WORD code, WORD wID, HWND hWnd, BOOL &);
	LRESULT OnEdKillFocus(WORD, WORD, HWND, BOOL &);

	LRESULT OnTreeReturn(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeClick(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeMoveElement(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeMoveElementOne(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeMoveElementSmart(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeMoveLeftElement(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeViewElement(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeViewElementSource(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeDeleteElement(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeMerge(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeUpdate(WORD, WORD, HWND, BOOL &);
	LRESULT OnTreeRestore(WORD, WORD, HWND, BOOL &);

	LRESULT OnGoToFootnote(WORD wNotifyCode, WORD wID, HWND hWndCtl);
	LRESULT OnGoToReference(WORD wNotifyCode, WORD wID, HWND hWndCtl);

	LRESULT OnZoomIn(WORD, WORD, HWND, BOOL &);
	LRESULT OnZoomOut(WORD, WORD, HWND, BOOL &);
	LRESULT OnZoom100(WORD, WORD, HWND, BOOL &);

	LRESULT OnSciModified(int id, NMHDR *hdr, BOOL &bHandled);
	LRESULT OnSciMarginClick(int id, NMHDR *hdr, BOOL &bHandled);
	LRESULT OnSciUpdateUI(int id, NMHDR *hdr, BOOL &bHandled);
	LRESULT OnSciPainted(int id, NMHDR *hdr, BOOL &bHandled);
	LRESULT OnSciCollapse(WORD cose, WORD wID, HWND, BOOL &);
	LRESULT OnSciExpand(WORD cose, WORD wID, HWND, BOOL &);
	LRESULT OnGoToMatchTag(WORD wNotifyCode, WORD wID, HWND hWndCtl);
	LRESULT OnGoToWrongTag(WORD wNotifyCode, WORD wID, HWND hWndCtl);

	LRESULT OnSpellCheck(WORD, WORD, HWND, BOOL &b);
	LRESULT OnToggleHighlight(WORD, WORD, HWND, BOOL &);

	// doc manipulate functions
	MSHTML::IHTMLDOMNodePtr MoveRightElementWithoutChildren(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr MoveRightElement(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr MoveLeftElement(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr RecursiveMoveRightElement(CTreeItem item);
	MSHTML::IHTMLDOMNodePtr GetFirstChildSection(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr GetNextSiblingSection(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr GetPrevSiblingSection(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr GetLastChildSection(MSHTML::IHTMLDOMNodePtr node);
	MSHTML::IHTMLDOMNodePtr CreateNestedSection(MSHTML::IHTMLDOMNodePtr section);
	bool IsNodeSection(MSHTML::IHTMLDOMNodePtr node);
	bool IsEmptySection(MSHTML::IHTMLDOMNodePtr section);

	bool IsEmptyText(BSTR text); //??

	void SourceGoTo(int line, int linePos);
	void GoToSelectedTreeItem();

	bool ShowSettingsDialog(HWND parent = ::GetActiveWindow());
	void ApplyConfChanges();
	void RestartProgram();

	bool BitmapInClipboard();

	// added by SeNS
	void UpdateViewSizeInfo();

private:
	void CMainFrame::DisplayOvrInsStatus();
	int TempCloseDialogs();
	void RestoreDialogs(int dialog_type);
	wchar_t strINS[MAX_LOAD_STRING + 1]; // localized string for INSERT mode
	wchar_t strOVR[MAX_LOAD_STRING + 1]; // localized string for OVERWRITEINSERT mode

	DWORD m_last_tree_update;
	BOOL m_last_sci_ovr : 1;
	bool m_last_ie_ovr : 1;
	bool m_sel_changed : 1;
	bool m_change_state : true;
	bool m_need_title_update : 1;

	bool m_sci_painted; //flag to await sci painted
	CSpeller *m_Speller;
	void DisplayCharCode();
	CStrings *m_strSuggestions;
	void ParaIDsToComboBox(CComboBox &box);
	void BinIDsToComboBox(CComboBox &box);
};

int StartScript(CMainFrame *mainframe);
void StopScript(void);
HRESULT ScriptLoad(const wchar_t *filename);
HRESULT ScriptCall(const wchar_t *func, VARIANT *arg, int argnum, VARIANT *ret);
bool ScriptFindFunc(const wchar_t *func);


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__38D356D4_C28B_47B0_A7AD_8C6B70F7F283__INCLUDED_)