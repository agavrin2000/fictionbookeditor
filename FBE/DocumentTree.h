#pragma once

#include "TreeView.h"

typedef CWinTraits<WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_LEFT, WS_EX_CLIENTEDGE> CTreeWithToolBarWinTraits;

class CTreeWithToolBar : public CFrameWindowImpl<CTreeWithToolBar, CWindow, CTreeWithToolBarWinTraits>
{
private:
	WTL::CToolBarCtrl m_toolbar;
	//WTL::CToolBarCtrl m_element_browser_tb;
	WTL::CReBarCtrl m_rebar;
	int m_maxTbwidth;

	CCommandBarCtrl m_view_bar;
	CMenu m_st_menu;
	BOOL ModifyStyle(DWORD dwRemove, DWORD dwAdd, UINT nFlags = 0) throw();
	void FillViewBar();
public:
	CTreeView m_tree;
	DECLARE_FRAME_WND_CLASS(L"DocumentTreeFrame", IDR_DOCUMENT_TREE)

	void GetDocumentStructure(MSHTML::IHTMLDocument2Ptr v);
	void UpdateDocumentStructure(MSHTML::IHTMLDocument2Ptr& v, MSHTML::IHTMLDOMNodePtr node);
	void HighlightItemAtPos(MSHTML::IHTMLElement* p);
	CTreeItem GetSelectedItem();

	BEGIN_MSG_MAP(CTreeWithToolBar)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SIZE, OnSize)

		COMMAND_ID_HANDLER(ID_DT_RIGHT_ONE, ForwardWMCommand)
		COMMAND_ID_HANDLER(ID_DT_RIGHT_SMART, ForwardWMCommand)
		COMMAND_ID_HANDLER(ID_DT_LEFT, ForwardWMCommand)
		COMMAND_ID_HANDLER(ID_DT_MERGE, ForwardWMCommand)
		COMMAND_ID_HANDLER(ID_DT_DELETE, ForwardWMCommand)

		// Menu commands
		COMMAND_RANGE_HANDLER(IDC_TREE_BASE, IDC_TREE_BASE + 99, OnMenuCommand)
		COMMAND_RANGE_HANDLER(IDC_TREE_ST_BASE, IDC_TREE_ST_BASE + 99, OnMenuStCommand)
		COMMAND_ID_HANDLER(IDC_TREE_CLEAR_ALL, OnMenuClear)
		COMMAND_ID_HANDLER(IDC_TREE_UNCHECK_ALL, OnMenuUncheckAll)

		NOTIFY_CODE_HANDLER(TTN_GETDISPINFO, OnToolTipText)
	END_MSG_MAP()
	
	LRESULT OnCreate(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSize(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT ForwardWMCommand(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	LRESULT OnMenuCommand(WORD, WORD wID, HWND, BOOL&);
	LRESULT OnMenuStCommand(WORD, WORD wID, HWND, BOOL&);
	LRESULT OnMenuClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnMenuUncheckAll(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnToolTipText(int idCtrl, LPNMHDR pnmh, BOOL&);
};

class CDocumentTree : public CPaneContainer
{
private:	
	//int m_current_tab;

public:
	CTreeWithToolBar m_tree;
	CDocumentTree(){};

	// TreeView methods
	void GetDocumentStructure(MSHTML::IHTMLDocument2Ptr v);
	void UpdateDocumentStructure(MSHTML::IHTMLDocument2Ptr v,MSHTML::IHTMLDOMNodePtr node);
	void HighlightItemAtPos(MSHTML::IHTMLElementPtr p);
	CTreeItem GetSelectedItem();

	BEGIN_MSG_MAP(CDocumentTree)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		CHAIN_MSG_MAP(CPaneContainer);
	END_MSG_MAP()
	
	LRESULT OnCreate(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
};
