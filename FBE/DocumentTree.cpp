#include "stdafx.h"
#include "resource.h"
#include "res1.h"
#include "utils.h"
#include "DocumentTree.h"

#include "ElementDescriptor.h"
#include "Settings.h"

extern CSettings _Settings;

extern CElementDescMnr _EDMnr;
//==================================================================================================================
// CTreeWithToolBar - Tree view window with toolbar

BOOL CTreeWithToolBar::ModifyStyle(DWORD dwRemove, DWORD dwAdd, UINT nFlags) throw()
{
	ATLASSERT(::IsWindow(m_hWnd));

	DWORD dwStyle = ::GetWindowLong(m_hWnd, GWL_STYLE);
	DWORD dwNewStyle = (dwStyle & ~dwRemove) | dwAdd;
	if(dwStyle == dwNewStyle)
		return FALSE;

	::SetWindowLong(m_hWnd, GWL_STYLE, dwNewStyle);
	if(nFlags != 0)
	{
		::SetWindowPos(m_hWnd, NULL, 0, 0, 0, 0,
			SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | nFlags);
	}

	return TRUE;
}

LRESULT CTreeWithToolBar::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	//m_toolbarOrientation = CTreeWithToolBar::bottom;

	LRESULT lRet = DefWindowProc(uMsg, wParam, lParam);

	m_tree.Create(*this, rcDefault);
	m_tree.SetBkColor(::GetSysColor(COLOR_WINDOW));
	m_rebar = CFrameWindowImplBase<>::CreateSimpleReBarCtrl(*this, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CS_HREDRAW);
	m_toolbar = CFrameWindowImplBase<>::CreateSimpleToolBarCtrl(*this, IDR_DOCUMENT_TREE, FALSE, ATL_SIMPLE_TOOLBAR_PANE_STYLE);
	CFrameWindowImplBase<>::AddSimpleReBarBandCtrl(m_rebar, m_toolbar);

    // Init standard element descriptors
	_EDMnr.InitStandartEDs();
	int edsCount = _EDMnr.GetStEDsCount();
	for(int i = 0; i < edsCount; ++i)
	{
		CElementDescriptor* eld = _EDMnr.GetStED(i);
		eld->SetViewInTree(_Settings.GetDocTreeItemState(eld->GetCaption(), eld->ViewInTree()));
	}

    // Init nonstandard element descriptors
	_EDMnr.InitScriptEDs();

    //Fill and add menu bar
	RECT rect;
	SetRect(&rect, 0, 0, 500, 20);
	this->ModifyStyle(0, WS_POPUP, 0);
	m_view_bar.Create(*this, rect, NULL, ATL_SIMPLE_TOOLBAR_PANE_STYLE);	
	m_view_bar.SetStyle(ATL_SIMPLE_TOOLBAR_PANE_STYLE);
	FillViewBar();
	this->ModifyStyle(WS_POPUP, 0, 0);

	m_maxTbwidth = 1000;
	::MoveWindow(m_rebar, 0, 0, m_maxTbwidth, 0, true);

//	m_toolbar.HideButton(ID_DT_DELETE);
	bHandled = FALSE;
	return lRet;
}

LRESULT CTreeWithToolBar::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled=FALSE;
	return 0;
}

LRESULT CTreeWithToolBar::OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{    
	return 0;
}

LRESULT CTreeWithToolBar::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	RECT clientRect = {0, 0, 0, 0};
	RECT rebarRect = {0, 0, 0, 0};
	RECT treeRect = {0, 0, 0, 0};
	RECT viewBarRect = {0, 0, 0, 0};
	
	this->GetClientRect(&clientRect);
	::GetWindowRect(m_toolbar, &rebarRect);
	::GetWindowRect(m_view_bar, &viewBarRect);
	int rebarHight = rebarRect.bottom - rebarRect.top + (/*GetSystemMetrics(SM_CYDLGFRAME) + */GetSystemMetrics(SM_CYEDGE))*2;
	rebarRect.left = treeRect.left = clientRect.left;
	rebarRect.right = treeRect.right = clientRect.right;

	int viewBarHight = viewBarRect.bottom - viewBarRect.top;
	int viewBarWidth = viewBarRect.right - viewBarRect.left;

	viewBarRect.left = clientRect.left;
	viewBarRect.right = clientRect.left + viewBarWidth;

	bool moved = false;

    if(rebarRect.top != clientRect.bottom - rebarHight)
        moved = true;

    rebarRect.top = clientRect.bottom - rebarHight;
    
    rebarRect.bottom = rebarRect.top + rebarHight;
    treeRect.top = clientRect.top + viewBarHight;
    treeRect.bottom = rebarRect.top;

    viewBarRect.top = clientRect.top;
    viewBarRect.bottom = viewBarRect.top + viewBarHight;

	if((rebarRect.right - rebarRect.left) > m_maxTbwidth || moved)
	{
		::MoveWindow(m_rebar, rebarRect.left, rebarRect.top, rebarRect.right - rebarRect.left, rebarRect.bottom - rebarRect.top, true);
		::MoveWindow(m_view_bar, viewBarRect.left, viewBarRect.top, viewBarRect.right - viewBarRect.left, viewBarRect.bottom - viewBarRect.top, true);
		m_maxTbwidth = rebarRect.right - rebarRect.left;
	}
	
	::MoveWindow(m_tree, treeRect.left, treeRect.top, treeRect.right - treeRect.left, treeRect.bottom - treeRect.top, true);	

	return 0;
}

LRESULT CTreeWithToolBar::OnToolTipText(int idCtrl, LPNMHDR pnmh, BOOL&)
{
    LPNMTTDISPINFOW pDispInfo = (LPNMTTDISPINFOW)pnmh;
    pDispInfo->szText[0] = 0;

    if((idCtrl != 0) && !(pDispInfo->uFlags & TTF_IDISHWND))
    {
        const int cchBuff = 256;
        wchar_t szBuff[cchBuff];
        szBuff[0] = 0;
        int nRet = ::LoadStringW(ModuleHelper::GetResourceInstance(), idCtrl, szBuff, cchBuff);
        for(int i = 0; i < nRet; i++)
        {
            if(szBuff[i] == L'\n')
            {
                SecureHelper::strncpyW_x(pDispInfo->szText, _countof(pDispInfo->szText), &szBuff[i + 1], _TRUNCATE);
                break;
            }
        }
        if(nRet > 0)   // string was loaded, save it
            pDispInfo->uFlags |= TTF_DI_SETITEM;
    }

    return 0;
}

void CTreeWithToolBar::GetDocumentStructure(MSHTML::IHTMLDocument2Ptr v)
{
	m_tree.GetDocumentStructure(v);
}

void CTreeWithToolBar::UpdateDocumentStructure(MSHTML::IHTMLDocument2Ptr& v,MSHTML::IHTMLDOMNodePtr node)
{
	m_tree.UpdateDocumentStructure(v, node);
}

void CTreeWithToolBar::HighlightItemAtPos(MSHTML::IHTMLElement *p)
{
	m_tree.HighlightItemAtPos(p);
}

CTreeItem CTreeWithToolBar::GetSelectedItem()
{
	return m_tree.GetSelectedItem();
}

LRESULT CTreeWithToolBar::ForwardWMCommand(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	DWORD wParam = MAKELONG(wID, wNotifyCode);
	DWORD lParam = (LPARAM)hWndCtl;
	return ::SendMessage(m_tree, WM_COMMAND, wParam, lParam);
}
// fill main menu for tree (descriptors)
void CTreeWithToolBar::FillViewBar()
{
	m_st_menu = ::CreateMenu();	// standard descroiptors
	HMENU menu = ::CreateMenu();// nonstandard descroiptors (scripts)
	HMENU bar = ::CreateMenu();

	wchar_t elsMenuItem[MAX_LOAD_STRING + 1];
	wchar_t scriptsMenuItem[MAX_LOAD_STRING + 1];

	::LoadString(_Module.GetResourceInstance(), IDS_DOCTREE_MENU_ELEMENTS, elsMenuItem, MAX_LOAD_STRING);
	::LoadString(_Module.GetResourceInstance(), IDS_DOCTREE_MENU_SCRIPTS, scriptsMenuItem, MAX_LOAD_STRING);

	::AppendMenu(bar, MF_POPUP|MF_STRING, (UINT)(HMENU)m_st_menu, elsMenuItem);
	::AppendMenu(bar, MF_POPUP|MF_STRING, (UINT)(HMENU)menu, scriptsMenuItem);

	int picType = 0;
	HANDLE picHandle = 0;
	unsigned int st_count = _EDMnr.GetStEDsCount();
	unsigned int count = _EDMnr.GetEDsCount();
    //append standard descriptors
	for(unsigned int i = 0; i < st_count; ++i)
	{
		::AppendMenu(m_st_menu, MF_STRING, IDC_TREE_ST_BASE + i, _EDMnr.GetStED(i)->GetCaption());
		if(_EDMnr.GetStED(i)->ViewInTree())
			::CheckMenuItem(m_st_menu, IDC_TREE_ST_BASE + i, MF_CHECKED);
	}
    // append "Clear" menu item
	::AppendMenu(m_st_menu, MF_SEPARATOR, 0, 0);
	CString s(L""); 
    s.LoadString(IDS_DOC_TREE_CLEANUP);
	::AppendMenu(m_st_menu, MF_STRING, IDC_TREE_UNCHECK_ALL, s);

    //append nonstandard descriptors
	for(unsigned int i = 0; i < count; ++i)
	{
		::AppendMenu(menu, MF_STRING, IDC_TREE_BASE + i, _EDMnr.GetED(i)->GetCaption());
		CElementDescriptor* ED = _EDMnr.GetED(i);

		if(ED->GetPic(picHandle, picType))
		{
			int imageID = 0;

			switch(picType)
			{
			case 0: //bitmap
				m_view_bar.AddBitmap((HBITMAP)picHandle, IDC_TREE_BASE + i);
				imageID = m_tree.AddImage(picHandle);
				break;
			case 1: //icon
				m_view_bar.AddIcon((HICON)picHandle, IDC_TREE_BASE + i);
				imageID = m_tree.AddIcon(picHandle);
				break;				
			}
			ED->SetImageID(imageID);
		}
	}
    // append "Clear" menu item
	::AppendMenu(menu, MF_SEPARATOR, 0, 0);
	s = L""; 
    s.LoadString(IDS_DOC_TREE_CLEANUP);
	::AppendMenu(menu, MF_STRING, IDC_TREE_CLEAR_ALL, s);

	m_view_bar.AttachMenu(bar);
}

// Clear all standard descriptors
LRESULT CTreeWithToolBar::OnMenuUncheckAll(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	_EDMnr.CleanTree();	
	int st_menu_item_count = m_st_menu.GetMenuItemCount();
	for(int i = 0; i < st_menu_item_count; ++i)
	{
		::CheckMenuItem(m_st_menu, IDC_TREE_ST_BASE + i, MF_UNCHECKED);
	}
	m_tree.UpdateAll();

	return 0;
}

// Clear nonstandard descriptors
LRESULT CTreeWithToolBar::OnMenuCommand(WORD, WORD wID, HWND, BOOL&)
{
	unsigned int index = wID - IDC_TREE_BASE;
	CElementDescriptor* ED = _EDMnr.GetED(index);
	ED->CleanUp();
	ED->ProcessScript();
	ED->SetViewInTree(true);
	m_tree.UpdateAll();

	return 0;
}

// toggle checked state in menu standard descriptor
LRESULT CTreeWithToolBar::OnMenuStCommand(WORD, WORD wID, HWND, BOOL&)
{
	unsigned int index = wID - IDC_TREE_ST_BASE;
	CElementDescriptor* ED = _EDMnr.GetStED(index);
	bool view = ED->ViewInTree();
	view = !view;
	ED->SetViewInTree(view);
	if(view)
		::CheckMenuItem(m_st_menu, wID, MF_CHECKED);
	else
		::CheckMenuItem(m_st_menu, wID, MF_UNCHECKED);
	m_tree.UpdateAll();
	return 0;
}
// clear nonstandard menu
LRESULT CTreeWithToolBar::OnMenuClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
{
	_EDMnr.CleanUpAll();
	return 0;
}

//==================================================================================================================
// CDocumentTree - PaneContainer class to display container window (with title and close button) for TreeView

LRESULT CDocumentTree::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	LRESULT lRet = DefWindowProc(uMsg, wParam, lParam);

	m_tree.Create(*this, rcDefault);
	m_tree.m_tree.SetMainwindow(GetParent());
	this->SetClient(m_tree);

	wchar_t capt[MAX_LOAD_STRING + 1];
	::LoadString(_Module.GetResourceInstance(), IDS_DOCUMENT_TREE_CAPTION, capt, MAX_LOAD_STRING);
	this->SetTitle(capt);    
    
    bHandled=FALSE;
    return lRet;
}

//WS_DLGFRAME  | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | CCS_NODIVIDER | CCS_NOPARENTALIGN | TBSTYLE_TOOLTIPS | TBSTYLE_BUTTON | TBSTYLE_AUTOSIZE

LRESULT CDocumentTree::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled=FALSE;
	return 0;
}

LRESULT CDocumentTree::OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	return 0;
}

void CDocumentTree::GetDocumentStructure(MSHTML::IHTMLDocument2Ptr v)
{
	m_tree.GetDocumentStructure(v);
}

void CDocumentTree::UpdateDocumentStructure(MSHTML::IHTMLDocument2Ptr v,MSHTML::IHTMLDOMNodePtr node)
{
	m_tree.UpdateDocumentStructure(v, node);
}

void CDocumentTree::HighlightItemAtPos(MSHTML::IHTMLElementPtr p)
{
	m_tree.HighlightItemAtPos(p);
}

CTreeItem CDocumentTree::GetSelectedItem()
{
	return m_tree.GetSelectedItem();
}
