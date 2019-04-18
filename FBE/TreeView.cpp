#include "stdafx.h"
#include "resource.h"
#include "res1.h"

#include "utils.h"

#include "FBEView.h"
#include "FBDoc.h"
#include "TreeView.h"

extern CElementDescMnr _EDMnr;

// redrawing the tree is _very_ ugly visually, so we first build a copy and compare them
struct TreeNode 
{
  TreeNode    *parent,*next,*last,*child; // link to neightborhoods
  CString     text; // node title
  int	      img; // image index
  MSHTML::IHTMLElement* pos; //corresponding HTML-element
  
  TreeNode() : parent(0), next(0), child(0), last(0), img(0), pos(0) 
  {
  }
  
  TreeNode(TreeNode *nn, const CString& s, int ii, MSHTML::IHTMLElement* pp) : 
                    parent(nn), next(0), last(0), child(0), text(s), img(ii), pos(pp)
  {
    if (pos)
      pos->AddRef();
  }
  
  ~TreeNode() 
  {
    if (pos)
      pos->Release();
    TreeNode *q;
    for (TreeNode *n=child; n ;n=q) 
    {
      q = n->next;
      delete n;
    }
  }
    ///<summary>Append new node</summary>
    ///<params name="s">Node title</params>
    ///<params name="ii">Image index</params>
    ///<params name="p">HTML-element</params>
    ///<returns>New Node</returns>
  TreeNode *Append(const CString& s, int ii, MSHTML::IHTMLElement* p) 
  {
    TreeNode *n = new TreeNode(this,s,ii,p);
    if (last) 
    {
      last->next = n;
      last = n;
    } 
    else
      last = child = n;
    return n;
  }
};

///<summary>Recursively search tree item with specified HTML-element</summary>
///<params name="ret">Founded tree item or NULL(OUT)</params>
///<params name="ii">Start tree item</params>
///<params name="p">Searched HTML-element</params>
///<returns>true if found</returns>
static bool SearchUnder(CTreeItem& ret, CTreeItem ii, MSHTML::IHTMLElementPtr p) 
{
  CTreeItem jj(ii.GetChild());
  
  while (jj) 
  {
	MSHTML::IHTMLElementPtr n = (MSHTML::IHTMLElement *)jj.GetData();
    if (n && n->contains(p)) 
    {
      ret = jj;
      if (jj.HasChildren())
        // search in children
        SearchUnder(ret, jj, p);
      return true;
    }
    jj = jj.GetNextSibling();
  }
  return false;
}

///<summary>Recursively create node and descendants from root HTML-element</summary>
///<params name="parent">Start tree node</params>
///<params name="elem">root HTML-element</params>
static void MakeNode(TreeNode* parent, MSHTML::IHTMLDOMNodePtr elem)
{
	if (elem->nodeType != NODE_ELEMENT)
		return;

	MSHTML::IHTMLElementPtr he = elem;

	CElementDescriptor* ED = nullptr;
	if(_EDMnr.GetElementDescriptor(he, &ED) && ED->ViewInTree())
	{
        // set node text = title or class name
		CString txt = ED->GetTitle(he);
		U::NormalizeInplace(txt);
		if (txt.IsEmpty())
		{
			txt.SetString(he->className);
			txt = L"<" + txt + L">";
		}
        // append node
		parent = parent->Append(txt, ED->GetDTImageID(), he);
	}
    
    // recursive make nodes
    elem = elem->firstChild;
	while(elem)
	{
		MakeNode(parent, elem);
		elem = elem->nextSibling;
	}

	return;
}

///<summary>Recursively create tree from BODY</summary>
///<params name="view">HTML Document</params>
///<returns>Root tree node</returns>
static TreeNode *GetDocTree(MSHTML::IHTMLDocument2Ptr& view)
{
  TreeNode	*root = new TreeNode();
  try 
  {
    MakeNode(root, view->body);
  }
  catch (_com_error&) 
  {
  }
  
  // if not added - rollback
  if (!root->child) 
  {
    delete root;
    return nullptr;
  }
  return root;
}

///<summary>Recursively transfer data from tree node to TreeItems</summary>
///<params name="n">Root tree node</params>
///<params name="ii">Root tree item</params>
///<params name="fRedraw">Flag/ Set if redraw needs (OUT)</params>
static void  CompareTreesAndSet(TreeNode *n, CTreeItem ii, bool& fRedraw) 
{
  bool	fH1 = (ii == TVI_ROOT) ? ii.m_pTreeView->GetCount() > 0 : ii.HasChildren() != 0;
  
  // walk them one by one and check
  CTreeItem   nc = fH1 ? ii.GetChild() : CTreeItem(NULL,ii.m_pTreeView);
  TreeNode    *ic = n->child;
  CString     text = L"";
  while (ic && nc) 
  {
    int	 img1, img2;
    nc.GetImage(img1, img2);
    nc.GetText(text);
    
    // copy different
    if (text != ic->text || img1 != ic->img) 
    { 
      // copy tree node to treeitem
      if (!fRedraw) 
      {
        ii.m_pTreeView->SetRedraw(FALSE);
        fRedraw = true;
      }
      // set title & icon
      nc.SetImage(ic->img, ic->img);
      nc.SetText(ic->text);
    } 
    // set HTML-Element link
    MSHTML::IHTMLElementPtr od = (MSHTML::IHTMLElement *)nc.GetData();
    if (od)
      od->Release();
    if (ic->pos)
      ic->pos->AddRef();
    nc.SetData((LPARAM)ic->pos);
    // recursive call
    CompareTreesAndSet(ic, nc, fRedraw);

    ic = ic->next;
    nc = nc.GetNextSibling();
  }
  
  if ((nc || ic) && !fRedraw) 
  {
    ii.m_pTreeView->SetRedraw(FALSE);
    fRedraw = true;
  }

  // remove extra children, staring with nc
  while (nc) 
  { 
    CTreeItem next = nc.GetNextSibling();
    nc.Delete();
    nc = next;
  }
  
  // append children to ii
  while (ic) 
  { 
    nc = ii.AddTail(ic->text,ic->img);
    if (ic->pos)
      ic->pos->AddRef();
    nc.SetData((LPARAM)ic->pos);
    if (ic->child)
      CompareTreesAndSet(ic,nc,fRedraw);
    ic = ic->next;
  }
}

///<summary>Recursively expand all child nodes</summary>
///<params name="n">Strat tree item</params>
///<params name="fEnable">Redraw flag (OUT)</params>
static void RecursiveExpand(CTreeItem n, bool& fEnable) 
{
  for (CTreeItem ch(n.GetChild()); ch; ch = ch.GetNextSibling()) 
  {
    if (!ch.HasChildren())
      continue;
    if (!(ch.GetState(TVIS_EXPANDED)&TVIS_EXPANDED)) 
    {
      if (!fEnable) 
      {
        fEnable = true;
        ch.GetTreeView()->SetRedraw(FALSE);
      }
      ch.Expand();
    }
    RecursiveExpand(ch,fEnable);
  }
}

///<summary>Locate tree item with specified HTML-element</summary>
///<params name="p">HTML-element</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::LocatePosition(MSHTML::IHTMLElementPtr p) 
{
  CTreeItem ret(TVI_ROOT, this);
  if (GetCount() == 0 || !p)
    return ret; // no items at all

  SearchUnder(ret, ret, p);
  return ret;
}

///<summary>Highlight tree item with specified HTML-element</summary>
///<params name="p">HTML-element</params>
void CTreeView::HighlightItemAtPos(MSHTML::IHTMLElementPtr p) 
{
  CTreeItem ii(LocatePosition(p));
  // if already selected - return
  if (ii == m_last_lookup_item)
    return;

  m_last_lookup_item = ii;
  ClearSelection();
 
  if (ii != TVI_ROOT) 
    SelectItem(ii);	
  else
    SelectItem(0);
}

///<summary>Build tree</summary>
///<params name="view">HTML Document</params>
void  CTreeView::GetDocumentStructure(MSHTML::IHTMLDocument2Ptr& view) 
{
  m_last_lookup_item = 0;

  // Build tree internal structure
  TreeNode  *root = GetDocTree(view);
  if (!root) 
  {
    SetRedraw(FALSE);
    DeleteAllItems();
    SetRedraw(TRUE);
    return;
  }

  //Refill tree info from internal structure
  bool	fRedraw = false;
  CompareTreesAndSet(root, CTreeItem(TVI_ROOT,this), fRedraw);
  if (fRedraw)
    SetRedraw(TRUE);

  delete root;
}

///<summary>Update tree</summary>
void CTreeView::UpdateAll()
{
	::SendMessage(m_main_window, WM_COMMAND, MAKEWPARAM(0,IDN_TREE_UPDATE_ME), (LPARAM)m_hWnd);
}

void  CTreeView::UpdateDocumentStructure(MSHTML::IHTMLDocument2Ptr& v,MSHTML::IHTMLDOMNodePtr node) 
{
  MSHTML::IHTMLElementPtr ce(node);
  CTreeItem  ii(LocatePosition(ce));

  // shortcut for the most common situation
  // update whole tree
  if (ii == TVI_ROOT) 
  { // huh?
    GetDocumentStructure(v);
    return;
  }

  MSHTML::IHTMLElementPtr jj = (MSHTML::IHTMLElement *)ii.GetData();

  // all changes confined to a P, which is not in the title
  if (U::scmp(jj->tagName, L"DIV")==0 && U::scmp(node->nodeName,L"P")==0) 
  {
        MSHTML::IHTMLElementPtr tn(U::FindTitleNode(jj));
        /* old variant
        if (!tn || (tn !=ce && tn->contains(ce) != VARIANT_TRUE))
            return;
        if (tn && (tn == ce || tn->contains(ce)==VARIANT_TRUE))
        {
            CString txt = tn->innerText;
            U::NormalizeInplace(txt);
            ii.SetText(txt);
            return;
        }
        */
        
        if (tn && (tn == ce || tn->contains(ce)==VARIANT_TRUE))
        {
            CString txt = tn->innerText;
            U::NormalizeInplace(txt);
            ii.SetText(txt);
        }
  }

  TreeNode *nn = new TreeNode;

  MakeNode(nn,jj);
  
  bool	fRedraw = false;
  // check the item itself

  // TO_DO ?? May be Check CompareTreesAndSet(nn, ii, fRedraw);
  CString text;
  int	  img1,img2;
  ii.GetImage(img1,img2);
  ii.GetText(text);
  if (text != nn->child->text || img1 != nn->child->img) 
  { 
    // copy differed item here
    SetRedraw(FALSE);
    fRedraw = true;
    ii.SetImage(nn->child->img, nn->child->img);
    ii.SetText(nn->child->text);
  }

  CompareTreesAndSet(nn->child, ii, fRedraw);
  if (fRedraw)
    SetRedraw(TRUE);

  delete nn;
}

BOOL CTreeView::PreTranslateMessage(MSG* pMsg)
{
  pMsg; //???
  return FALSE;
}

LRESULT CTreeView::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
  // "CTreeViewCtrl::OnCreate()"
  LRESULT lRet = DefWindowProc(uMsg, wParam, lParam);
  
  // "OnInitialUpdate"
  m_ImageList.CreateFromImage(IDB_STRUCTURE,16,32,RGB(255,0,255),IMAGE_BITMAP);
  SetImageList(m_ImageList,TVSIL_NORMAL);

  SetScrollTime(1);
  FillEDMnr();
  
  bHandled = TRUE;
  
  return lRet;
}

LRESULT CTreeView::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
  SetImageList(NULL,TVSIL_NORMAL);
  m_ImageList.Destroy();
  
  // Say that we didn't handle it so that the treeview and anyone else
  //  interested gets to handle the message
  bHandled = FALSE;
  return 0;
}

LRESULT CTreeView::OnClick(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled = SetMultiSelection(wParam, lParam);
  // check if we are going to hit an item
  /*UINT	  flags=0;
  CTreeItem ii(HitTest(CPoint(LOWORD(lParam),HIWORD(lParam)),&flags));
  if (flags&TVHT_ONITEM && !ii.IsNull()) { // try to select and expand it
    CTreeItem	sel(GetSelectedItem());
    if (sel!=ii)
      ii.Select();
    if (!(ii.GetState(TVIS_EXPANDED)&TVIS_EXPANDED))
      ii.Expand();
    SetFocus();
  } //else*/
    //bHandled=FALSE;
  return 0;
}

LRESULT CTreeView::OnDblClick(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
  // check if we double-clicked an already selected item item
  UINT  flags = 0;
  CTreeItem ii(HitTest(CPoint(LOWORD(lParam),HIWORD(lParam)),&flags));
  if (flags&TVHT_ONITEM && ii && ii == GetSelectedItem())
    ::SendMessage(m_main_window,WM_COMMAND,MAKEWPARAM(0,IDN_TREE_CLICK),(LPARAM)m_hWnd);
  return 0;
}

LRESULT CTreeView::OnChar(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
  switch (wParam) 
  {
  case VK_RETURN: // swallow
    break;
  case '*': // expand the entire tree
    if (GetCount()>0) 
    {
      bool fEnable = false;
      RecursiveExpand(CTreeItem(TVI_ROOT,this), fEnable);
      if (fEnable)	 
        SetRedraw(TRUE);
    }
    break;
  default: // pass to control
    break;
  }
  bHandled=FALSE;
  return 0;
}
// Onkey down - select handling
LRESULT CTreeView::OnKeyDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{  
  if (wParam==VK_RETURN) // expand & goto element
    ::PostMessage(m_main_window,WM_COMMAND,MAKEWPARAM(0,IDN_TREE_RETURN),(LPARAM)m_hWnd);

  if ((wParam==VK_UP || wParam==VK_DOWN) && (GetKeyState(VK_SHIFT) & 0x8000))
	{
		// Initialize the reference item if this is the first shift selection
		if(!m_hItemFirstSel)
		{
			m_hItemFirstSel = GetSelectedItem(); 
			ClearSelection();
		}

		HTREEITEM hItemPrevSel = GetSelectedItem(); //currently selected item

		HTREEITEM hItemNext; // closest item in tree
		if (wParam==VK_UP)
			hItemNext = GetPrevVisibleItem(hItemPrevSel);
		else
			hItemNext = GetNextVisibleItem(hItemPrevSel);

		if (hItemNext)
		{
			/* Old variant
			// Determine if we need to reselect previously selected item
			BOOL bReselect = !(GetItemState(hItemNext, TVIS_SELECTED) & TVIS_SELECTED );

			// Select the next item - this will also deselect the previous item
                SelectItem(hItemNext);
            
			// Reselect the previously selected item
			if ( bReselect )
				SetItemState( hItemPrevSel, TVIS_SELECTED, TVIS_SELECTED );
            */
            SelectItems(m_hItemFirstSel, hItemNext, false);
		}
		bHandled = TRUE;
		return 0;
	}
	else if (wParam >= VK_SPACE)
	{
        // reset selection if press any alphanum key
		m_hItemFirstSel = nullptr;
		ClearSelection();
	}
  
  bHandled=FALSE;
  return 0;
}
// TO_DO No DRAG-DROP features
LRESULT CTreeView::OnBegindrag(int idCtrl, LPNMHDR mhdr, BOOL& bHandled)
{
	LPNMTREEVIEW pnmtv = (LPNMTREEVIEW) mhdr;
	
	m_move_from.m_hTreeItem = pnmtv->itemNew.hItem;	
	
	m_himlDrag = TreeView_CreateDragImage(*this, pnmtv->itemNew.hItem);
	ImageList_BeginDrag(this->m_himlDrag, 0, 0, 0);
	ImageList_DragEnter(*this, pnmtv->ptDrag.x,pnmtv->ptDrag.y);
    SetCapture();
	m_drag = true;
	bHandled = false;
	return 0;
}

LRESULT CTreeView::OnLButtonUp(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	if(m_drag)
	{
		EndDrag();
		::SendMessage(m_main_window,WM_COMMAND,MAKEWPARAM(0,IDN_TREE_MOVE_ELEMENT),(LPARAM)m_hWnd);
	}
	bHandled = false;
	return 0;
}

LRESULT CTreeView::OnMouseMove(UINT, WPARAM, LPARAM lParam, BOOL& bHandled)
{
	if(m_drag)
	{
		POINTS Pos = MAKEPOINTS(lParam);
		POINT point;
		point.x = Pos.x;
		point.y = Pos.y;

		ImageList_DragMove(Pos.x, Pos.y);
		UINT		flags;
		CTreeItem	hitem(this->HitTest(point, &flags), this);		
		if (!hitem.IsNull())
		{
			if(IsDropChangePosition(flags, hitem))
			{		
				CImageList::DragShowNolock(FALSE);
				if(IsParent(m_move_from, hitem))
				{
					SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);	
					GetItemImage(hitem, m_drop_item_nimage, m_drop_item_nimage);
					m_insert_type = CTreeView::none;	
					SelectItem(hitem);
				}
                else if(flags & TVHT_ONITEMICON)
				{
					SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);
					GetItemImage(hitem, m_drop_item_nimage, m_drop_item_nimage);
					SetItemImage(hitem, m_drop_item_nimage + 1, m_drop_item_nimage + 1);
					m_insert_type = CTreeView::sibling;
					SelectItem(hitem);
				}
				else if(flags & TVHT_ONITEMLABEL)
				{
					SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);
					GetItemImage(hitem, m_drop_item_nimage, m_drop_item_nimage);
					SetItemImage(hitem, m_drop_item_nimage + 2, m_drop_item_nimage + 2);				
					m_insert_type = CTreeView::child;
					SelectItem(hitem);
				}
				else
				{
					SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);	
					GetItemImage(hitem, m_drop_item_nimage, m_drop_item_nimage);
					m_insert_type = CTreeView::none;										
				}
				m_move_to.m_hTreeItem = hitem;
				
				CImageList::DragShowNolock(TRUE);
			}
			else
			{
				//SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);
			}		
		}
	}
	
	bHandled = false;
	return 0;	
}

LRESULT CTreeView::OnRClick(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
{
	if(m_drag)
		EndDrag();

	UINT flags = 0;
	CTreeItem htItem(HitTest(CPoint(LOWORD(lParam),HIWORD(lParam)),&flags), this);
	if(!htItem.GetState(TVIS_SELECTED))
	{
		ClearSelection();
		SelectItem(htItem);
	}
	SendMessage(WM_CONTEXTMENU, (WPARAM) m_hWnd, GetMessagePos());	
    return 0;
 }

LRESULT CTreeView::OnContextMenu(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
	CPoint ptMousePos = (CPoint)lParam;
		
	// if Shift-F10
	if (ptMousePos.x == -1 && ptMousePos.y == -1)
	{
		ptMousePos = (CPoint)GetMessagePos();		
	}

	ScreenToClient(&ptMousePos);

	UINT uFlags;
	CTreeItem htItem(HitTest( ptMousePos, &uFlags ), this);

	if( htItem.IsNull() )
		return 0;

	//m_hActiveItem = htItem;

	HMENU menu;
	HMENU pPopup;

	// the font popup is stored in a resource
	menu = ::LoadMenu(_Module.GetResourceInstance(), MAKEINTRESOURCEW(IDR_DOCUMENT_TREE));
	//TO_DO Disable menu items if multiple selection
	pPopup = ::GetSubMenu(menu, 0);
	ClientToScreen(&ptMousePos);
	::TrackPopupMenu(pPopup, TPM_LEFTALIGN, ptMousePos.x, ptMousePos.y, 0, *this, 0);

	return 1;
}

LRESULT CTreeView::OnDeleteItem(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
    NMTREEVIEW *tvn = (NMTREEVIEW*)pnmh;
    if (tvn->itemOld.lParam)
      ((MSHTML::IHTMLElement*)tvn->itemOld.lParam)->Release();
    return 0;
}

bool CTreeView::IsDropChangePosition(UINT flags, HTREEITEM hitem)
{
	if(hitem != m_move_to.m_hTreeItem)
		return true;

	if(flags & TVHT_ONITEMICON) 
		if(m_insert_type != CTreeView::sibling)
			return true;
		else
			return false;

	if(flags & TVHT_ONITEMLABEL) 
		if (m_insert_type != CTreeView::child)
			return true;
		else
			return false;

	if (m_insert_type != CTreeView::none)
		return true;
	else
		return false;
}

bool CTreeView::IsParent(CTreeItem	parent, CTreeItem child)
{
	if(parent == child)
		return true;
	while((child = child.GetParent()) && !child.IsNull())
	{
        if(parent == child)
			return true;
	}
	return false;
}

bool CTreeView::IsSibling(CTreeItem	item, CTreeItem	sibling)
{
	return (IsOneOfNextSibling(item, sibling) || IsOneOfPrevSibling(item, sibling));
}

bool CTreeView::IsOneOfNextSibling(CTreeItem	item, CTreeItem	sibling)
{
	if(item == sibling)
		return true;
	while((sibling = sibling.GetNextSibling()) && !sibling.IsNull())
	{
        if(item == sibling)
			return true;
	}
	return false;
}
bool CTreeView::IsOneOfPrevSibling(CTreeItem	item, CTreeItem	sibling)
{
	if(item == sibling)
		return true;
	while((sibling = sibling.GetPrevSibling()), !sibling.IsNull())
	{
        if(item == sibling)
			return true;
	}
	return false;
}

void CTreeView::EndDrag()
{
	ImageList_DragLeave(*this);
	ImageList_EndDrag();
	ReleaseCapture();  	
    
	SetItemImage(m_move_to, m_drop_item_nimage, m_drop_item_nimage);
	SelectDropTarget(NULL);
	m_drag = false;	        			
}

LRESULT CTreeView::OnCut(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// remove last selection
	SetItemState(m_move_from, 0, TVIS_CUT);
	HTREEITEM hitem = GetSelectedItem();
	m_move_from = hitem;
	SetItemState(hitem, TVIS_CUT, TVIS_CUT);	
	return 0;	
}
LRESULT CTreeView::OnPaste(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// remove selection
	SetItemState(m_move_from, 0, TVIS_CUT);
	m_move_to = GetSelectedItem();
	m_insert_type = CTreeView::child;
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT),(LPARAM)m_hWnd);
	return 0;
}
LRESULT CTreeView::OnView(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{	
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_VIEW_ELEMENT),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnViewSource(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_VIEW_ELEMENT_SOURCE),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnDelete(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_DELETE_ELEMENT),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnRight(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	//m_move_from = GetSelectedItem();
	
	// ???? ????????? ??????? ?????? ?????? ????????
	HTREEITEM from = GetSelectedItem();
	HTREEITEM item = GetSelectedItem();// = TreeView_GetNextSibling(*this, m_move_from);
	HTREEITEM prevItem = 0;
	while(item = TreeView_GetNextSibling(*this, item))
	{
		prevItem = item;
	}

	item = prevItem;

	m_move_to = GetSelectedItem();
	m_insert_type = CTreeView::child;
	while(item = TreeView_GetPrevSibling(*this, item))
	{
		if(prevItem == from)
			break;

		m_move_from = prevItem;
		::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT),(LPARAM)m_hWnd);
		prevItem = item;
	}

	m_move_from = from;
	m_move_to = TreeView_GetPrevSibling(*this, m_move_from);

	if(m_move_to)
	{
		item = TreeView_GetPrevSibling(*this, m_move_from);
		if(!item || ! m_move_from)
			return 0;

		if(TreeView_GetChild(*this, item))
		{
			item = TreeView_GetChild(*this, item);
			while (item)
			{		
				m_move_to = item;
				item = TreeView_GetNextSibling(*this, item);		
			}
			m_insert_type = CTreeView::sibling;
		}
		else
		{
			m_move_to = item;
			m_insert_type = CTreeView::child;
		}
	}
	else
	{

	}
	
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnRightOne(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{	
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT_ONE),(LPARAM)m_hWnd);
	/*HTREEITEM item = GetFirstSelectedItem();
	if(!item)
		return 0;

	do
	{
		MoveRightOne(item);		
	}while(item = GetNextSelectedItem(item));*/
	
	return 0;
}

LRESULT CTreeView::OnRightSmart(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT_SMART),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnLeftOne(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_LEFT_ONE),(LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnLeftWithChildren(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_LEFT),(LPARAM)m_hWnd);
	return 0;
}
LRESULT CTreeView::OnMerge(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window, WM_COMMAND, MAKELONG(0, IDN_TREE_MERGE), (LPARAM)m_hWnd);
	return 0;
}

LRESULT CTreeView::OnLeft(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_LEFT),(LPARAM)m_hWnd);
	/*m_move_from = GetSelectedItem();
	m_move_to = TreeView_GetParent(*this, m_move_from);
	m_insert_type = InsertType::sibling;
	::SendMessage(m_main_window,WM_COMMAND,MAKELONG(0,IDN_TREE_MOVE_ELEMENT),(LPARAM)m_hWnd);*/
	return 0;
}

bool CTreeView::SetMultiSelection(UINT nFlags, CPoint point)
{
	// Set focus to control if key strokes are needed.
	// Focus is not automatically given to control on lbuttondown

	//m_dwDragStart = GetTickCount();
	
	UINT flag;
	HTREEITEM hItem = HitTest(point, &flag);

	if (nFlags & MK_SHIFT)
	{
		// Shift key is down	

		// Initialize the reference item if this is the first shift selection
		if(!m_hItemFirstSel)
			m_hItemFirstSel = GetSelectedItem();

		// Select new item
		if(GetSelectedItem() == hItem)
			SelectItem( NULL );			// to prevent edit
//		OnLButtonDown(nFlags, point);

		if(m_hItemFirstSel)
		{
			SelectItems(m_hItemFirstSel, hItem, !(BOOL)(nFlags & MK_CONTROL));
			return true;
		}
	}
	else if(nFlags & MK_CONTROL )
	{
		// Control key is down
		if( hItem )
		{			
			// Toggle selection state
			UINT uNewSelState =
				GetItemState(hItem, TVIS_SELECTED) & TVIS_SELECTED ?
							0 : TVIS_SELECTED;

			// Get old selected (focus) item and state
			//HTREEITEM hItemOld = GetSelectedItem();
			//UINT uOldSelState  = hItemOld ?	GetItemState(hItemOld, TVIS_SELECTED) : 0;

			// Select new item
			if( GetSelectedItem() == hItem )
				SelectItem( NULL );		// to prevent edit				

			// Set proper selection (highlight) state for new item
			SelectItem(hItem);
			SetItemState(hItem, uNewSelState,  TVIS_SELECTED);

			// Restore state of old selected item
			//if (hItemOld && hItemOld != hItem)
			//	SetItemState(hItemOld, uOldSelState, TVIS_SELECTED);

			m_hItemFirstSel = NULL;

			return true;
		}
	}
	else
	{
		// Normal - remove all selection and let default 
		// handler do the rest
		ClearSelection();
		//SelectItem(hItem);
		if(flag & TVHT_ONITEMICON || flag & TVHT_ONITEMLABEL)
			TreeView_SetItemState(*this, hItem, TVIS_SELECTED, TVIS_SELECTED);
		m_hItemFirstSel = NULL;
	}
	//		OnLButtonDown(nFlags, point);
	return false;
}
// clear all selections in Tree
void CTreeView::ClearSelection()
{
	// This can be time consuming for very large trees 
	// and is called every time the user does a normal selection
	// If performance is an issue, it may be better to maintain 
	// a list of selected items
	for (CTreeItem hItem(GetRootItem(), this); hItem; hItem = GetNextItem(hItem))
		if (GetItemState(hItem, TVIS_SELECTED) & TVIS_SELECTED)
			SetItemState(hItem, 0, TVIS_SELECTED);
}

///<summary>Selects items from hItemFrom to hItemTo. 
///Does not select child item if parent is collapsed. 
///Removes selection from all other items</summary>
///<params name="hItemFrom">item to start selecting from</params>
///<params name="hItemTo">item to end selection at</params>
///<params name="clearPrevSelection">true=clear selection upto the first item</params>
///<returns>TRUE on success, FALSE - item not visible</returns>
BOOL CTreeView::SelectItems(HTREEITEM hItemFrom, HTREEITEM hItemTo, bool clearPrevSelection)
{
	HTREEITEM hItem = GetRootItem();

	// Clear selection upto the first item
	while ( hItem && hItem != hItemFrom && hItem != hItemTo )
	{
		hItem = GetNextVisibleItem(hItem);
		if(clearPrevSelection)
			SetItemState( hItem, 0, TVIS_SELECTED );
	}	

	if (!hItem)
		return FALSE;	// start Item is not visible

	SelectItem(hItemTo);

	// Rearrange hItemFrom and hItemTo so that hItemFirst is at top
	if( hItem == hItemTo )
	{
		hItemTo = hItemFrom;
		hItemFrom = hItem;
	}


	// Go through remaining visible items
	BOOL bSelect = TRUE;
	while ( hItem )
	{
		// Select or remove selection depending on whether item
		// is still within the range.
		SetItemState( hItem, bSelect ? TVIS_SELECTED : 0, TVIS_SELECTED );

		// Do we need to start removing items from selection
		if( hItem == hItemTo )
			bSelect = FALSE;

		hItem = GetNextVisibleItem( hItem );
	}

	return TRUE;
}

///<summary>Get first selected tree item</summary>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetFirstSelectedItem()
{
	for (CTreeItem hItem(GetRootItem(), this); !hItem.IsNull(); hItem = GetNextItem(hItem))
		if (GetItemState(hItem, TVIS_SELECTED) & TVIS_SELECTED)
			return hItem;

	return nullptr;
}

///<summary>Get last selected tree item</summary>
///<params name="view">HTML Document</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetLastSelectedItem()
{
	CTreeItem ret;
	for (CTreeItem hItem(GetRootItem(), this); !hItem.IsNull(); hItem = GetNextItem(hItem))
		if (GetItemState( hItem, TVIS_SELECTED ) & TVIS_SELECTED)
			ret = hItem;

	return ret;
}

///<summary>Get next selected tree item</summary>
///<params name="hItem">Current item</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetNextSelectedItem(CTreeItem hItem)
{
	for ( hItem = GetNextItem(hItem); !hItem.IsNull(); hItem = GetNextItem(hItem) )
		if (GetItemState( hItem, TVIS_SELECTED ) & TVIS_SELECTED)
			return hItem;

	return NULL;
}

///<summary>Get previous selected tree item</summary>
///<params name="hItem">Current item</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetPrevSelectedItem( CTreeItem hItem )
{
	for ( hItem = GetPrevItem( hItem ); !hItem.IsNull(); hItem = GetPrevItem( hItem ) )
		if ( GetItemState( hItem, TVIS_SELECTED ) & TVIS_SELECTED )
			return hItem;

	return NULL;
}

///<summary>Get next tree item</summary>
///<params name="hitem">Current item</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetNextItem(CTreeItem hitem)
{
	//????????? ???? ?? child
	CTreeItem child = hitem.GetChild();
	if(!child.IsNull())
		return child;

	// ????????? ???? ?? sibling
	CTreeItem sibling = hitem.GetNextSibling();
	if(!sibling.IsNull())
		return sibling;
	
	// ???? ?????? sibling parent'?
	CTreeItem parent = hitem;
	while(parent = parent.GetParent())
	{
		CTreeItem parent_sibling = parent.GetNextSibling();
		if(parent_sibling)
			return parent_sibling;
	}

	return 0;
}

///<summary>Get previous tree item</summary>
///<params name="hitem">Current item</params>
///<returns>Tree item</returns>
CTreeItem CTreeView::GetPrevItem(CTreeItem hitem)
{
	// ????????? ???? ?? prev sibling
	CTreeItem sibling = hitem.GetPrevSibling();
	if(!sibling.IsNull())
	{
		// ?????????? ?????? ?????????? ??????
		CTreeItem next_item = sibling;
		CTreeItem ret = sibling;
		while(next_item = GetNextItem(next_item))
		{
			if(next_item == hitem)
				return ret;

			ret = next_item;
		}
		return 0;
	}

	// ????????? ???? ?? parent
	CTreeItem parent = hitem.GetParent();
	if(!parent.IsNull())
	{
		return parent;
	}

	return 0;
}

CTreeItem CTreeView::GetNextSelectedSibling(CTreeItem item)
{
	if(item.IsNull())
		return 0;

	CTreeItem next = GetNextSelectedItem(item);
	while(!next.IsNull())
	{
		if(IsSibling(item, next))
			return next;
		next = GetNextSelectedItem(next);
	}
	return 0;
}

// search and tree item with specific HTML-element
void CTreeView::SelectElement(MSHTML::IHTMLElement *p)
{
	CTreeItem item = this->LocatePosition(p);
	SelectItem(item);
}

// search and expand tree item with specific HTML-element
void CTreeView::ExpandElem(MSHTML::IHTMLElement *p, UINT mode)
{
	CTreeItem item = this->LocatePosition(p);
	Expand(item, mode);
}

void CTreeView::Collapse(CTreeItem item, int level2Collapse, bool mode)
{
	if(item.IsNull())
		item = GetRootItem();// root ? ??? ??? 0 - ?? ???????
    
	do
	{
		if(level2Collapse == 0)
		{
			item.Expand(mode ? TVE_EXPAND : TVE_COLLAPSE);
		}
		else
		{
			CTreeItem child = item.GetChild();
			if(!child.IsNull())
				Collapse(child, level2Collapse - 1, mode);
		}
		item = item.GetNextSibling();
	} while(!item.IsNull());
}

//TO_DO Not used
void CTreeView::FillEDMnr()
{
	/*m_bodyED = new CBodyED;
	m_sectionED = new CSectionED;
	m_imageED =new CImageED;
	m_poemED = new CPoemED;
	_EDMnr.AddElementDescriptor(m_bodyED);
	_EDMnr.AddElementDescriptor(m_sectionED);
	_EDMnr.AddElementDescriptor(m_imageED);
	_EDMnr.AddElementDescriptor(m_poemED);*/
}

// add image to image list
// return Image index
int CTreeView::AddImage(HANDLE img)
{
	return m_ImageList.Add((HBITMAP)img);
}

// add icon to image list
int CTreeView::AddIcon(HANDLE icon)
{
	return m_ImageList.AddIcon((HICON)icon);
}
