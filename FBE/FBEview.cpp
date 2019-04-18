// FBEView.cpp : implementation of the CFBEView class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "res1.h"

#include "utils.h"

#include "FBEView.h"
#include "SearchReplace.h"
#include "Scintilla.h"
#include "ElementDescriptor.h"
#include "atlimageex.h"
#include "TableDialog.h"

extern CElementDescMnr _EDMnr;

// normalization helpers
static void PackText(MSHTML::IHTMLElement2Ptr elem);
static void KillDivs(MSHTML::IHTMLElement2Ptr elem);
static void FixupParagraphs(MSHTML::IHTMLElement2Ptr elem);
static void RelocateParagraphs(MSHTML::IHTMLDOMNode *node);
static void KillStyles(MSHTML::IHTMLElement2Ptr elem);

/// <summary>Convert VARIANT to boolean</summary>
/// <param name="vt">VARIANT</param>
/// <returns>bool value</returns>
static bool vt2bool(const _variant_t &vt)
{
	if (V_VT(&vt) == VT_DISPATCH)
		return V_DISPATCH(&vt) != 0;
	if (V_VT(&vt) == VT_BOOL)
		return V_BOOL(&vt) == VARIANT_TRUE;
	if (V_VT(&vt) == VT_I4)
		return V_I4(&vt) != 0;
	if (V_VT(&vt) == VT_UI4)
		return V_UI4(&vt) != 0;
	return false;
}

/// <summary>Check empty node.</summary>
/// <param name="node">HTMLNode</param>
/// <returns>true for whitespace string</returns>
static bool IsEmptyNode(MSHTML::IHTMLDOMNodePtr node)
{
	if (node->nodeType != NODE_ELEMENT) //ELEMENT_NODE
		return false;

	CString name = node->nodeName;
	CString class_name = MSHTML::IHTMLElementPtr(node)->className;
	CString test = MSHTML::IHTMLElementPtr(node)->outerHTML;

	// images are always empty
	// links can be meaningful even if the contain only ws
	// empty <P> means <empty-line/>
	if ((name == L"DIV" && class_name == L"image") || name == L"IMG" || name == L"A" || name == L"P")
		return false;

	if (!node->hasChildNodes())
		return true;

	// More than one child
	if (node->firstChild->nextSibling)
		return false;
	
	//Has one child and it is element
	if (node->firstChild->nodeType != NODE_TEXT) //TEXT_ELEMENT
		return false;

	//Has one child and it is nonempty text
	if (U::is_whitespace(node->firstChild->nodeValue.bstrVal))
		return true;

	return false;
}

/// <summary>Remove empty leaf nodes</summary>
/// <param name="node">HTMLNode</param>
static void RemoveEmptyNodes(MSHTML::IHTMLDOMNode *node)
{
	if (node->nodeType != NODE_ELEMENT)
	{
		if (U::is_whitespace(node->nodeValue.bstrVal))
			node->removeNode(VARIANT_TRUE);
		return;
	}

	MSHTML::IHTMLDOMNodePtr cur(node->firstChild);
	while (cur)
	{
		RemoveEmptyNodes(cur);
		if (IsEmptyNode(cur))
			cur->removeNode(VARIANT_TRUE);
		cur = cur->nextSibling;
	}
}

/// <summary>Find closest parent DIV</summary>
/// <param name="hp">HTML-element</param>
/// <returns>HTML-element or NULL</returns>
static MSHTML::IHTMLElementPtr GetHP(MSHTML::IHTMLElementPtr hp)
{
	while (hp && U::scmp(hp->tagName, L"DIV"))
		hp = hp->parentElement;
	return hp;
}

/// <summary>Cleaning up HTML-element. Remove all DIVs</summary>
/// <param name="elem">HTML Element</param>
static void KillDivs(MSHTML::IHTMLElement2Ptr elem)
{
	MSHTML::IHTMLElementCollectionPtr divs(elem->getElementsByTagName(L"DIV"));
	while (divs->length > 0)
		MSHTML::IHTMLDOMNodePtr(divs->item(0))->removeNode(VARIANT_FALSE);
}

/// <summary>Cleaning up styles.
/// Remove all styles from <P></summary>
/// <param name="elem">HTML Element</param>
static void KillStyles(MSHTML::IHTMLElement2Ptr elem)
{
	MSHTML::IHTMLElementCollectionPtr ps(elem->getElementsByTagName(L"P"));
	for (long l = 0; l < ps->length; ++l)
		CheckError(MSHTML::IHTMLElementPtr(ps->item(l))->put_className(NULL));
}

//////////////////////////////////////////////////////////////////////////////
/// @fn static bool	MergeEqualHTMLElements(MSHTML::IHTMLDOMNode *node)
///
/// функция объединяет стоящие рядом одинаковые HTML элементы
///
/// @params MSHTML::IHTMLDOMNode *node [in, out] - нода, внутри которой будет производиться преобразование
///
/// @note сливаются следующие элементы: EM, STRONG
/// при этом пробельные символы, располагающиеся между закрывающем и открывающим тегами остаются, т.е.
/// '<EM>хороший</EM> <EM>пример</EM>' преобразуется в '<EM>хороший пример</EM>'
///
/// @author Ильин Иван @date 31.03.08
//////////////////////////////////////////////////////////////////////////////
static bool MergeEqualHTMLElements(MSHTML::IHTMLDOMNodePtr node, MSHTML::IHTMLDocument2Ptr doc)
{
	if (node->nodeType != NODE_ELEMENT) 
		return false;

	bool fRet = false;

	MSHTML::IHTMLDOMNodePtr cur(node->firstChild);
	while (cur)
	{
		MSHTML::IHTMLDOMNodePtr next;
		next = cur->nextSibling;
		if (MergeEqualHTMLElements(cur, doc))
		{
			cur = node->firstChild;
			continue;
		}

		// если нет следующего элемента, то сливать будет не с чем
		if (!next)
			return false;

		_bstr_t name(cur->nodeName);
		MSHTML::IHTMLElementPtr curelem(cur);
		MSHTML::IHTMLElementPtr nextElem(next);

		if (U::scmp(name, L"EM") == 0 || U::scmp(name, L"STRONG") == 0)
		{
			// отлавливаем ситуацию с пробелом, обрамленным тегами EM т.д.
			bstr_t curText = curelem->innerText;
			if (curText.length() == 0 || U::is_whitespace(curelem->innerText))
			{
				// удаляем обрамляющие теги
				MSHTML::IHTMLDOMNodePtr prev = cur->previousSibling;
				if ((bool)prev)
				{
					if (prev->nodeType == NODE_TEXT) 
					{
						prev->nodeValue = (bstr_t)prev->nodeValue.bstrVal + curelem->innerText;
					}
					else
					{
						MSHTML::IHTMLElementPtr prevElem(prev);
						prevElem->innerHTML = prevElem->innerHTML + curelem->innerText;
					}
					cur->removeNode(VARIANT_TRUE);
					cur = prev;
					continue;
				}

				if ((bool)next)
				{
					MSHTML::IHTMLDOMNodePtr parent = cur->parentNode;
					if (next->nodeType == NODE_TEXT) //text
					{
						next->nodeValue = (bstr_t)curelem->innerText + next->nodeValue.bstrVal;
					}
					else
					{
						MSHTML::IHTMLElementPtr nextElem2(next);
						nextElem2->innerHTML = curelem->innerText + nextElem2->innerHTML;
					}
					cur->removeNode(VARIANT_TRUE);
					cur = parent->firstChild;
					continue;
				}
			}

			if (next->nodeType == NODE_TEXT)
			{
				MSHTML::IHTMLDOMNodePtr afterNext(next->nextSibling);
				if (!(bool)afterNext)
				{
					cur = next;
					continue;
				}

				MSHTML::IHTMLElementPtr afterNextElem(afterNext);

				bstr_t afterNextName = afterNext->nodeName;
				if (U::scmp(name, afterNextName)) // если следующий элемент другого типа
				{
					cur = next;
					continue;
				}

				// проверяем между одинаковыми элементами стоят одни пробелы
				if (!U::is_whitespace(next->nodeValue.bstrVal))
				{
					cur = next;
					continue; // <EM>123</EM>45<EM>678</EM> абсолютно нормальная ситуация
				}

				// объединяем элементы
				MSHTML::IHTMLElementPtr newelem(doc->createElement(name));
				MSHTML::IHTMLDOMNodePtr newnode(newelem);
				newelem->innerHTML = curelem->innerHTML + next->nodeValue.bstrVal + afterNextElem->innerHTML;
				cur->replaceNode(newnode);
				afterNext->removeNode(VARIANT_TRUE);
				next->removeNode(VARIANT_TRUE);
				cur = newnode;
				fRet = true;
			}
			else
			{
				bstr_t nextName(next->nodeName);
				if (U::scmp(name, nextName)) // если следующий элемент другого типа
				{
					cur = next;
					continue;
				}

				// объединяем элементы
				MSHTML::IHTMLElementPtr newelem(doc->createElement(name));
				MSHTML::IHTMLDOMNodePtr newnode(newelem);
				newelem->innerHTML = curelem->innerHTML + nextElem->innerHTML;
				cur->replaceNode(newnode);
				next->removeNode(VARIANT_TRUE);
				cur = newnode;
				fRet = true;
				continue;
			}
		}
		cur = next;
	}
	return fRet;
}

///<summary>Get input value from HTML-input element</summary>
///<params name="m_cur_input_element">INPUT element</params>
///<return>String value</return>
static CString GetInputElementValue(MSHTML::IHTMLElementPtr m_cur_input_element)
{
	CString ret(L"");
	CString tName = m_cur_input_element->tagName;
	if (tName == L"INPUT")
	{
		ret.SetString(MSHTML::IHTMLInputElementPtr(m_cur_input_element)->value);
	}
	if (tName == L"TEXTAREA")
	{
		ret.SetString(MSHTML::IHTMLTextAreaElementPtr(m_cur_input_element)->value);
	}
	if (tName == L"SELECT")
	{
		ret.SetString(MSHTML::IHTMLSelectElementPtr(m_cur_input_element)->value);
	}
	return ret;
}

///<summary> Remove unsupported elements from DOM
/// Recursive function</summary>
///<params name="node">Starting node (first start=fbw_body)</params>
///<return>flag to restart scan after renaming</return>
static bool RemoveUnk(MSHTML::IHTMLDOMNodePtr node, MSHTML::IHTMLDocument2Ptr doc)
{
	if (node->nodeType != NODE_ELEMENT)
		return false;

	// restart:
	MSHTML::IHTMLDOMNodePtr cur = node->firstChild;
	while (cur)
	{
		MSHTML::IHTMLDOMNodePtr next = cur->nextSibling;
		if (RemoveUnk(cur, doc))
		{
			/*		if (RemoveUnk(cur, doc))
					goto restart;
					*/
					// Rename node B=>STRONG? I=>EM
			MSHTML::IHTMLElementPtr curelem(cur);
			CString name = curelem->tagName;
			if (name == L"B" || name == L"I")
			{
				name = (name == L"B") ? L"STRONG" : L"EM";
				MSHTML::IHTMLElementPtr newelem(doc->createElement(name.GetString()));
				newelem->innerHTML = curelem->innerHTML;
				//MSHTML::IHTMLDOMNodePtr newnode(newelem);
				cur = cur->replaceNode(MSHTML::IHTMLDOMNodePtr(newelem));
				//			cur = newnode;
				//			fRet = true;
				//			goto restart; //??
			}

			/*		CString text;
					if (curelem != NULL)
						text.SetString(curelem->outerHTML);*/
			CString cls("");
			CString parent_cls("");

			if (curelem && MSHTML::IHTMLDOMNodePtr(curelem)->nodeType == NODE_ELEMENT)
			{
				cls.SetString(curelem->className);
			}
			if (curelem && curelem->parentElement)
			{
				parent_cls.SetString(curelem->parentElement->className);
			}

			// Skip allowed tags
			if (name != L"P" && name != L"STRONG" && name != L"STRIKE" && name != L"SUP" && name != L"SUB" &&
				name != L"EM" && name != L"A" && !(name == L"SPAN" && cls == L"code") &&
				name != L"BR" && name != L"IMG" && !(name == L"SPAN" && cls == L"image"))
			{

				if (name == L"DIV")
				{
					// check allowed id and class names
					CString id = curelem->id;
					if (cls == L"body" || cls == L"section" || cls == L"table" || cls == L"tr" || cls == L"th" ||
						cls == L"td" || cls == L"output" || cls == L"part" || cls == L"output-document-class" ||
						cls == L"annotation" || cls == L"title" || cls == L"epigraph" || cls == L"poem" ||
						cls == L"stanza" || cls == L"cite" || cls == L"date" || cls == L"history" || cls == L"image" ||
						cls == L"code" || id == L"fbw_desc" || id == L"fbw_body" || id == L"fbw_updater")
					{
						cur = next;
						continue;
					}
				}
				// check in allowed class descriptors
				CElementDescriptor *ED;
				if (_EDMnr.GetElementDescriptor(cur, &ED))
				{
					cur = next;
					continue;
				}

				// remove non-standart node
				MSHTML::IHTMLDOMNodePtr ce(cur->previousSibling);
				cur->removeNode(VARIANT_FALSE); //remove, but keep children
				if (ce)
					next = ce->nextSibling;
				else
					next = node->firstChild;
			}
		}
		cur = next;
	}
	return true;
}

// move the paragraph up one level
static void MoveUp(bool fCopyFmt, MSHTML::IHTMLDOMNodePtr &node)
{
	MSHTML::IHTMLDOMNodePtr parent(node->parentNode);
	MSHTML::IHTMLElement2Ptr elem(parent);

    // move children up
	if (fCopyFmt)
	{
        // clone parent without children (it can be A/EM/STRONG/SPAN)
		MSHTML::IHTMLDOMNodePtr clone(parent->cloneNode(VARIANT_FALSE));

        // move all children to cloned parent
		while (node->firstChild)
			clone->appendChild(node->firstChild);
        // add to current node
		node->appendChild(clone);
	}

	// clone parent once more and move siblings after node to it
	if (node->nextSibling)
	{
		MSHTML::IHTMLDOMNodePtr clone(parent->cloneNode(VARIANT_FALSE));
		while (node->nextSibling)
			clone->appendChild(node->nextSibling);
		elem->insertAdjacentElement(L"afterEnd", MSHTML::IHTMLElementPtr(clone));
//		if (U::scmp(parent->nodeName, L"P") == 0)
//			MSHTML::IHTMLElement3Ptr(clone)->inflateBlock = VARIANT_TRUE;
	}

	// now move node to parent level, the tree may be in some weird state
	node->removeNode(VARIANT_TRUE); // delete from tree
	node = elem->insertAdjacentElement(L"afterEnd", MSHTML::IHTMLElementPtr(node));
}

static void BubbleUp(MSHTML::IHTMLDOMNode *node, const wchar_t *name)
{
	MSHTML::IHTMLElement2Ptr elem(node);
	MSHTML::IHTMLElementCollectionPtr elements(elem->getElementsByTagName(name));
    
	long len = elements->length;
	for (long i = 0; i < len; ++i)
	{
		MSHTML::IHTMLDOMNodePtr ce(elements->item(i));
		if (!ce)
			break;
		for (int ll = 0; ce->parentNode != node && ll < 30; ++ll)
			MoveUp(true, ce);
		MoveUp(false, ce);
	}
}

///<summary>Place child window in center of parent</summary>
///<params name="parent">Parent window</params>
///<params name="child">Child window</params>
static void CenterChildWindow(CWindow parent, CWindow child)
{
	RECT rcParent, rcChild;
	parent.GetWindowRect(&rcParent);
	child.GetWindowRect(&rcChild);
	int parentW = rcParent.right - rcParent.left;
	
	int parentH = rcParent.bottom - rcParent.top;
	int childW = rcChild.right - rcChild.left;
	int childH = rcChild.bottom - rcChild.top;
	child.MoveWindow(rcParent.left + parentW/2 - childW/2, rcParent.top + parentH/2 - childH/2, childW, childH);
}

// split paragraphs containing BR elements inside text content
static void SplitBRs(MSHTML::IHTMLElement2Ptr elem)
{
	CString text = MSHTML::IHTMLElementPtr(elem)->outerHTML;
	bool changed = false;
	changed = text.Replace(L"<br>", L"</p><p>") > 0;
	changed = text.Replace(L"<br/>", L"</p><p>") > 0;

	if (changed)
		MSHTML::IHTMLElementPtr(elem)->outerHTML = bstr_t(text);
}

// this sub should locate any nested paragraphs and bubble them up
static void RelocateParagraphs(MSHTML::IHTMLDOMNode *node)
{
	MSHTML::IHTMLDOMNodePtr cur(node->firstChild);
	while (cur)
	{
		if (cur->nodeType == 1) //ELEMENT_NODE
		{
			if (!U::scmp(cur->nodeName, L"P"))
			{
				BubbleUp(cur, L"P");
				BubbleUp(cur, L"DIV");
			}
			else
				RelocateParagraphs(cur);
		}
		cur = cur->nextSibling;
	}
}

///<summary>Move text content of DIV into P elements, so DIV contains P&DIV only </summary>
///<params name="elem">HTML element</params>
///<params name="doc">HTML document</params>
static void PackText(MSHTML::IHTMLElement2Ptr elem)
{
	if (!elem)
		return;
	MSHTML::IHTMLDocument2Ptr doc = MSHTML::IHTMLElementPtr(elem)->document;
	MSHTML::IHTMLElementCollectionPtr elements(elem->getElementsByTagName(L"DIV"));
	for (long i = 0; i < elements->length; ++i)
	{
		MSHTML::IHTMLDOMNodePtr div(elements->item(i));
		if (U::scmp(MSHTML::IHTMLElementPtr(div)->className, L"image") == 0)
			continue;
		
        MSHTML::IHTMLDOMNodePtr cur(div->firstChild);
		while (cur)
		{
			//_bstr_t cur_name(cur->nodeName);
			CString cur_name = cur->nodeName;
			if (cur_name != L"P" && cur_name != L"DIV")
			{
				// create a paragraph from a run of !P && !DIV
				MSHTML::IHTMLElementPtr newp(doc->createElement(L"P"));
				MSHTML::IHTMLDOMNodePtr newn(newp);
				cur->replaceNode(newn);
				newn->appendChild(cur);

				while (newn->nextSibling)
				{
					cur_name.SetString(newn->nextSibling->nodeName);
					if (cur_name == L"P" || cur_name == L"DIV")
						break;
					newn->appendChild(newn->nextSibling);
				}
				cur = newn->nextSibling;
			}
			else
				cur = cur->nextSibling;
		}
	}
}
///<summary>Change href attribute for <a> from "file://" to internal link</summary>
///<params name="dom">HTML node</params>
static void FixupLinks(MSHTML::IHTMLDOMNode *dom)
{
	MSHTML::IHTMLElement2Ptr elem(dom);

	if (!elem)
		return;

	MSHTML::IHTMLElementCollectionPtr coll(elem->getElementsByTagName(L"a"));
	if (!coll)
		return;

	/*if (coll->length == 0)
		coll = elem->getElementsByTagName(L"A");*/

	for (long l = 0; l < coll->length; ++l)
	{
		MSHTML::IHTMLElementPtr a(coll->item(l));
		/*if (!a)
			continue;
        */
        CString sHref(a->getAttribute(L"href", 0));
        if (sHref.GetLength()>11 && sHref.Left(7) == L"file://")
        {
            int pos = sHref.Find(L"#");
            if(pos != -1)
            {
                a->setAttribute(L"href", _variant_t(sHref.Right(sHref.GetLength()-pos)), 0);
                
            }
        }
        /*
		_variant_t href(a->getAttribute(L"href", 2)); // return BSTR
		if (V_VT(&href) == VT_BSTR && V_BSTR(&href) &&
			::SysStringLen(V_BSTR(&href)) > 11 &&
			memcmp(V_BSTR(&href), L"file://", 6 * sizeof(wchar_t)) == 0)
		{
			wchar_t *pos = wcschr((wchar_t *)V_BSTR(&href), L'#');
			if (!pos)
				continue;
			a->setAttribute(L"href", pos, 0);
		}*/
	}
}

_ATL_FUNC_INFO CFBEView::DocumentCompleteInfo =
	{CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, (VT_BYREF | VT_VARIANT)}};
_ATL_FUNC_INFO CFBEView::BeforeNavigateInfo =
	{CC_STDCALL, VT_EMPTY, 7, {
								  VT_DISPATCH,
								  (VT_BYREF | VT_VARIANT),
								  (VT_BYREF | VT_VARIANT),
								  (VT_BYREF | VT_VARIANT),
								  (VT_BYREF | VT_VARIANT),
								  (VT_BYREF | VT_VARIANT),
								  (VT_BYREF | VT_BOOL),
							  }};
_ATL_FUNC_INFO CFBEView::VoidInfo =
	{CC_STDCALL, VT_EMPTY, 0};
_ATL_FUNC_INFO CFBEView::EventInfo =
	{CC_STDCALL, VT_BOOL, 1, {VT_DISPATCH}};
_ATL_FUNC_INFO CFBEView::VoidEventInfo =
	{CC_STDCALL, VT_EMPTY, 1, {VT_DISPATCH}};

CFBEView::~CFBEView()
{
	if (m_hdoc)
	{
		DocumentEvents::DispEventUnadvise(Document(), &DIID_HTMLDocumentEvents2);
		TextEvents::DispEventUnadvise(Document()->body, &DIID_HTMLTextContainerEvents2);
		//m_mkc->UnRegisterForDirtyRange(m_dirtyRangeCookie);
		TraceChanges(false);
	}
	if (m_browser)
		BrowserEvents::DispEventUnadvise(m_browser, &DIID_DWebBrowserEvents2);

	if (m_find_dlg)
	{
		CloseFindDialog(m_find_dlg);
		delete m_find_dlg;
	}
}

LRESULT CFBEView::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	if (DefWindowProc(uMsg, wParam, lParam))
		return 1;
	if (!SUCCEEDED(QueryControl(&m_browser)))
		return 1;

	// register browser events handler
	BrowserEvents::DispEventAdvise(m_browser, &DIID_DWebBrowserEvents2);

	return 0;
}

LRESULT CFBEView::OnFocus(UINT, WPARAM, LPARAM, BOOL &)
{
	// pass to document
	if (m_hdoc)
		MSHTML::IHTMLDocument4Ptr(Document())->focus();
	return 0;
}

BOOL CFBEView::PreTranslateMessage(MSG *pMsg)
{
	return SendMessage(WM_FORWARDMSG, 0, (LPARAM)pMsg) != 0;
}

// --------------- Browser editing commands -------------------
void CFBEView::QueryStatus(OLECMD *cmd, int ncmd)
{
	IOleCommandTargetPtr ct(m_browser);
	if (ct)
		ct->QueryStatus(&CGID_MSHTML, ncmd, cmd, NULL);
}

//Return OLE command status
CString CFBEView::QueryCmdText(ULONG cmd)
{
	IOleCommandTargetPtr ct(m_browser);
	if (ct)
	{
		OLECMD oc = { cmd };
		struct
		{
			OLECMDTEXT oct;
			wchar_t buffer[512];
		} oct = { {OLECMDTEXTF_NAME, 0, 512} };
		if (SUCCEEDED(ct->QueryStatus(&CGID_MSHTML, 1, &oc, &oct.oct)))
			return oct.oct.rgwz;
	}
	return CString();
}

//Exec MSHTML Command
LRESULT CFBEView::ExecCommand(int cmd)
{
	IOleCommandTargetPtr ct(m_browser);
	if (ct)
		ct->Exec(&CGID_MSHTML, cmd, 0, NULL, NULL);
	return 0;
}

LRESULT CFBEView::OnUndo(WORD, WORD, HWND hWnd, BOOL &)
{
	LRESULT res = ExecCommand(IDM_UNDO);
	// update tree view
	::SendMessage(m_frame, WM_COMMAND, MAKEWPARAM(0, IDN_TREE_RESTORE), 0);
	return res;
}

LRESULT CFBEView::OnRedo(WORD, WORD, HWND, BOOL &)
{
	LRESULT res = ExecCommand(IDM_REDO);
	// update tree view
	::SendMessage(m_frame, WM_COMMAND, MAKEWPARAM(0, IDN_TREE_RESTORE), 0);
	return res;
}

LRESULT CFBEView::OnCut(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_CUT);
}

LRESULT CFBEView::OnCopy(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_COPY);
}

LRESULT CFBEView::OnBold(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_BOLD);
}

LRESULT CFBEView::OnItalic(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_ITALIC);
}

LRESULT CFBEView::OnStrik(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_STRIKETHROUGH);
}

LRESULT CFBEView::OnSup(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_SUPERSCRIPT);
}

LRESULT CFBEView::OnSub(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_SUBSCRIPT);
}

//Create Link from selection
LRESULT CFBEView::OnStyleLink(WORD, WORD, HWND, BOOL &)
{
	try
	{
		m_mk_srv->BeginUndoUnit(L"Create Link");
		if (Document()->execCommand(L"CreateLink", VARIANT_FALSE, _variant_t(L"")) == VARIANT_TRUE)
		{
			::SendMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_SEL_CHANGE), (LPARAM)m_hWnd);
		//	::SendMessage(m_frame, WM_COMMAND, MAKELONG(IDC_HREF, IDN_WANTFOCUS), (LPARAM)m_hWnd);
		}
		m_mk_srv->EndUndoUnit();
	}
	catch (_com_error &)
	{
	}
	return 0;
}

//Create Footnote from selection
LRESULT CFBEView::OnStyleFootnote(WORD, WORD, HWND, BOOL &)
{
	try
	{
		m_mk_srv->BeginUndoUnit(L"Create Footnote");
		if (Document()->execCommand(L"CreateLink", VARIANT_FALSE, _variant_t(L"")) == VARIANT_TRUE)
		{
			MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
			MSHTML::IHTMLDOMRangePtr range = sel->getRangeAt(0);
			MSHTML::IHTMLElementPtr pe = range->commonAncestorContainer->parentNode;
			if (U::scmp(pe->tagName, L"A") == 0)
				pe->className = L"note";
		}
		m_mk_srv->EndUndoUnit();
		::SendMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_SEL_CHANGE), (LPARAM)m_hWnd);
		//::SendMessage(m_frame, WM_COMMAND, MAKELONG(IDC_HREF, IDN_WANTFOCUS), (LPARAM)m_hWnd);
	}
	catch (_com_error &)
	{
	}
	return 0;
}

//Create Footnote from selection
LRESULT CFBEView::OnStyleCode(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_UNDERLINE);
	//return 0;
}

bool CFBEView::CheckCommand(WORD wID)
{
	switch (wID)
	{
	case ID_ADD_BODY:
		return true;
	case ID_ADD_BODY_NOTES:
		return canApplyCommand(L"body_notes");
	case ID_ADD_ANN:
		return canApplyCommand(L"annotation");
	case ID_ADD_SECTION:
		return canApplyCommand(L"section");
	case ID_ADD_TITLE:
		return canApplyCommand(L"title");
	case ID_STYLE_NORMAL:
		return Document()->queryCommandEnabled(L"removeFormat") == VARIANT_TRUE;
	case ID_STYLE_BOLD:
		return Document()->queryCommandEnabled(L"bold") == VARIANT_TRUE;
	case ID_STYLE_ITALIC:
		return Document()->queryCommandEnabled(L"italic") == VARIANT_TRUE;
	case ID_STYLE_SUPERSCRIPT:
		return Document()->queryCommandEnabled(L"superscript") == VARIANT_TRUE;
	case ID_STYLE_SUBSCRIPT:
		return Document()->queryCommandEnabled(L"subscript") == VARIANT_TRUE;
	case ID_STYLE_STRIKETROUGH: 
		return Document()->queryCommandEnabled(L"strikeThrough") == VARIANT_TRUE;
	case ID_STYLE_CODE:
		return canApplyCommand(L"title") && Document()->queryCommandEnabled(L"strikeThrough") == VARIANT_TRUE;
			//Document()->queryCommandEnabled(L"underline") == VARIANT_TRUE;
	case ID_ADD_SUBTITLE:
		return canApplyCommand(L"subtitle");
	case ID_ADD_TEXTAUTHOR:
		return canApplyCommand(L"text-author");
	case ID_ADD_IMAGE:
		return canApplyCommand(L"image");
	case ID_ATTACH_BINARY:
		return true;
	case ID_ADD_EPIGRAPH:
		return canApplyCommand(L"epigraph");
	case ID_EDIT_SPLIT:
		return true; // SplitContainer(true);
	case ID_ADD_POEM:
		return canApplyCommand(L"poem");
	case ID_ADD_STANZA:
		return canApplyCommand(L"stanza");
	case ID_ADD_CITE:
		return canApplyCommand(L"cite");
	case ID_INS_TABLE:
		return canApplyCommand(L"table");
	case ID_GOTO_FOOTNOTE:
	case ID_GOTO_REFERENCE:
		return canApplyCommand(L"link");
	case ID_EDIT_CLONE:
		return false;
//		return bCall(L"CloneContainer", SelectionsStruct(L""));
	case ID_EDIT_MERGE:
		return false;
//		return bCall(L"MergeContainers", SelectionsStruct(L""));
	case ID_EDIT_REMOVE_OUTER:
		return false;
//		return bCall(L"RemoveOuterContainer", SelectionsStruct(L""));
	case ID_STYLE_LINK:
	case ID_STYLE_NOTE:
		try
		{
			return canApplyCommand(L"link") && Document()->queryCommandEnabled(L"CreateLink") == VARIANT_TRUE;
		}
		catch (_com_error)
		{
		}
		break;
	}
	return false;
}

bool CFBEView::CheckSetCommand(WORD wID)
{
	//switch (wID)
	//{
	//case ID_STYLE_CODE:
		//return bCall(L"IsCode");
//		return canApplyCommand(L"code");
	//}

	return false;
}

// changes tracking
MSHTML::IHTMLDOMNodePtr CFBEView::GetChangedNode()
{
	MSHTML::IMarkupPointerPtr p1, p2;
	m_mk_srv->CreateMarkupPointer(&p1);
	m_mk_srv->CreateMarkupPointer(&p2);

	m_mkc->GetAndClearDirtyRange(m_dirtyRangeCookie, p1, p2);

	MSHTML::IHTMLElementPtr e1, e2;
	p1->CurrentScope(&e1);
	p2->CurrentScope(&e2);
	p1.Release();
	p2.Release();

	while (e1 && e1 != e2 && e1->contains(e2) != VARIANT_TRUE)
	{
		e1 = e1->parentElement;
	}
	
	return e1;
}

// Splitting
bool CFBEView::SplitContainer(bool fCheck)
{
	try
	{
		MSHTML::IHTMLTxtRangePtr rng(Document()->selection->createRange());
		if (!rng)
			return false;

		MSHTML::IHTMLElementPtr pe(rng->parentElement());
		while (pe && U::scmp(pe->tagName, L"DIV"))
			pe = pe->parentElement;

		if (!pe || (U::scmp(pe->className, L"section") && U::scmp(pe->className, L"stanza")))
			return false;

		MSHTML::IHTMLTxtRangePtr r2(rng->duplicate());
		r2->moveToElementText(pe);

		if (rng->compareEndPoints(L"StartToStart", r2) == 0)
			return false;

		MSHTML::IHTMLTxtRangePtr r3(rng->duplicate());
		r3->collapse(true);
		MSHTML::IHTMLTxtRangePtr r4(rng->duplicate());
		r4->collapse(false);

		if (!(bool)pe || GetHP(r3->parentElement()) != pe || GetHP(r4->parentElement()) != pe)
			return false;

		if (fCheck)
			return true;

		// At this point we are ready to split

		// Create an undo unit
		CString name(L"split ");
		name += (const wchar_t *)pe->className;
		m_mk_srv->BeginUndoUnit((TCHAR *)(const TCHAR *)name);

		//// Create a new element
		MSHTML::IHTMLElementPtr ne(Document()->createElement(L"DIV"));
		ne->className = pe->className;
		_bstr_t className = pe->className;
		// SeNS: issue #153
		_bstr_t id = pe->id;
		pe->id = L"";

		MSHTML::IHTMLElementPtr peTitle(Document()->createElement(L"DIV"));
		MSHTML::IHTMLElementCollectionPtr peColl = pe->children;
		{
			MSHTML::IHTMLElementPtr peChild = peColl->item(0);
			if (!U::scmp(peChild->tagName, L"DIV") && !U::scmp(peChild->className, L"title"))
				peTitle->innerHTML = peChild->outerHTML;
			else
				peTitle = NULL;
		}

		// Create and position markup pointers
		MSHTML::IMarkupPointerPtr selstart, selend, elembeg, elemend;

		m_mk_srv->CreateMarkupPointer(&selstart);
		m_mk_srv->CreateMarkupPointer(&selend);
		m_mk_srv->CreateMarkupPointer(&elembeg);
		m_mk_srv->CreateMarkupPointer(&elemend);

		MSHTML::IHTMLTxtRangePtr titleRng(rng->duplicate());
		m_mk_srv->MovePointersToRange(titleRng, selstart, selend);
		U::ElTextHTML title(titleRng->htmlText, titleRng->text);

		MSHTML::IHTMLTxtRangePtr preRng(rng->duplicate());
		elembeg->MoveAdjacentToElement(pe, MSHTML::ELEM_ADJ_AfterBegin);
		m_mk_srv->MoveRangeToPointers(elembeg, selstart, preRng);
		U::ElTextHTML pre(preRng->htmlText, preRng->text);

		MSHTML::IHTMLElementCollectionPtr peChilds = pe->children;
		MSHTML::IHTMLElementPtr elLast = peChilds->item(peChilds->length - 1);
		if (U::scmp(elLast->innerText, L"") == 0)
			elLast->innerText = L"123";

		MSHTML::IHTMLTxtRangePtr postRng(rng->duplicate());
		elemend->MoveAdjacentToElement(pe, MSHTML::ELEM_ADJ_BeforeEnd);
		m_mk_srv->MoveRangeToPointers(selend, elemend, postRng);
		U::ElTextHTML post(postRng->htmlText, postRng->text);

		// Check if title needs to be created and further text to be copied
		bool fTitle = !title.text.IsEmpty();
		bool fContent = !post.html.IsEmpty();

		if (fTitle && title.html.Find(L"<P") == -1)
			title.html = CString(L"<P>") + title.html + CString(L"</P>");
		if (fContent && post.html.Find(L"<P") == -1)
			post.html = CString(L"<P>") + post.html + CString(L"</P>");

		title.html.Remove(L'\r');
		title.html.Remove(L'\n');
		post.html.Remove(L'\r');
		post.html.Remove(L'\n');

		if (post.html.Find(L"<P>&nbsp;</P>") == 0 && post.html.GetLength() > 13 && fTitle && title.html.Find(L"<P>&nbsp;</P>") != title.html.GetLength() - 14)
			post.html.Delete(0, 13);
		if (post.html.Find(L"<P>123</P>") != -1)
			post.html.Replace(L"<P>123</P>", L"<P>&nbsp;</P>");

		// Insert it after pe
		MSHTML::IHTMLElement2Ptr(pe)->insertAdjacentElement(L"afterEnd", ne);

		// Move content or create new
		if (fContent)
		{
			// Create and position destination markup pointer
			if (post.html == L"<P>&nbsp;</P>")
				post.html += L"<P>&nbsp;</P>";
			ne->innerHTML = post.html.AllocSysString();
			// SeNS: issue #153
			ne->id = id;
		}
		else
		{
			MSHTML::IHTMLElementPtr para(Document()->createElement(L"P"));
//			MSHTML::IHTMLElement3Ptr(para)->inflateBlock = VARIANT_TRUE;
			MSHTML::IHTMLElement2Ptr(ne)->insertAdjacentElement(L"beforeEnd", para);
		}

		// Create and move title if needed
		if (fTitle)
		{
			MSHTML::IHTMLElementPtr elTitle(Document()->createElement(L"DIV"));
			elTitle->className = L"title";
			MSHTML::IHTMLElement2Ptr(ne)->insertAdjacentElement(L"afterBegin", elTitle);

			// Create and position destination markup pointer
			elTitle->innerHTML = title.html.AllocSysString();

			// Delete all containers from title
			KillDivs(elTitle);
			KillStyles(elTitle);
		}

		if (pre.html.Find(L"<P") == -1)
		{
			if (pre.html == L"")
				pre.html = L"<P>&nbsp;</P>";
			else
				pre.html = CString(L"<P>") + pre.html + CString(L"</P>");
		}
		pe->innerHTML = pre.html.AllocSysString();

		// Ensure we have good html
		FixupParagraphs(ne);
		PackText(ne);

		peColl = pe->children;
		if (peColl->length == 1)
		{
			MSHTML::IHTMLElementPtr peChild = peColl->item(0);
			if (!U::scmp(peChild->tagName, L"DIV") && !U::scmp(peChild->className, className.GetBSTR()))
				m_mk_srv->RemoveElement(peChild);
		}

		MSHTML::IHTMLElementCollectionPtr neColl = ne->children;
		if (neColl->length == 1)
		{
			MSHTML::IHTMLElementPtr neChild = neColl->item(0);
			if (!U::scmp(neChild->tagName, L"DIV") && !U::scmp(neChild->className, className.GetBSTR()))
				m_mk_srv->RemoveElement(neChild);
		}

		CString peTitSect;
		if (peTitle)
		{
			peTitSect = peTitle->innerHTML.GetBSTR();
			peTitSect += L"<P>&nbsp;</P>";
		}

		CString b = pe->innerText;
		b.Remove(L'\r');
		b.Remove(L'\n');
		CString c = peTitle ? peTitle->innerText : L"";
		c.Remove(L'\r');
		c.Remove(L'\n');

		if (peTitle && !U::scmp(b, c))
			pe->innerHTML = peTitSect.AllocSysString();

		// Close undo unit
		m_mk_srv->EndUndoUnit();

		// Move cursor to newly created item
		GoTo(ne, false);
	}
	catch (_com_error &e)
	{
		U::ReportError(e);
	}

	return false;
}

// searching
bool CFBEView::CanFindNext()
{
	return !m_fo.pattern.IsEmpty();
}

void CFBEView::StopIncSearch()
{
	if (m_is_start)
		m_is_start.Release();
}

bool CFBEView::DoIncSearch(const CString &str, bool fMore)
{
	++m_ignore_changes;
	m_fo.pattern = str;
	bool ret = DoSearch(fMore);
	--m_ignore_changes;
	return ret;
}

void CFBEView::StartIncSearch()
{
	try
	{
		m_is_start = Document()->selection->createRange();
		m_is_start->collapse(VARIANT_TRUE);
	}
	catch (_com_error &)
	{
	}
}

void CFBEView::CancelIncSearch()
{
	if (m_is_start)
	{
		m_is_start->raw_select();
		m_is_start.Release();
	}
}

CString CFBEView::LastSearchPattern()
{
	return m_fo.pattern;
}

int CFBEView::ReplaceAllRe(const CString &re, const CString &str, MSHTML::IHTMLElementPtr elem, CString cntTag)
{
	m_fo.pattern = re;
	m_fo.replacement = str;
	m_fo.fRegexp = true;
	m_fo.flags = 0;
	return GlobalReplace(elem, cntTag);
}

int CFBEView::ReplaceToolWordsRe(const CString &re,
								 const CString &str,
								 MSHTML::IHTMLElementPtr fbw_body,
								 bool replace,
								 CString cntTag,
								 int *pIndex,
								 int *globIndex,
								 int replNum)
{
	m_fo.pattern = re;
	m_fo.replacement = str;
	m_fo.fRegexp = true;
	m_fo.flags = FRF_CASE | FRF_WHOLE;
	m_fo.replNum = replNum;
	return ToolWordsGlobalReplace(fbw_body, pIndex, globIndex, !replace, cntTag);
}

// script calls
//TO_DO not used??
void CFBEView::ImgSetURL(IDispatch *elem, const CString &url)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t ve(elem);
		_variant_t vu((const TCHAR *)url);
		dd.Invoke2(L"ImgSetURL", &ve, &vu);
	}
	catch (_com_error &)
	{
	}
}

IDispatchPtr CFBEView::Call(const wchar_t *name)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t ret;
		_variant_t vt2((false));
		dd.Invoke1(name, &vt2, &ret);
		if (V_VT(&ret) == VT_DISPATCH)
			return V_DISPATCH(&ret);
	}
	catch (_com_error &)
	{
	}
	return IDispatchPtr();
}
IDispatchPtr CFBEView::Call(const wchar_t *name, IDispatch *pDisp)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt;
		if (pDisp)
			vt = pDisp;
		_variant_t vt2(false);
		_variant_t ret;
		dd.Invoke2(name, &vt, &vt2, &ret);
		if (V_VT(&ret) == VT_DISPATCH)
			return V_DISPATCH(&ret);
	}
	catch (_com_error &)
	{
	}
	return IDispatchPtr();
}

bool CFBEView::bCall(const wchar_t *name, int nParams, VARIANT *params)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t ret;
		dd.InvokeN(name, params, nParams, &ret);
		return vt2bool(ret);
	}
	catch (_com_error &err)
	{
		U::ReportError(err);
	}

	return false;
}

bool CFBEView::bCall(const wchar_t *name, IDispatch *pDisp)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt;
		if (pDisp)
			vt = pDisp;
		_variant_t vt2(true);
		_variant_t ret;
		dd.Invoke2(name, &vt, &vt2, &ret);
		return vt2bool(ret);
	}
	catch (_com_error &)
	{
	}
	return false;
}

bool CFBEView::bCall(const wchar_t *name)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt2(true);
		_variant_t ret;
		dd.Invoke1(name, &vt2, &ret);
		return vt2bool(ret);
	}
	catch (_com_error &)
	{
	}
	return false;
}

// Check if can run command
bool CFBEView::canApplyCommand(const CString name)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt1(name);
		_variant_t ret;
		dd.Invoke1(L"CanApplyStyle", &vt1, &ret);
		return vt2bool(ret);
	}
	catch (_com_error &)
	{
		return false;
	}
}

// utilities
//Get HTML-path for HTML-element
static CString GetPath(MSHTML::IHTMLElementPtr elem)
{
    CString path(L"");

    if (!(bool)elem)
        return path;
    // climb up from source element till body
    while (elem)
    {
        CString cur = elem->tagName;
        CString cid = elem->id;
        if (cur == L"BODY" || cid == "fbw_body")
            return path;

        CString cls = elem->className;
        if (cls.GetLength() > 0)
            cur = cls;
        if (cid.GetLength() > 0)
            cur += L"#" + cid;
        
        if (!path.IsEmpty())
            path = L"/" + path;
        path = cur + path;
        elem = elem->parentElement;
    }
    return path;
}

CString CFBEView::SelPath()
{
	return GetPath(SelectionContainer());
}

MSHTML::IHTMLElementPtr CFBEView::SelectionContainer()
{
	//if (m_cur_sel)
	//	return m_cur_sel;
	return SelectionContainerImp();
}

///Return HTML-Element which contains selection
MSHTML::IHTMLElementPtr CFBEView::SelectionContainerImp()
{
	MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
	if (!sel)
		return nullptr;
	//return MSHTML::IHTMLElementPtr();
	if (sel->rangeCount == 0)
		return nullptr;
	//return MSHTML::IHTMLElementPtr();
	MSHTML::IHTMLDOMRangePtr range = sel->getRangeAt(0);
	/*else
	{
	MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(Document())->createRange();
	sel->addRange(range);
	}*/

	if (range)
	{
		if (range->commonAncestorContainer->nodeType != NODE_ELEMENT)
			return range->commonAncestorContainer->parentNode;
		else
			return range->commonAncestorContainer;
	}
	return nullptr;
	//return MSHTML::IHTMLElementPtr();
	/* textrange variant
	try
	{
	IDispatchPtr selrange(Document()->selection->createRange());
	MSHTML::IHTMLTxtRangePtr range(selrange);
	if (range)
	{
	return range->parentElement();
	}
	MSHTML::IHTMLControlRangePtr coll(selrange);
	if (coll)
	return coll->commonParentElement();
	}
	catch (_com_error &err)
	{
	U::ReportError(err);
	}

	return MSHTML::IHTMLElementPtr();*/
}
/// <summary>Scroll window to center HTML-element</summary>
/// <param name="e">HTML-element</param>
void CFBEView::ScrollToElement(MSHTML::IHTMLElementPtr e)
{
    if(!e)
        return;
    MSHTML::IHTMLRectPtr rect = MSHTML::IHTMLElement2Ptr(e)->getBoundingClientRect();
    MSHTML::IHTMLWindow2Ptr window(MSHTML::IHTMLDocument2Ptr(Document())->parentWindow);
    if (rect && window)
    {
		if (rect->bottom - rect->top <= MSHTML::IHTMLWindow7Ptr(window)->innerHeight)
			window->scrollBy(0, (rect->top + rect->bottom - MSHTML::IHTMLWindow7Ptr(window)->innerHeight) / 2);
		else
			window->scrollBy(0, rect->top);
    }
}

void CFBEView::ScrollToSelection()
{
    // scroll window to selection
    MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
    MSHTML::IHTMLDOMNodePtr el = nullptr;
    if (sel && sel->rangeCount > 0)
    {
        MSHTML::IHTMLDOMRangePtr range = sel->getRangeAt(0)->cloneRange();
        range->collapse(VARIANT_FALSE);
        el = range->endContainer;
		if (el->nodeType == NODE_ELEMENT)
			// range offset = number of child element
			if(range->endOffset>0)
				el = MSHTML::IHTMLDOMChildrenCollectionPtr(el->childNodes)->item(range->endOffset-1);
			else
				el = MSHTML::IHTMLDOMChildrenCollectionPtr(el->childNodes)->item(0);
		else
        {
            // search node with ELEMENT type
            MSHTML::IHTMLDOMNodePtr o = el;
            while (o && o->nodeType != NODE_ELEMENT)
                o = o->previousSibling;
            el = o ? o : el->parentNode;
        }
        range->Detach();
    }
	ScrollToElement(el);
}

void CFBEView::GoTo(MSHTML::IHTMLElementPtr e, bool fScroll)
{
	if (!e)
		return;
	if (fScroll)
		ScrollToElement(e);

	MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(Document())->createRange();
	if (!range)
		return;

	range->selectNodeContents(e);
	range->collapse(VARIANT_TRUE);

	MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
	if (!sel)
		return;
	sel->removeAllRanges();
	sel->addRange(range);
}

MSHTML::IHTMLElementPtr CFBEView::SelectionAnchor(MSHTML::IHTMLElementPtr cur)
{
	try
	{
		while (cur)
		{
			CString tn = cur->tagName;
			CString cn = cur->className;
			if (tn == L"A" || ((tn == L"DIV" || tn == L"SPAN") && cn == L"image"))
				return cur;
			cur = cur->parentElement;
		}
	}
	catch (_com_error &)
	{
	}
	return nullptr;
}

// New universal functions
/// <summary>Get closest P or DIV parent element for current selection</summary>
/// <summary>Get closest SPAN parent element for current selection</summary>
/// <summary>Get closest parent DIV with image element for current selection</summary>
/// <summary>Get closest parent table element for current selection</summary>
/// <returns>DIV with image element or empty element</returns>
MSHTML::IHTMLElementPtr CFBEView::SelectionsStruct(const CString search)
{
	try
	{
		MSHTML::IHTMLElementPtr cur(SelectionContainer());
		while (cur)
		{
            CString tn = cur->tagName;
            CString cn = cur->className;
            
			if (search == L"")   //anchor
			{
				if (tn == L"P" || tn == L"DIV")
					return cur;
			}
			if(search==L"anchor")   //anchor
            {
                if (tn == L"A" || ((tn == L"DIV" || tn == L"SPAN") && cn == L"image"))
                    return cur;
            }
            if(search==L"struct")   //struct
            {
                if (tn == L"P" || tn == L"DIV")
                    return cur;
            }
            if(search==L"th")   //TableCon
            {
                if (cn == L"th" || cn == L"td")
                    return cur;
            }
            if(search==L"image")   //image
            {
                if (cn == L"image" && tn != L"SPAN")
                    return cur;
            }
			if (search == L"table")   //table
			{
				if (cn == L"table")
					return cur;
			}
			if(search==L"section")   //section
            {
                if (cn == L"section")
                    return cur;
            }
            if(search==L"code")   //code
            {
           		if (tn == L"SPAN" && cn == L"code")
                    return cur;
            }
			cur = cur->parentElement;
		}
	}
	catch (_com_error &)
	{
	}
	return MSHTML::IHTMLElementPtr();
}

// New universal function
MSHTML::IHTMLElementPtr CFBEView::SelectionsStructB(const CString search, _bstr_t &style)
{
	try
	{
		MSHTML::IHTMLElementPtr cur(SelectionContainer());
		while (cur)
		{
            CString cn = cur->className;
            
            if(search==L"table")
            {
                if (cn == L"table")
                {
                    style = U::GetAttrB(cur, L"fbstyle");
                    return cur;
                }
            }
            if(search==L"th")
            {
            
       			if (cn == L"th" || cn == L"td")
                {
                    style = U::GetAttrB(cur, L"fbstyle");
                    return cur;
                }
            }
            if(search==L"colspan")
            {
                if (cn == L"th" || cn == L"td")
                {
                    style = U::GetAttrB(cur, L"fbcolspan");
                    return cur;
                }
            }
            if(search == L"rowspan")
            {
                if (cn == L"th" || cn == L"td")
                {
                    style = U::GetAttrB(cur, L"fbrowspan");
                    return cur;
                }
            }
            if(search==L"tralign")
            {
                if (cn == L"tr")
                {
                    style = U::GetAttrB(cur, L"fbalign");
                    return cur;
                }
            }
            if(search==L"align")
            {
                if (cn == L"tr" || cn == L"th" || cn == L"td")
                {
                    style = U::GetAttrB(cur, L"fbalign");
                    return cur;
                }
            }
            if(search==L"valign")
            {
                if (cn == L"th" || cn == L"td")
                {
                    style = U::GetAttrB(cur, L"fbvalign");
                    return cur;
                }
            }
			cur = cur->parentElement;
		}
	}
	catch (_com_error &)
	{
	}
	return MSHTML::IHTMLElementPtr();
}

void CFBEView::Normalize()
{
	try
	{
		MSHTML::IHTMLDOMNodePtr el = MSHTML::IHTMLDocument3Ptr(m_hdoc)->getElementById(L"fbw_body");
		if (!el)
			return;

		// wrap in an undo unit
		m_mk_srv->BeginUndoUnit(L"Normalize");
		CString cc1 = m_hdoc->body->outerHTML;
		// remove unsupported elements
		RemoveUnk(el, m_hdoc);
		CString cc2 = m_hdoc->body->outerHTML;

		// merge similar neighbor elements
		MergeEqualHTMLElements(el, m_hdoc);
		CString cc3 = m_hdoc->body->outerHTML;

		// bubble up of nested DIVs and Ps
		RelocateParagraphs(el);
		CString cc4 = m_hdoc->body->outerHTML;

		// delete empty nodes
		RemoveEmptyNodes(el);
		CString cc5 = m_hdoc->body->outerHTML;

        // make sure text appears under Ps only
		PackText(el);
		CString cc6 = m_hdoc->body->outerHTML;

        // get rid of nested Ps once more
		RelocateParagraphs(el);
		CString cc7 = m_hdoc->body->outerHTML;

        // convert BRs to separate paragraphs
		SplitBRs(el);
		CString cc8 = m_hdoc->body->outerHTML;

        // delete empty nodes again
		RemoveEmptyNodes(el);
		CString cc9 = m_hdoc->body->outerHTML;

        // fixup links in <a>
		FixupLinks(el);
        
		m_mk_srv->EndUndoUnit();
	}
	catch (_com_error &e)
	{
		U::ReportError(e);
	}
}
///<summary>Set inflateBlock for <P></summary>
///<params name="elem">HTML element</params>
static void FixupParagraphs(MSHTML::IHTMLElement2Ptr elem)
{
	if (!elem)
		return;
	MSHTML::IHTMLElementCollectionPtr pp = elem->getElementsByTagName(L"P");
	for (long l = 0; l < pp->length; ++l)
		MSHTML::IHTMLElement3Ptr(pp->item(l))->inflateBlock = VARIANT_TRUE;
}

LRESULT CFBEView::OnPaste(WORD, WORD, HWND, BOOL &)
{
	try
	{
		m_mk_srv->BeginUndoUnit(L"Paste");
		++m_enable_paste;

		// added by SeNS: process clipboard and change nbsp
		if (OpenClipboard())
		{
			// process text
			if (IsClipboardFormatAvailable(CF_TEXT) || IsClipboardFormatAvailable(CF_UNICODETEXT))
			{
				if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
				{
					HANDLE hData = GetClipboardData(CF_UNICODETEXT);
					TCHAR *buffer = (TCHAR *)GlobalLock(hData);
					CString fromClipboard(buffer);
					GlobalUnlock(hData);

					fromClipboard.Replace(L"\u00A0", _Settings.GetNBSPChar());
					HGLOBAL clipbuffer = GlobalAlloc(GMEM_DDESHARE, (fromClipboard.GetLength() + 1) * sizeof(TCHAR));
					buffer = (TCHAR *)GlobalLock(clipbuffer);
					wcscpy(buffer, fromClipboard);
					GlobalUnlock(clipbuffer);
					SetClipboardData(CF_UNICODETEXT, clipbuffer);
				}
			}
			// process bitmaps from clipboard
			else if (IsClipboardFormatAvailable(CF_BITMAP))
			{
				HBITMAP hBitmap = (HBITMAP)GetClipboardData(CF_BITMAP);
				TCHAR szPathName[MAX_PATH] = {0};
				TCHAR szFileName[MAX_PATH] = {0};
				if (::GetTempPath(sizeof(szPathName) / sizeof(TCHAR), szPathName))
					if (::GetTempFileName(szPathName, L"img", ::GetTickCount(), szFileName))
					{
						int quality = _Settings.m_jpeg_quality;

						CString fileName(szFileName);
						FBE::CImageEx image;
						image.Attach(hBitmap);

						if (_Settings.GetImageType() == 0)
						{
							fileName.Replace(L".tmp", L".png");
							image.Save(fileName, Gdiplus::ImageFormatPNG);
						}
						else
						{
							fileName.Replace(L".tmp", L".jpg");
							// set encoder quality
							Gdiplus::EncoderParameters encoderParameters[1];
							encoderParameters[0].Count = 1;
							encoderParameters[0].Parameter[0].Guid = Gdiplus::EncoderQuality;
							encoderParameters[0].Parameter[0].NumberOfValues = 1;
							encoderParameters[0].Parameter[0].Type = Gdiplus::EncoderParameterValueTypeLong;
							encoderParameters[0].Parameter[0].Value = &quality;
							image.Save(fileName, Gdiplus::ImageFormatJPEG, &encoderParameters[0]);
						}

						AddImage(fileName);
						::DeleteFile(fileName);
					}
			}
			CloseClipboard();
		}

		IOleCommandTargetPtr(m_browser)->Exec(&CGID_MSHTML, IDM_PASTE, 0, NULL, NULL);
		--m_enable_paste;
		// if (m_normalize)
			Normalize();
		m_mk_srv->EndUndoUnit();
	}
	catch (_com_error &err)
	{
		U::ReportError(err);
	}

	return 0;
}

// searching
bool CFBEView::DoSearch(bool fMore)
{
#ifndef USE_PCRE
	if (m_fo.match)
		m_fo.match.Release();
#endif
	if (m_fo.pattern.IsEmpty())
	{
		if (m_is_start)
			m_is_start->raw_select();
		return true;
	}

	// added by SeNS
	if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
		m_fo.pattern.Replace(L"\u00A0", _Settings.GetNBSPChar());

	return m_fo.fRegexp ? DoSearchRegexp(fMore) : DoSearchStd(fMore);
}

// Removes HTML tags
void RemoveTags(CString &src)
{
	int openTag = 0, closeTag = 0;
	while (openTag != -1)
	{
		openTag = src.Find(L"<", 0);
		closeTag = src.Find(L">", openTag + 1);
		if (openTag != -1 && closeTag > openTag)
			src.Delete(openTag, closeTag - openTag + 1);
	}
}

// Returns text offset including images (treated as a 3 chars each)
int CFBEView::TextOffset(MSHTML::IHTMLTxtRange *rng, U::ReMatch rm, CString txt, CString htmlTxt)
{
	CString text(txt);
	CString match = rm->Value;
	// special fix for "Words" dialog
	match = match.TrimRight(10);
	match = match.TrimRight(13);
	int num = 0, pos = 1;
	if (text.IsEmpty())
		text.SetString(rng->text);
	while (num < text.GetLength())
	{
		num = text.Find(match, num);
		if ((num == rm->FirstIndex) || (num == -1))
			break;
		num += 1;
		pos++;
	}
	CString html(htmlTxt);
	if (html.IsEmpty())
		html.SetString(rng->htmlText);

	// change <IMG to "afro-american" ☻<IMG LOL
	html.Replace(L"<IMG", L"☻<IMG");
	RemoveTags(html);

	num = 0;
	for (int i = 0; i < pos; i++)
	{
		num = html.Find(match, num);
		num += 1;
	}
	// find number of images occurences
	html = html.Left(num);
	pos = num = 0;
	while (num < html.GetLength())
	{
		num = html.Find(L"☻", num);
		if (num == -1)
			break;
		num += 1;
		pos++;
	}
	return pos * 3;
}

void CFBEView::SelMatch(MSHTML::IHTMLTxtRange *tr, U::ReMatch rm)
{
	// SeNS: fix for issue #147
	int numImages = TextOffset(tr, rm);

	tr->collapse(VARIANT_TRUE);
	tr->move(L"character", rm->FirstIndex + numImages);
	if (tr->moveStart(L"character", 1) == 1)
		tr->move(L"character", -1);
	tr->moveEnd(L"character", rm->Length);
	// set focus to editor if selection empty
	if (!rm->Length)
		SetFocus();
	tr->select();
	m_fo.match = rm;
}

bool CFBEView::DoSearchRegexp(bool fMore)
{
	try
	{
		// well, try to compile it first
		U::RegExp re;
#ifdef USE_PCRE
		re = new U::IRegExp2();
#else
		CheckError(re.CreateInstance(L"VBScript.RegExp"));
#endif
		re->IgnoreCase = m_fo.flags & 4 ? VARIANT_FALSE : VARIANT_TRUE;
		re->Global = VARIANT_TRUE;
		re->Pattern = (const wchar_t *)m_fo.pattern;

		// locate starting paragraph
		MSHTML::IHTMLTxtRangePtr sel(Document()->selection->createRange());
		if (!fMore && (bool)m_is_start)
			sel = m_is_start->duplicate();
		if (!(bool)sel)
			return false;

		MSHTML::IHTMLElementPtr sc(SelectionsStruct(L""));
		long s_idx = 0;
		long s_off1 = 0;
		long s_off2 = 0;
		if ((bool)sc)
		{
			s_idx = sc->sourceIndex;
			if (U::scmp(sc->tagName, L"P") == 0 && (bool)sel)
			{
				s_off2 = sel->text.length();
				MSHTML::IHTMLTxtRangePtr pr(sel->duplicate());
				pr->moveToElementText(sc);
				pr->setEndPoint(L"EndToStart", sel);
				s_off1 = pr->text.length();
				s_off2 += s_off1;
			}
		}

		// walk the all collection now, looking for the next P
		MSHTML::IHTMLElementCollectionPtr all(Document()->all);
		long all_len = all->length;
		long incr = m_fo.flags & 1 ? -1 : 1;
		bool fWrapped = false;

		// * search in starting element
		if ((bool)sc && U::scmp(sc->tagName, L"P") == 0)
		{
			sel->moveToElementText(sc);
			CString selText = sel->text;

			U::ReMatches rm(re->Execute(sel->text));
			// changed by SeNS: fix for issue #62
			if (rm->Count > 0 && !(selText.IsEmpty() && fMore))
			{
				if (incr > 0)
				{
					for (long l = m_startMatch; l < rm->Count; ++l)
					{
						U::ReMatch crm(rm->Item[l]);
						if (crm->FirstIndex >= s_off2)
						{
							m_startMatch = l + 1;
							SelMatch(sel, crm);
							return true;
						}
					}
				}
				else
				{
					for (long l = m_endMatch; l >= 0; --l)
					{
						U::ReMatch crm(rm->Item[l]);
						if (crm->FirstIndex < s_off1)
						{
							SelMatch(sel, crm);
							m_endMatch = l - 1;
							return true;
						}
					}
				}
			}
		}

		// search all others
		for (long cur = s_idx + incr;; cur += incr)
		{
			// adjust out of bounds indices
			if (cur < 0)
			{
				cur = all_len - 1;
				fWrapped = true;
			}
			else if (cur >= all_len)
			{
				cur = 0;
				fWrapped = true;
			}

			// check for wraparound
			if (cur == s_idx)
				break;

			// check current element type
			MSHTML::IHTMLElementPtr elem(all->item(cur));
			if (!(bool)elem || U::scmp(elem->tagName, L"P"))
				continue;

			// search inside current element
			sel->moveToElementText(elem);

			U::ReMatches rm(re->Execute(sel->text));
			if (rm->Count <= 0)
				continue;
			if (incr > 0)
			{
				SelMatch(sel, rm->Item[0]);
				m_startMatch = 1;
			}
			else
			{
				SelMatch(sel, rm->Item[rm->Count - 1]);
				m_endMatch = rm->Count - 2;
			}
			if (fWrapped)
				::MessageBeep(MB_ICONASTERISK);

			return true;
		}

		// search again in starting element
		if ((bool)sc && U::scmp(sc->tagName, L"P") == 0)
		{
			sel->moveToElementText(sc);
			U::ReMatches rm(re->Execute(sel->text));
			if (rm->Count > 0)
			{
				if (incr > 0)
				{
					for (long l = 0; l < rm->Count; ++l)
					{
						U::ReMatch crm(rm->Item[l]);
						if (crm->FirstIndex < s_off1)
						{
							SelMatch(sel, crm);
							if (fWrapped)
								::MessageBeep(MB_ICONASTERISK);

							return true;
						}
					}
				}
				else
				{
					for (long l = rm->Count - 1; l >= 0; --l)
					{
						U::ReMatch crm(rm->Item[l]);
						if (crm->FirstIndex >= s_off2)
						{
							SelMatch(sel, crm);
							if (fWrapped)
								::MessageBeep(MB_ICONASTERISK);
							return true;
						}
					}
				}
			}
		}
	}
	catch (_com_error &err)
	{
		U::ReportError(err);
	}

	return false;
}

bool CFBEView::DoSearchStd(bool fMore)
{
	try
	{
		// fetch selection
		MSHTML::IHTMLTxtRangePtr sel(Document()->selection->createRange());
		if (!fMore && (bool)m_is_start)
			sel = m_is_start->duplicate();
		if (!(bool)sel)
			return false;

		MSHTML::IHTMLTxtRangePtr org(sel->duplicate());
		// check if it is collapsed
		if (sel->compareEndPoints(L"StartToEnd", sel) != 0)
		{
			// collapse and advance
			if (m_fo.flags & FRF_REVERSE)
				sel->collapse(VARIANT_TRUE);
			else
				sel->collapse(VARIANT_FALSE);
		}

		// search for text
		if (sel->findText((const wchar_t *)m_fo.pattern, 1073741824, m_fo.flags) == VARIANT_TRUE)
		{
			// ok, found
			sel->select();
			return true;
		}

		// not found, try searching from start to sel
		sel = MSHTML::IHTMLBodyElementPtr(Document()->body)->createTextRange();
		sel->collapse(m_fo.flags & 1 ? VARIANT_FALSE : VARIANT_TRUE);
		if (sel->findText((const wchar_t *)m_fo.pattern, 1073741824, m_fo.flags) == VARIANT_TRUE && org->compareEndPoints("StartToStart", sel) * (m_fo.flags & 1 ? -1 : 1) > 0)
		{
			// found
			sel->select();
			MessageBeep(MB_ICONASTERISK);
			return true;
		}
	}
	catch (_com_error &)
	{
		//U::ReportError(err);
	}

	return false;
}

static CString GetSM(U::ReSubMatches sm, int idx)
{
	if (!sm)
		return CString();

	if (idx < 0 || idx >= sm->Count)
		return CString();

	_variant_t vt(sm->Item[idx]);

	if (V_VT(&vt) == VT_BSTR)
		return V_BSTR(&vt);

	return CString();
}

struct RR
{
	enum
	{
		STRONG = 1,
		EMPHASIS = 2,
		UPPER = 4,
		LOWER = 8,
		TITLE = 16
	};

	int flags;
	int start;
	int len;
};

typedef CSimpleValArray<RR> RRList;

static CString GetReplStr(const CString &rstr, U::ReMatch rm, RRList &rl)
{
	CString rep;
	rep.GetBuffer(rstr.GetLength());
	rep.ReleaseBuffer(0);

	U::ReSubMatches rs(rm->SubMatches);

	RR cr;
	memset(&cr, 0, sizeof(cr));
	int flags = 0;

	CString rv;
	bool emptyParam = false;

	for (int i = 0; i < rstr.GetLength(); ++i)
	{
		if ((rstr[i] == L'$' && i < rstr.GetLength() - 1) ||
			(rstr[i] == L'\\' && i < rstr.GetLength() - 1))
		{
			switch (rstr[++i])
			{
				//case L'&': // whole match
			case L'0': // whole match
				rv = (const wchar_t *)rm->Value;
				break;
			case L'+': // last submatch
				rv = GetSM(rs, rs->Count - 1);
				break;
			case L'1':
			case L'2':
			case L'3':
			case L'4':
			case L'5':
			case L'6':
			case L'7':
			case L'8':
			case L'9':
				rv = GetSM(rs, rstr[i] - L'0' - 1);
				if (rv.IsEmpty())
					emptyParam = true;
				break;
			case L'T': // title case
				flags |= RR::TITLE;
				continue;
			case L'U': // uppercase
				flags |= RR::UPPER;
				continue;
			case L'L': // lowercase
				flags |= RR::LOWER;
				continue;
			case L'S': // strong
				flags |= RR::STRONG;
				continue;
			case L'E': // emphasis
				flags |= RR::EMPHASIS;
				continue;
			case L'Q': // turn off flags
				flags = 0;
				continue;
			default: // ignore
				continue;
			}
		}

		if (cr.flags != flags && cr.flags && cr.start < rep.GetLength())
		{
			cr.len = rep.GetLength() - cr.start;
			rl.Add(cr);
			cr.flags = 0;
		}

		if (flags)
		{
			cr.flags = flags;
			cr.start = rep.GetLength();
		}

		// SeNS: fix for issue #142
		if (!emptyParam)
		{
			if (!rv.IsEmpty())
			{
				rep += rv;
				rv.Empty();
			}
			else
				rep += rstr[i];
		}
		else
			emptyParam = false;
	}

	if (cr.flags && cr.start < rep.GetLength())
	{
		cr.len = rep.GetLength() - cr.start;
		rl.Add(cr);
	}

	// process case conversions here
	int tl = rep.GetLength();
	TCHAR *cp = rep.GetBuffer(tl);
	for (int j = 0; j < rl.GetSize();)
	{
		RR rr = rl[j];
		if (rr.flags & RR::UPPER)
			LCMapString(CP_ACP, LCMAP_UPPERCASE, cp + rr.start, rr.len, cp + rr.start, rr.len);
		else if (rr.flags & RR::LOWER)
			LCMapString(CP_ACP, LCMAP_LOWERCASE, cp + rr.start, rr.len, cp + rr.start, rr.len);
		else if (rr.flags & RR::TITLE && rr.len > 0)
		{
			LCMapString(CP_ACP, LCMAP_UPPERCASE, cp + rr.start, 1, cp + rr.start, 1);
			LCMapString(CP_ACP, LCMAP_LOWERCASE, cp + rr.start + 1, rr.len - 1, cp + rr.start + 1, rr.len - 1);
		}

		if ((rr.flags & ~(RR::UPPER | RR::LOWER | RR::TITLE)) == 0)
			rl.RemoveAt(j);
		else
			++j;
	}

	rep.ReleaseBuffer(tl);

	return rep;
}

void CFBEView::DoReplace()
{
	try
	{
		MSHTML::IHTMLTxtRangePtr sel(Document()->selection->createRange());
		if (!(bool)sel)
			return;
		MSHTML::IHTMLTxtRangePtr x2(sel->duplicate());
		int adv = 0;

		m_mk_srv->BeginUndoUnit(L"replace");

		if (m_fo.match)
		{ // use regexp match
			RRList rl;
			CString rep(GetReplStr(m_fo.replacement, m_fo.match, rl));

			// added by SeNS
			if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
				rep.Replace(L"\u00A0", _Settings.GetNBSPChar());

			sel->text = (const wchar_t *)rep;
			// change bold/italic where needed
			for (int i = 0; i < rl.GetSize(); ++i)
			{
				RR rr = rl[i];
				x2 = sel->duplicate();
				x2->move(L"character", rr.start - rep.GetLength());
				x2->moveEnd(L"character", rr.len);
				if (rr.flags & RR::STRONG)
					x2->execCommand(L"Bold", VARIANT_FALSE);
				if (rr.flags & RR::EMPHASIS)
					x2->execCommand(L"Italic", VARIANT_FALSE);
			}
			adv = rep.GetLength();
		}
		else
		{ // plain text
			sel->text = (const wchar_t *)m_fo.replacement;
			adv = m_fo.replacement.GetLength();
		}
		sel->moveStart(L"character", -adv);
		sel->select();
	}
	catch (_com_error &e)
	{
		U::ReportError(e);
	}
	m_mk_srv->EndUndoUnit();
}

int CFBEView::GlobalReplace(MSHTML::IHTMLElementPtr elem, CString cntTag)
{
	if (m_fo.pattern.IsEmpty())
		return 0;

	try
	{
		MSHTML::IHTMLTxtRangePtr sel(MSHTML::IHTMLBodyElementPtr(Document()->body)->createTextRange());
		if (elem)
			sel->moveToElementText(elem);
		if (!(bool)sel)
			return 0;

		U::RegExp re;
#ifdef USE_PCRE
		re = new U::IRegExp2();
#else
		CheckError(re.CreateInstance(L"VBScript.RegExp"));
#endif
		re->IgnoreCase = m_fo.flags & 4 ? VARIANT_FALSE : VARIANT_TRUE;
		re->Global = VARIANT_TRUE;

		// added by SeNS
		if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
			m_fo.pattern.Replace(L"\u00A0", _Settings.GetNBSPChar());

		re->Pattern = (const wchar_t *)m_fo.pattern;

		m_mk_srv->BeginUndoUnit(L"replace");

		sel->collapse(VARIANT_TRUE);

		int nRepl = 0;

		if (m_fo.fRegexp)
		{
			MSHTML::IHTMLTxtRangePtr s3;
			MSHTML::IHTMLElementCollectionPtr all;
			if (elem)
				all = MSHTML::IHTMLElement2Ptr(elem)->getElementsByTagName(cntTag.AllocSysString());
			else
				all = MSHTML::IHTMLDocument3Ptr(Document())->getElementsByTagName(cntTag.AllocSysString());
			_bstr_t charstr(L"character");
			RRList rl;
			CString repl;

			for (long l = 0; l < all->length; ++l)
			{
				MSHTML::IHTMLElementPtr elem2(all->item(l));
				sel->moveToElementText(elem2);
				;
				U::ReMatches rm(re->Execute(sel->text));
				if (rm->Count <= 0)
					continue;

				// SeNS: fix for issue #147
				MSHTML::IHTMLTxtRangePtr rng = sel->duplicate();
				CString text = rng->text;
				CString html = rng->htmlText;

				// Replace
				sel->collapse(VARIANT_TRUE);
				long last = 0;
				for (long i = 0; i < rm->Count; ++i)
				{
					U::ReMatch cur(rm->Item[i]);
					long delta = cur->FirstIndex - last;

					// SeNS
					delta += TextOffset(rng, cur, text, html);

					if (delta)
					{
						sel->move(charstr, delta);
						last += delta;
					}
					if (sel->moveStart(charstr, 1) == 1)
						sel->move(charstr, -1);
					delta = cur->Length;
					last += cur->Length;
					sel->moveEnd(charstr, delta);
					rl.RemoveAll();
					repl = GetReplStr(m_fo.replacement, cur, rl);

					// added by SeNS
					if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
						repl.Replace(L"\u00A0", _Settings.GetNBSPChar());

					sel->text = (const wchar_t *)repl;
					for (int k = 0; k < rl.GetSize(); ++k)
					{
						RR rr = rl[k];
						s3 = sel->duplicate();
						s3->move(L"character", rr.start - repl.GetLength());
						s3->moveEnd(L"character", rr.len);
						if (rr.flags & RR::STRONG)
							s3->execCommand(L"Bold", VARIANT_FALSE);
						if (rr.flags & RR::EMPHASIS)
							s3->execCommand(L"Italic", VARIANT_FALSE);
					}
					++nRepl;
				}
			}
		}
		else
		{
			DWORD flags = m_fo.flags & ~FRF_REVERSE;
			_bstr_t pattern((const wchar_t *)m_fo.pattern);
			_bstr_t repl((const wchar_t *)m_fo.replacement);
			while (sel->findText(pattern, 1073741824, flags) == VARIANT_TRUE)
			{
				sel->text = repl;
				++nRepl;
			}
		}

		m_mk_srv->EndUndoUnit();
		return nRepl;
	}
	catch (_com_error &err)
	{
		U::ReportError(err);
	}

	return 0;
}

int CFBEView::ToolWordsGlobalReplace(MSHTML::IHTMLElementPtr fbw_body,
									 int *pIndex,
									 int *globIndex,
									 bool find,
									 CString cntTag)
{
	if (m_fo.pattern.IsEmpty())
		return 0;

	int nRepl = 0;

	try
	{
		U::RegExp re;
#ifdef USE_PCRE
		re = new U::IRegExp2();
#else
		CheckError(re.CreateInstance(L"VBScript.RegExp"));
#endif
		re->IgnoreCase = m_fo.flags & FRF_CASE ? VARIANT_FALSE : VARIANT_TRUE;
		re->Global = m_fo.flags & FRF_WHOLE ? VARIANT_TRUE : VARIANT_FALSE;
		re->Multiline = VARIANT_TRUE;
		re->Pattern = (const wchar_t *)m_fo.pattern;

		MSHTML::IHTMLElementCollectionPtr paras = MSHTML::IHTMLElement2Ptr(fbw_body)->getElementsByTagName(cntTag.AllocSysString());
		if (!paras->length)
			return 0;

		int iNextElem = pIndex != NULL ? *pIndex : 0;
		CSimpleArray<CFBEView::pElAdjacent> pAdjElems;

		while (iNextElem < paras->length)
		{
			pAdjElems.RemoveAll();

			MSHTML::IHTMLElementPtr currElem(paras->item(iNextElem));
			CString innerText = currElem->innerText;
			pAdjElems.Add(pElAdjacent(currElem));

			if (pIndex != NULL)
				*pIndex = iNextElem;

			MSHTML::IHTMLDOMNodePtr currNode(currElem);
			if (MSHTML::IHTMLElementPtr siblElem = currNode->nextSibling)
			{
				int jNextElem = iNextElem + 1;
				for (int i = jNextElem; i < paras->length; ++i)
				{
					MSHTML::IHTMLElementPtr nextElem = paras->item(i);
					if (siblElem == nextElem)
					{
						pAdjElems.Add(pElAdjacent(siblElem));
						innerText += L"\n";
						innerText += siblElem->innerText.GetBSTR();
						iNextElem++;
						siblElem = MSHTML::IHTMLDOMNodePtr(nextElem)->nextSibling;
					}
					else
					{
						break;
					}
				}
			}
			innerText += L"\n";

			if (innerText.IsEmpty())
			{
				iNextElem++;
				continue;
			}

			// Replace
#ifdef USE_PCRE
			U::ReMatches rm(re->Execute(innerText));
#else
			U::ReMatches rm(re->Execute(innerText.AllocSysString()));
#endif
			if (rm->Count <= 0)
			{
				iNextElem++;
				continue;
			}

			for (long i = 0; i < rm->Count; ++i)
			{
				U::ReMatch cur(rm->Item[i]);

				long matchIdx = cur->FirstIndex;
				long matchLen = cur->Length - 1;

				long pAdjLen = 0;
				bool begin = false, end = false;
				int first = 0, last = 0;

				for (int b = 0; b < pAdjElems.GetSize(); ++b)
				{
					int pElemLen = pAdjElems[b].innerText.length() + 1;

					if (!pElemLen)
						continue;

					pAdjLen += pElemLen;

					if (matchIdx < pAdjLen && !begin)
					{
						begin = true;
						first = b;
					}

					if (matchIdx + (matchLen - 1) < pAdjLen && !end)
					{
						end = true;
						last = b;
						break;
					}
				}

				int skip = 0;
				while (skip < first)
				{
					matchIdx -= (pAdjElems[skip].innerText.length() + 1);
					skip++;
				}

				CString newCont;
				int icat = first;
				while (icat <= last)
				{
					newCont += pAdjElems[icat].innerText.GetBSTR();
					icat++;
				}

				if (find)
				{
					if (i == rm->Count - 1)
					{
						(*globIndex) = -1;
						(*pIndex) += (pAdjElems.GetSize());
					}
					else
						(*globIndex)++;

					if (*globIndex > i)
					{
						(*globIndex)--;
						continue;
					}

					MSHTML::IHTMLTxtRangePtr found(Document()->selection->createRange());
					found->moveToElementText(pAdjElems[first].elem);

					// SeNS: fix for issue #148
					found->moveStart(L"character", matchIdx + TextOffset(found, cur));
					found->collapse(TRUE);
					//					int diff = last - first;
					found->moveEnd(L"character", matchLen);
					found->select();

					return 0;
				}
				else
				{
					MSHTML::IHTMLTxtRangePtr found(Document()->selection->createRange());
					found->moveToElementText(pAdjElems[first].elem);

					// SeNS: fix for issue #148
					found->moveStart(L"character", matchIdx + TextOffset(found, cur));
					found->collapse(TRUE);
					//					int diff = last - first;
					found->moveEnd(L"character", matchLen);
					found->select();
					CString strRepl;
					GetDlgItem(IDC_WORDS_FR_EDIT_REPL).GetWindowText(strRepl);

					found->text = L"";
					found->text = m_fo.replacement.AllocSysString();

					newCont.Delete(matchIdx, matchLen - (last - first));
					newCont.Insert(matchIdx, m_fo.replacement);

					pAdjElems[first].innerText = newCont.AllocSysString();
					//pAdjElems[first].elem->innerText = pAdjElems[first].innerText;

					for (int c = first + 1; c <= last; ++c)
					{
						//	MSHTML::IHTMLDOMNodePtr(pAdjElems[c].elem)->removeNode(VARIANT_TRUE);
						iNextElem--;
					}

					for (int c = first + 1; c < last; ++c)
						pAdjElems.RemoveAt(c);

					if (nRepl >= m_fo.replNum)
						goto stop;

					CString again;
					for (int c = 0; c < pAdjElems.GetSize(); ++c)
					{
						//pAdjElems[c].innerText = pAdjElems[c].elem->innerText;
						again += pAdjElems[c].innerText.GetBSTR();
						again += L"\n";
					}
#ifdef USE_PCRE
					rm = re->Execute(again);
#else
					rm = re->Execute(again.AllocSysString());
#endif
					i--;

					nRepl++;
				}
			}

			iNextElem++;
		}

	stop:
#ifndef USE_PCRE
		re.Release();
#endif

		if (find)
		{
			Document()->selection->empty();
			return -1;
		}
	}
	catch (_com_error &err)
	{
		U::ReportError(err);
	}

	return nRepl;
}
//------------CViewReplaceDlg---
class CViewReplaceDlg : public CReplaceDlgBase {
public:
	CViewReplaceDlg(CFBEView *view) : CReplaceDlgBase(view) { }

	virtual void DoFind() {
		if (!m_view->DoSearch())
		{
			CString strMessage;
			strMessage.Format(IDS_SEARCH_END_MSG, (LPCTSTR)m_view->m_fo.pattern);
			AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_WARNING_ICON);
		}
		else {
			SaveString();
			SaveHistory();
			m_selvalid = true;
			MakeClose();
		}
	}
	virtual void DoReplace() {
		if (m_selvalid) { // replace
			m_view->DoReplace();
			m_selvalid = false;
		}
		m_view->m_startMatch = m_view->m_endMatch = 0;
		DoFind();
	}
	virtual void DoReplaceAll()
	{
		int nRepl = m_view->GlobalReplace();

		CString strMessage;
		if (nRepl > 0)
		{
			SaveString();
			SaveHistory();

			strMessage.Format(IDS_REPL_DONE_MSG, nRepl);
			AtlTaskDialog(*this, IDS_REPL_ALL_CAPT, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
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

LRESULT CFBEView::OnFind(WORD, WORD, HWND, BOOL &)
{
	m_fo.pattern = GetSelection();
	if (!m_find_dlg)
		m_find_dlg = new CViewFindDlg(this);

	if (!m_find_dlg->IsValid())
		m_find_dlg->ShowDialog(*this); // show modeless
	else
		m_find_dlg->SetFocus();
	return 0;
}

LRESULT CFBEView::OnReplace(WORD, WORD, HWND, BOOL &)
{
	m_fo.pattern = GetSelection();
	if (!m_replace_dlg)
		m_replace_dlg = new CViewReplaceDlg(this);

	if (!m_replace_dlg->IsValid())
		m_replace_dlg->ShowDialog(*this);
	else
		m_replace_dlg->SetFocus();
	return 0;
}

LRESULT CFBEView::OnFindNext(WORD, WORD, HWND, BOOL &)
{
	if (!DoSearch())
	{
		CString strMessage;
		strMessage.Format(IDS_SEARCH_FAIL_MSG, (LPCTSTR)m_fo.pattern);
		AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_WARNING_ICON);
	}
	return 0;
}

// change notifications
void CFBEView::EditorChanged(int id)
{
	switch (id)
	{
	case FWD_SINK:
		break;
	case BACK_SINK:
		break;
	case RANGE_SINK:
		m_startMatch = m_endMatch = 0;
		if (!m_ignore_changes)
			::SendMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_ED_CHANGED), (LPARAM)m_hWnd);
		break;
	}
}

// DWebBrowserEvents2
void CFBEView::OnDocumentComplete(IDispatch *pDisp, VARIANT *vtUrl)
{
	m_complete = true;
}

SHD::IWebBrowser2Ptr CFBEView::Browser()
{
	return m_browser;
}

MSHTML::IHTMLDocument2Ptr CFBEView::Document()
{
	return m_hdoc;
}

IDispatchPtr CFBEView::Script()
{
	return MSHTML::IHTMLDocumentPtr(m_hdoc)->Script;
}

CString CFBEView::NavURL()
{
	return m_nav_url;
}

bool CFBEView::Loaded()
{
	bool cmp = m_complete;
	m_complete = false;
	return cmp;
}

/// <summary>Set font&colors for BODY view</summary>
void CFBEView::ApplyConfChanges()
{
    try
    {
        MSHTML::IHTMLStylePtr hs = m_hdoc->body->style;
        // set font family
        CString fss(_Settings.GetFont());
        if (!fss.IsEmpty())
            hs->fontFamily = (const wchar_t *)fss;

        // set font size
        DWORD fs = _Settings.GetFontSize();
        if (fs > 1)
        {
            fss.Format(_T("%dpt"), fs);
            hs->fontSize = (const wchar_t *)fss;
        }

        // set fg color
        fs = _Settings.GetColorFG();
        if (fs == CLR_DEFAULT)
            fs = ::GetSysColor(COLOR_WINDOWTEXT);
        fss.Format(_T("rgb(%d,%d,%d)"), GetRValue(fs), GetGValue(fs), GetBValue(fs));
        hs->color = (const wchar_t *)fss;

        // set bg color
        fs = _Settings.GetColorBG();
        if (fs == CLR_DEFAULT)
            fs = ::GetSysColor(COLOR_WINDOW);
        fss.Format(_T("rgb(%d,%d,%d)"), GetRValue(fs), GetGValue(fs), GetBValue(fs));
        hs->backgroundColor = (const wchar_t *)fss;
    }
    catch (_com_error &)
    {
    }
}

long CFBEView::GetVersionNumber()
{
	return m_mkc ? m_mkc->GetVersionNumber() : -1;
}

void CFBEView::BeginUndoUnit(const wchar_t *name)
{
	m_mk_srv->BeginUndoUnit((wchar_t *)name);
}
void CFBEView::EndUndoUnit()
{
	m_mk_srv->EndUndoUnit();
}
///<summary>Init Body</summary>
void CFBEView::Init()
{
	// save document pointer
	m_hdoc = m_browser->Document;

	m_mk_srv = m_hdoc;
	m_mk_srv.AddRef();
	m_mkc = m_hdoc;

	// attach external helper
	//SetExternalDispatch(CreateHelper());

	// fixup all P elements (inflate_block)
	FixupParagraphs(Document()->body);

//	if (m_normalize)
		//Normalize();
/*
	if (!m_normalize)
	{
		// set ID,version, date, programms fields of document-info if they empty
		MSHTML::IHTMLInputElementPtr ii(Document()->all->item(L"diID"));
		if (ii && ii->value.length() == 0)
		{ 
            // generate new ID
			UUID uuid;
			unsigned char *str;
			if (UuidCreate(&uuid) == RPC_S_OK && UuidToStringA(&uuid, &str) == RPC_S_OK)
			{
				CString us(str);
				RpcStringFreeA(&str);
				us.MakeUpper();
				ii->value = (const wchar_t *)us;
			}
		}
		
        ii = Document()->all->item(L"diVersion");
		if (ii && ii->value.length() == 0)
			ii->value = L"1.0";
		
        ii = Document()->all->item(L"diDate");
		MSHTML::IHTMLInputElementPtr jj(Document()->all->item(L"diDateVal"));
		if (ii && jj && ii->value.length() == 0 && jj->value.length() == 0)
		{
			time_t tt;
			time(&tt);
			char buffer[128];
			strftime(buffer, sizeof(buffer), "%Y-%m-%d", localtime(&tt));
			ii->value = buffer;
			jj->value = buffer;
		}
		
        ii = Document()->all->item(L"diProgs");
		if ((bool)ii && ii->value.length() == 0)
			ii->value = L"FB Tools";
	}
	*/
	// attach document events handler
	DocumentEvents::DispEventUnadvise(Document(), &DIID_HTMLDocumentEvents2);
	DocumentEvents::DispEventAdvise(Document(), &DIID_HTMLDocumentEvents2);
	TextEvents::DispEventUnadvise(Document()->body, &DIID_HTMLTextContainerEvents2);
	TextEvents::DispEventAdvise(Document()->body, &DIID_HTMLTextContainerEvents2);

	// turn off browser's d&d
	m_browser->put_RegisterAsDropTarget(VARIANT_FALSE);
	m_initialized = true;
}

void CFBEView::TraceChanges(bool state)
{
	if (state == m_trace_changes)
		return;
	if (state)
	{
		// attach editing changed handlers
		m_mkc->RegisterForDirtyRange((RangeSink *)this, &m_dirtyRangeCookie);
		//		m_ignore_changes = 0;
	/*	if (!m_pChangeLog)
			m_mkc->CreateChangeLog((RangeSink *)this, &m_pChangeLog, true, true);*/
		// added by SeNS
		m_elementsNum = MSHTML::IHTMLElement2Ptr(Document()->body)->getElementsByTagName("*")->length;
	}
	else
	{
		//To stop receiving change notifications, call Release on the IHTMLChangeLog pointer.

		// detach editing changed handlers
/*		if (m_pChangeLog)
		{
			m_pChangeLog->Release();
			m_pChangeLog = nullptr;
		}
		m_ignore_changes++;*/
		m_mkc->UnRegisterForDirtyRange(m_dirtyRangeCookie);
	}
	m_trace_changes = state;
}
void CFBEView::OnBeforeNavigate(IDispatch *pDisp, VARIANT *vtUrl, VARIANT *vtFlags,
								VARIANT *vtTargetFrame, VARIANT *vtPostData,
								VARIANT *vtHeaders, VARIANT_BOOL *fCancel)
{
	if (!m_initialized)
		return;

	if (vtUrl && V_VT(vtUrl) == VT_BSTR)
	{
		m_nav_url = V_BSTR(vtUrl);

		if (m_nav_url.Left(13) == _T("fbw-internal:"))
			return;

		// changed by SeNS: possible fix for issue #87
		// tested on Windows Vista Ultimate
		::PostMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_NAVIGATE), (LPARAM)m_hWnd);
	}

	// disable navigating away
	*fCancel = VARIANT_TRUE;
}

// HTMLDocumentEvents
void CFBEView::OnSelChange(IDispatch *evt)
{
	if (!m_ignore_changes)
		::SendMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_SEL_CHANGE), (LPARAM)m_hWnd);
//	if (m_cur_sel)
//		m_cur_sel.Release();
}

VARIANT_BOOL CFBEView::OnContextMenu(IDispatch *evt)
{
	MSHTML::IHTMLEventObjPtr oe(evt);
	oe->cancelBubble = VARIANT_TRUE;
	oe->returnValue = VARIANT_FALSE;
/*	if (!m_normalize)
	{
		MSHTML::IHTMLElementPtr elem(oe->srcElement);
		if (!(bool)elem)
			return VARIANT_TRUE;
		if (U::scmp(elem->tagName, L"INPUT") && U::scmp(elem->tagName, L"TEXTAREA"))
			return VARIANT_TRUE;
	}*/

	// display custom context menu here
	CMenu menu;
	CString itemName;

	menu.CreatePopupMenu();
	menu.AppendMenu(MF_STRING, ID_EDIT_UNDO, _T("&Undo"));
	menu.AppendMenu(MF_SEPARATOR);

	itemName.LoadString(IDS_CTXMENU_CUT);
	menu.AppendMenu(MF_STRING, ID_EDIT_CUT, itemName);

	itemName.LoadString(IDS_CTXMENU_COPY);
	menu.AppendMenu(MF_STRING, ID_EDIT_COPY, itemName);

	itemName.LoadString(IDS_CTXMENU_PASTE);
	menu.AppendMenu(MF_STRING, ID_EDIT_PASTE, itemName);

/*	if (m_normalize)
	{*/
		MSHTML::IHTMLElementPtr cur(SelectionContainer());
		MSHTML::IHTMLElementPtr initial(cur);
		if (initial)
		{
			menu.AppendMenu(MF_SEPARATOR);
			int cmd = ID_SEL_BASE;
			itemName.LoadString(IDS_CTXMENU_SELECT);

			while ((bool)cur && U::scmp(cur->tagName, L"BODY") && U::scmp(cur->id, L"fbw_body"))
			{
				menu.AppendMenu(MF_STRING, cmd, itemName + L" " + GetPath(cur));
				cur = cur->parentElement;
				++cmd;
			}
			if (U::scmp(initial->className, L"image") == 0)
			{
				MSHTML::IHTMLImgElementPtr image = MSHTML::IHTMLDOMNodePtr(initial)->firstChild;
				CString src = image->src.GetBSTR();
				src.Delete(src.Find(L"fbw-internal:"), 13);
				if (src != L"#undefined")
				{
					menu.AppendMenu(MF_SEPARATOR);
					itemName.LoadString(IDS_CTXMENU_IMG_SAVEAS);
					menu.AppendMenu(MF_STRING, ID_SAVEIMG_AS, itemName);
				}
			}
		}
//	}

	U::TRACKPARAMS tp;
	tp.hMenu = menu;
	tp.uFlags = TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON;
	tp.x = oe->screenX;
	tp.y = oe->screenY;
	::SendMessage(m_frame, U::WM_TRACKPOPUPMENU, 0, (LPARAM)&tp);

	return VARIANT_TRUE;
}

LRESULT CFBEView::OnSelectElement(WORD, WORD wID, HWND, BOOL &)
{
	int steps = wID - ID_SEL_BASE;
	try
	{
		MSHTML::IHTMLElementPtr cur(SelectionContainer());

		while ((bool)cur && steps-- > 0)
			cur = cur->parentElement;

		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
		MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(Document())->createRange();
		range->selectNodeContents(cur);
		++m_ignore_changes;
		sel->removeAllRanges();
		sel->addRange(range);
		--m_ignore_changes;

		/*
		MSHTML::IHTMLTxtRangePtr r(MSHTML::IHTMLBodyElementPtr(Document()->body)->createTextRange());

		r->moveToElementText(cur);

		++m_ignore_changes;
		r->select();
		--m_ignore_changes;
		*/
		//m_cur_sel = cur;
		::SendMessage(m_frame, WM_COMMAND, MAKELONG(0, IDN_SEL_CHANGE), (LPARAM)m_hWnd);
	}
	catch (_com_error &e)
	{
		U::ReportError(e);
	}

	return 0;
}

VARIANT_BOOL CFBEView::OnClick(IDispatch *evt)
{
	MSHTML::IHTMLEventObjPtr oe(evt);
	MSHTML::IHTMLElementPtr elem(oe->srcElement);

	m_startMatch = m_endMatch = 0;

	if (!(bool)elem)
		return VARIANT_FALSE;

	MSHTML::IHTMLElementPtr parent_element = elem->parentElement;

	if (!(bool)parent_element)
		return VARIANT_FALSE;

	bstr_t pc = parent_element->className;

	if (!U::scmp(pc, L"image"))
	{
		// make image selected
		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
		MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(Document())->createRange();
		range->selectNode(parent_element);
		sel->removeAllRanges();
		sel->addRange(range);
		
		/*IHTMLControlRangePtr r(((MSHTML::IHTMLElement2Ptr)(Document()->body))->createControlRange());
		HRESULT hr = r->add((IHTMLControlElementPtr)elem->parentElement);
		hr = r->select();*/
		//::SendMessage(m_frame, WM_COMMAND, MAKELONG(IDC_HREF, IDN_WANTFOCUS), (LPARAM)m_hWnd);
		
		return VARIANT_TRUE;
	}

	if (U::scmp(elem->tagName, L"A"))
	{
		return VARIANT_FALSE;
	}
	/*else
	{
	  ::SendMessage(m_frame, WM_COMMAND, MAKELONG(IDC_HREF, IDN_WANTFOCUS), (LPARAM)m_hWnd);
	  return VARIANT_FALSE;
	}*/

	if (oe->altKey != VARIANT_TRUE || oe->shiftKey == VARIANT_TRUE || oe->ctrlKey == VARIANT_TRUE)
		return VARIANT_FALSE;

	CString sref(U::GetAttrCS(elem, L"href"));
	if (sref.IsEmpty() || sref[0] != L'#')
		return VARIANT_FALSE;

	sref.Delete(0);

	MSHTML::IHTMLElementPtr targ(Document()->all->item((const wchar_t *)sref));

	if (!(bool)targ)
		return VARIANT_FALSE;

	GoTo(targ);

	oe->cancelBubble = VARIANT_TRUE;
	oe->returnValue = VARIANT_FALSE;

	return VARIANT_TRUE;
}

VARIANT_BOOL CFBEView::OnKeyDown(IDispatch *evt)
{
	MSHTML::IHTMLEventObjPtr oe(evt);
	if (oe->keyCode == VK_LEFT || oe->keyCode == VK_UP || oe->keyCode == VK_PRIOR || oe->keyCode == VK_HOME)
		m_startMatch = m_endMatch = 0;
	return VARIANT_TRUE;
}

VARIANT_BOOL CFBEView::OnRealPaste(IDispatch *evt)
{
	MSHTML::IHTMLEventObjPtr oe(evt);
	oe->cancelBubble = VARIANT_TRUE;
	if (!m_enable_paste)
	{
		// Blocks first OnRealPaste to stop double-insertion
		SendMessage(WM_COMMAND, MAKELONG(ID_EDIT_PASTE, 0), 0);
		oe->returnValue = VARIANT_FALSE;
	}
	else
	{
		oe->returnValue = VARIANT_TRUE;
	}

	return VARIANT_TRUE;
}

void __stdcall CFBEView::OnDrop(IDispatch *)
{
//	if (m_normalize)
	Normalize();
}

VARIANT_BOOL __stdcall CFBEView::OnDragDrop(IDispatch *)
{
	return VARIANT_FALSE;
}

/// <summary>Check synched state=no changes after switching from source editor</summary>
/// <returns>true if weren't changes</returns>
bool CFBEView::IsSynched()
{
	return m_synched;
}

/// <summary>(Re)set synched state in true</summary>
void CFBEView::Synched(bool state)
{
	m_synched = state;
}

/// <summary>Trace changes in input elements in DESC view
/// compare input values on focusin\focusout events</summary>
void CFBEView::OnFocusInOut(IDispatch *evt)
{
	if (m_cur_input_element)
	{
		// check previous value in input
		bool cv = GetInputElementValue(m_cur_input_element) != m_cur_val;
		m_cur_input_element.Release();
		m_cur_input_element = nullptr;
		m_decr_form_changed = m_decr_form_changed || cv;
		m_synched = !cv;
	}
	if (m_decr_form_changed)
		return;

	// get focused element from event
    MSHTML::IHTMLEventObjPtr oe(evt);
	if (!oe)
		return;

	m_cur_input_element = oe->srcElement;

	if (!m_cur_input_element)
		return;

	// store just input elements value
	CString tName = m_cur_input_element->tagName;
	if (tName != L"INPUT" != 0 && tName != L"SELECT" && tName != L"TEXTAREA")
	{
		m_cur_input_element.Release();
		m_cur_input_element = nullptr;
		return;
	}
	m_cur_val = GetInputElementValue(m_cur_input_element);
}

// find/replace support for scintilla
bool CFBEView::SciFindNext(HWND src, bool fFwdOnly, bool fBarf)
{
	if (m_fo.pattern.IsEmpty())
		return true;

	int flags = 0;
	if (m_fo.flags & FRF_WHOLE)
		flags |= SCFIND_WHOLEWORD;
	if (m_fo.flags & FRF_CASE)
		flags |= SCFIND_MATCHCASE;
	if (m_fo.fRegexp)
		flags |= SCFIND_REGEXP | SCFIND_POSIX;
	int rev = m_fo.flags & FRF_REVERSE && !fFwdOnly;

	// added by SeNS
	if (_Settings.GetNBSPChar().Compare(L"\u00A0") != 0)
		m_fo.pattern.Replace(L"\u00A0", _Settings.GetNBSPChar());

	DWORD len = ::WideCharToMultiByte(CP_UTF8, 0, m_fo.pattern, m_fo.pattern.GetLength(), NULL, 0, NULL, NULL);
	char *tmp = (char *)malloc(len + 1);
	if (tmp)
	{
		::WideCharToMultiByte(CP_UTF8, 0, m_fo.pattern, m_fo.pattern.GetLength(), tmp, len, NULL, NULL);
		tmp[len] = '\0';
		int p1 = ::SendMessage(src, SCI_GETSELECTIONSTART, 0, 0);
		int p2 = ::SendMessage(src, SCI_GETSELECTIONEND, 0, 0);
		if (p2 > p1 && !rev)
			p1 = p2;
		//   if (p1!=p2 && !rev) ++p1;
		if (rev)
			--p1;
		if (p1 < 0)
			p1 = 0;
		p2 = rev ? 0 : ::SendMessage(src, SCI_GETLENGTH, 0, 0);
		int p3 = p2 == 0 ? ::SendMessage(src, SCI_GETLENGTH, 0, 0) : 0;
		::SendMessage(src, SCI_SETTARGETSTART, p1, 0);
		::SendMessage(src, SCI_SETTARGETEND, p2, 0);
		::SendMessage(src, SCI_SETSEARCHFLAGS, flags, 0);
		// this sometimes hangs in reverse search :)
		int ret = ::SendMessage(src, SCI_SEARCHINTARGET, len, (LPARAM)tmp);
		if (ret == -1)
		{ // try wrap
			if (p1 != p3)
			{
				::SendMessage(src, SCI_SETTARGETSTART, p3, 0);
				::SendMessage(src, SCI_SETTARGETEND, p1, 0);
				::SendMessage(src, SCI_SETSEARCHFLAGS, flags, 0);
				ret = ::SendMessage(src, SCI_SEARCHINTARGET, len, (LPARAM)tmp);
			}
			if (ret == -1)
			{
				free(tmp);
				if (fBarf)
				{
					CString strMessage;
					strMessage.Format(IDS_SEARCH_FAIL_MSG, (LPCTSTR)m_fo.pattern);
					AtlTaskDialog(*this, IDR_MAINFRAME, (LPCTSTR)strMessage, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_WARNING_ICON);
				}
				return false;
			}
			::MessageBeep(MB_ICONASTERISK);
		}
		free(tmp);
		p1 = ::SendMessage(src, SCI_GETTARGETSTART, 0, 0);
		p2 = ::SendMessage(src, SCI_GETTARGETEND, 0, 0);
		::SendMessage(src, SCI_SETSELECTIONSTART, p1, 0);
		::SendMessage(src, SCI_SETSELECTIONEND, p2, 0);
		::SendMessage(src, SCI_SCROLLCARET, 0, 0);
		return true;
	}
	else
	{
		AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, IDS_OUT_OF_MEM_MSG, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
	}

	return false;
}

//Get selection text
CString CFBEView::GetSelection()
{
	MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
	if (!sel)
		return L"";
	if (sel->isCollapsed)
		return L"";
	else
		return sel->toString().GetBSTR();
}

// Modification by Pilgrim
/*static bool IsTable(MSHTML::IHTMLDOMNode *node)
{
	MSHTML::IHTMLElementPtr elem(node);
	return U::scmp(elem->className, L"table") == 0;
}

static bool IsTR(MSHTML::IHTMLDOMNode *node)
{
	MSHTML::IHTMLElementPtr elem(node);
	return U::scmp(elem->className, L"tr") == 0;
}

static bool IsTH(MSHTML::IHTMLDOMNode *node)
{
	MSHTML::IHTMLElementPtr elem(node);
	return U::scmp(elem->className, L"th") == 0;
}

static bool IsTD(MSHTML::IHTMLDOMNode *node)
{
	MSHTML::IHTMLElementPtr elem(node);
	return U::scmp(elem->className, L"td") == 0;
}
*/
bool CFBEView::GoToFootnote(bool fCheck)
{
	// * create selection range
	MSHTML::IHTMLTxtRangePtr rng(Document()->selection->createRange());
	if (!(bool)rng)
		return false;

	MSHTML::IHTMLTxtRangePtr next_rng = rng->duplicate();
	MSHTML::IHTMLTxtRangePtr prev_rng = rng->duplicate();
	next_rng->moveEnd(L"character", +1);
	prev_rng->moveStart(L"character", -1);

	CString sref(U::GetAttrCS(SelectionsStruct(L"anchor"), L"href"));
	if (sref.IsEmpty())
		sref = U::GetAttrCS(SelectionAnchor(next_rng->parentElement()), L"href");
	if (sref.IsEmpty())
		sref = U::GetAttrCS(SelectionAnchor(prev_rng->parentElement()), L"href");

	if (sref.Find(L"file") == 0)
		sref = sref.Mid(sref.ReverseFind(L'#'), 1024);
	if (sref.IsEmpty() || sref[0] != _T('#'))
		return false;

	// * ok, all checks passed
	if (fCheck)
		return true;

	sref.Delete(0);

	MSHTML::IHTMLElementPtr targ(Document()->all->item((const wchar_t *)sref));

	if (!(bool)targ)
		return false;

	MSHTML::IHTMLDOMNodePtr childNode;
	MSHTML::IHTMLDOMNodePtr node(targ);
	if (!(bool)node)
		return false;

	// added by SeNS: move caret to the footnote text
	if (!U::scmp(node->nodeName, L"DIV") && !U::scmp(targ->className, L"section"))
	{
		if (node->firstChild)
		{
			childNode = node->firstChild;
			while (childNode && !U::scmp(childNode->nodeName, L"DIV") &&
				   (!U::scmp(MSHTML::IHTMLElementPtr(childNode)->className, L"image") ||
					!U::scmp(MSHTML::IHTMLElementPtr(childNode)->className, L"title")))
				childNode = childNode->nextSibling;
		}
	}
	if (!childNode)
		childNode = node;
	if (childNode)
	{
		GoTo(MSHTML::IHTMLElementPtr(childNode));
		targ->scrollIntoView(true);
	}

	return true;
}
bool CFBEView::GoToReference(bool fCheck)
{
	// * create selection range
	MSHTML::IHTMLTxtRangePtr rng(Document()->selection->createRange());
	if (!(bool)rng)
		return false;

	// * get its parent element
	MSHTML::IHTMLElementPtr pe(GetHP(rng->parentElement()));
	if (!(bool)pe)
		return false;

	if (rng->compareEndPoints(L"StartToEnd", rng) != 0)
		return false;

	while ((bool)pe && (U::scmp(pe->tagName, L"DIV") != 0 || U::scmp(pe->className, L"section") != 0))
		pe = pe->parentElement; // Find parent division
	if (!(bool)pe)
		return false;

	MSHTML::IHTMLElementPtr body = pe->parentElement;

	while ((bool)body && (U::scmp(body->tagName, L"DIV") != 0 || U::scmp(body->className, L"body") != 0))
		body = body->parentElement; // Find body

	if (!(bool)body)
		return false;

	CString id = MSHTML::IHTMLElementPtr(rng->parentElement())->id;
	CString sfbname(U::GetAttrCS(body, L"fbname"));
	if (id.IsEmpty() && (sfbname.IsEmpty() || !(sfbname.CompareNoCase(L"notes") == 0 || sfbname.CompareNoCase(L"comments") == 0)))
		return false;
	id = L"#" + id;

	// * ok, all checks passed
	if (fCheck)
		return true;

	MSHTML::IHTMLElement2Ptr elem(Document()->body);
	MSHTML::IHTMLElementCollectionPtr coll(elem->getElementsByTagName(L"A"));
	if (!(bool)coll || coll->length == 0)
	{
		AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, IDS_GOTO_REF_FAIL_MSG, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_INFORMATION_ICON);
		return false;
	}

	for (long l = 0; l < coll->length; ++l)
	{
		MSHTML::IHTMLElementPtr a(coll->item(l));
		if (!(bool)a)
			continue;

		CString href(U::GetAttrCS((MSHTML::IHTMLElementPtr)coll->item(l), L"href"));

		// changed by SeNS
		if (href.Find(L"file") == 0)
			href = href.Mid(href.ReverseFind(L'#'), 1024);
		else if (href.Find(_T("://"), 0) != -1)
			continue;

		CString snote = L"#" + pe->id;

		if (href == snote || href == id)
		{
			GoTo(a);
			MSHTML::IHTMLTxtRangePtr r(MSHTML::IHTMLBodyElementPtr(Document()->body)->createTextRange());
			r->moveToElementText(a);
			r->collapse(VARIANT_TRUE);
			// move selection to position after reference
			CString sa = a->innerText;
			r->move(L"character", sa.GetLength());
			r->select();
			// scroll to the center of view
			MSHTML::IHTMLRectPtr rect = MSHTML::IHTMLElement2Ptr(a)->getBoundingClientRect();
			MSHTML::IHTMLWindow2Ptr window(MSHTML::IHTMLDocument2Ptr(Document())->parentWindow);
			if (rect && window)
			{
				if (rect->bottom - rect->top <= _Settings.GetViewHeight())
					window->scrollBy(0, (rect->top + rect->bottom - _Settings.GetViewHeight()) / 2);
				else
					window->scrollBy(0, rect->top);
			}
			break;
		}
	}

	return false;
}

LRESULT CFBEView::OnEditInsertTable(WORD wNotifyCode, WORD wID, HWND hWndCtl)
{
	CTableDlg dlg;
	if (dlg.DoModal() == IDOK)
	{
		int nRows = dlg.m_nRows;
		bool bTitle = dlg.m_bTitle;
		InsertTable(false, bTitle, nRows);
	}
	return 0;
}

LRESULT CFBEView::OnEditInsImage(WORD, WORD cmdID, HWND, BOOL&)
{
	// added by SeNS
//	bool bInline = (cmdID != ID_EDIT_INS_IMAGE);

	if (_Settings.m_insimage_ask)
	{
		CTaskDialog imgDialog;
		imgDialog.SetWindowTitle(IDR_MAINFRAME);
		imgDialog.SetMainInstructionText(IDS_ADD_CLEARIMG_TEXT);
		imgDialog.SetCommonButtons(TDCBF_YES_BUTTON | TDCBF_NO_BUTTON);
		imgDialog.SetMainIcon(TD_WARNING_ICON);
		imgDialog.SetVerificationText(IDS_DONT_ASK_AGAIN);
		if (!_Settings.m_insimage_ask)
			imgDialog.m_tdc.dwFlags |= TDF_VERIFICATION_FLAG_CHECKED;
		int nButton;
		BOOL fVerificationFlagChecked;
		imgDialog.DoModal(*this, &nButton, NULL, &fVerificationFlagChecked);
		_Settings.SetIsInsClearImage(nButton == IDYES ? true : false);
		_Settings.SetInsImageAsking(fVerificationFlagChecked == TRUE ? false : true);
	}

	if (!_Settings.m_ins_clear_image)
	{
		CFileDialog dlg(
			TRUE,
			NULL,
			NULL,
			OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR,
			L"FBE supported (*.jpg;*.jpeg;*.png)\0*.jpg;*.jpeg;*.png\0JPEG (*.jpg)\0*.jpg\0PNG (*.png)\0*.png\0Bitmap (*.bmp"\
			L")\0*.bmp\0GIF (*.gif)\0*.gif\0TIFF (*.tif)\0*.tif\0\0"
		);
		dlg.m_ofn.Flags &= ~OFN_ENABLEHOOK;
		dlg.m_ofn.lpfnHook = NULL;

		wchar_t dlgTitle[MAX_LOAD_STRING + 1];
		::LoadString(_Module.GetResourceInstance(), IDS_ADD_IMAGE_FILEDLG, dlgTitle, MAX_LOAD_STRING);
		dlg.m_ofn.lpstrTitle = dlgTitle;
		dlg.m_ofn.nFilterIndex = 1;

		if (dlg.DoModal(*this) == IDOK)
		{
			AddImage(dlg.m_szFileName);
		}
	}
	else
	{
		Call(L"InsImage");
	}

	return 0;
}

LRESULT CFBEView::OnStyleNolink(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_UNLINK);
}

LRESULT CFBEView::OnStyleNormal(WORD, WORD, HWND, BOOL &)
{
	return ExecCommand(IDM_REMOVEFORMAT);
}

LRESULT CFBEView::OnStyleTextAuthor(WORD, WORD, HWND, BOOL &)
{
	Call(L"StyleTextAuthor", SelectionsStruct(L""));
	return 0;
}

LRESULT CFBEView::OnEditAddSubtitle(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddSubtitle");
	return 0;
}

LRESULT CFBEView::OnStyleSubtitle(WORD, WORD, HWND, BOOL &)
{
	Call(L"StyleSubtitle", SelectionsStruct(L""));
	return 0;
}
LRESULT CFBEView::OnViewHTML(WORD, WORD, HWND, BOOL &)
{
	IOleCommandTargetPtr ct(m_browser);
	if (ct)
		ct->Exec(&CGID_MSHTML, IDM_VIEWSOURCE, 0, NULL, NULL);
	return 0;
}

LRESULT CFBEView::OnEditAddTitle(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddTitle");
	return 0;
}
LRESULT CFBEView::OnEditAddEpigraph(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddEpigraph", SelectionsStruct(L""));
	return 0;
}
LRESULT CFBEView::OnEditAddBody(WORD, WORD, HWND, BOOL &)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt1(false);
		dd.Invoke1(L"AddBody", &vt1);
	}
	catch (_com_error &)
	{
	}
	//TO_DO goto last body
	return 0;
}
LRESULT CFBEView::OnEditAddBodyNotes(WORD, WORD, HWND, BOOL &)
{
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t vt1(true);
		dd.Invoke1(L"AddBody", &vt1);
	}
	catch (_com_error &)
	{
	}
	//TO_DO goto last body
	return 0;
}
LRESULT CFBEView::OnEditAddTA(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddTextAuthor");
	return 0;
}
LRESULT CFBEView::OnEditClone(WORD, WORD, HWND, BOOL &)
{
	Call(L"CloneContainer", SelectionsStruct(L""));
	return 0;
}

LRESULT CFBEView::OnEditAddAnn(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddAnnotation");
	return 0;
}

LRESULT CFBEView::OnEditAddPoem(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddPoem");
	return 0;
}

LRESULT CFBEView::OnEditAddStanza(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddStanza");
	return 0;
}

LRESULT CFBEView::OnEditAddSection(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddSection");
	return 0;
}

LRESULT CFBEView::OnEditMerge(WORD, WORD, HWND, BOOL &)
{
	Call(L"MergeContainers", SelectionsStruct(L""));
	return 0;
}
LRESULT CFBEView::OnEditSplit(WORD, WORD, HWND, BOOL &)
{
	SplitContainer(false);
	return 0;
}

LRESULT CFBEView::OnEditAddCite(WORD, WORD, HWND, BOOL &)
{
	Call(L"AddCite");
	return 0;
}
LRESULT CFBEView::OnEditRemoveOuter(WORD, WORD, HWND, BOOL &)
{
	Call(L"RemoveOuterContainer", SelectionsStruct(L""));
	return 0;
}

LRESULT CFBEView::OnSaveImageAs(WORD, WORD, HWND, BOOL &)
{
	CString src;
	MSHTML::IHTMLImgElementPtr image = MSHTML::IHTMLDOMNodePtr(SelectionContainer())->firstChild;
	src = image->src.GetBSTR();
	src.Delete(src.Find(L"fbw-internal:#"), 14);

	_variant_t data;
	try
	{
		CComDispatchDriver dd(Script());
		_variant_t arg(src);
		if (!SUCCEEDED(dd.Invoke1(L"GetImageData", &arg, &data)) || data.vt == VT_EMPTY)
			return 0;
	}
	catch (_com_error &)
	{
		return 0;
	}

	CFileDialog imgSaveDlg(FALSE, NULL, src);
	if (m_file_path != L"")
		imgSaveDlg.m_ofn.lpstrInitialDir = m_file_path;

	// add file types
	imgSaveDlg.m_ofn.lpstrFilter = L"JPEG files (*.jpg)\0*.jpg\0PNG files (*.png)\0*.png\0All files (*.*)\0*.*\0\0";
	imgSaveDlg.m_ofn.nFilterIndex = 0;
	imgSaveDlg.m_ofn.lpstrDefExt = L"jpg";
	imgSaveDlg.m_ofn.Flags &= ~OFN_ENABLEHOOK;
	imgSaveDlg.m_ofn.lpfnHook = NULL;

	if (imgSaveDlg.DoModal(m_hWnd) == IDOK)
	{
		HANDLE imgFile = ::CreateFile(imgSaveDlg.m_szFileName,
									  GENERIC_WRITE,
									  NULL,
									  NULL,
									  CREATE_ALWAYS,
									  FILE_ATTRIBUTE_NORMAL,
									  NULL);
		long elnum = 0;
		void *pData;
		::SafeArrayPtrOfIndex(data.parray, &elnum, &pData);
		long size = 0;
		::SafeArrayGetUBound(data.parray, 1, &size);
		DWORD written;
		::WriteFile(imgFile, pData, size, &written, NULL);

		CloseHandle(imgFile);
	}

	return 0;
}

bool CFBEView::InsertTable(bool fCheck, bool bTitle, int nrows)
{
	try
	{
		// * create selection range
		MSHTML::IHTMLTxtRangePtr rng(Document()->selection->createRange());
		if (!(bool)rng)
			return false;

		// * get its parent element
		MSHTML::IHTMLElementPtr pe(GetHP(rng->parentElement()));
		if (!(bool)pe)
			return false;

		// * get parents for start and end ranges and ensure they are the same as pe
		MSHTML::IHTMLTxtRangePtr tr(rng->duplicate());
		tr->collapse(VARIANT_TRUE);
		if (GetHP(tr->parentElement()) != pe)
			return false;

		// * check if it possible to insert a table there
		_bstr_t cls(pe->className);
		if (U::scmp(cls, L"section") && U::scmp(cls, L"epigraph") &&
			U::scmp(cls, L"annotation") && U::scmp(cls, L"history") && U::scmp(cls, L"cite"))
			return false;

		// * ok, all checks passed
		if (fCheck)
			return true;

		// at this point we are ready to create a table

		// * create an undo unit
		m_mk_srv->BeginUndoUnit(L"insert table");

		MSHTML::IHTMLElementPtr te(Document()->createElement(L"DIV"));

		for (int row = nrows; row != -1; --row)
		{
			// * create tr
			MSHTML::IHTMLElementPtr tre(Document()->createElement(L"DIV"));
			tre->className = L"tr";
			// * create th and td
			MSHTML::IHTMLElementPtr the(Document()->createElement(L"P"));
			if (row == 0)
			{
				if (bTitle)
				{							//Нужен заголовок таблицы
					the->className = L"th"; // * create th - заголовок
					MSHTML::IHTMLElement2Ptr(tre)->insertAdjacentElement(L"afterBegin", the);
				}
			}
			else
			{
				the->className = L"td"; // * create td - строки
				MSHTML::IHTMLElement2Ptr(tre)->insertAdjacentElement(L"afterBegin", the);
			}

			// * create table
			te->className = L"table";
			MSHTML::IHTMLElement2Ptr(te)->insertAdjacentElement(L"afterBegin", tre);
		}

		// * paste the results back
		rng->pasteHTML(te->outerHTML);

		// * ensure we have good html
		RelocateParagraphs(MSHTML::IHTMLDOMNodePtr(pe));
		FixupParagraphs(pe);

		// * close undo unit
		m_mk_srv->EndUndoUnit();
	}
	catch (_com_error &e)
	{
		U::ReportError(e);
	}
	return false;
}

long CFBEView::InsertCode()
{
	if (bCall(L"IsCode", SelectionsStruct(L"code")))
	{
		HRESULT hr;
		BeginUndoUnit(L"insert code");
		hr = m_mk_srv->RemoveElement(SelectionsStruct(L"code"));
		EndUndoUnit();
		return hr == S_OK ? 0 : -1;
	}
	else
	{
		try
		{
			BeginUndoUnit(L"insert code");

			int offset = -1;
			MSHTML::IHTMLTxtRangePtr rng(Document()->selection->createRange());
			if (!(bool)rng)
				return -1;

			CString rngHTML((wchar_t *)rng->htmlText);

			// empty selection case - select current word
			if (rngHTML == L"")
			{
				// select word
				rng->moveStart(L"word", -1);
				CString txt = rng->text;
				offset = txt.GetLength();
				rng->expand(L"word");
				rngHTML.SetString(rng->htmlText);
			}

			if (iswspace(rngHTML[rngHTML.GetLength() - 1]))
			{
				rng->moveEnd(L"character", -1);
				rngHTML.SetString(rng->htmlText);
				if (offset > rngHTML.GetLength())
					offset--;
			}

			// save selection
			MSHTML::IMarkupPointerPtr selBegin, selEnd;
			m_mk_srv->CreateMarkupPointer(&selBegin);
			m_mk_srv->CreateMarkupPointer(&selEnd);
			m_mk_srv->MovePointersToRange(rng, selBegin, selEnd);

			if (rngHTML.Find(L"<P") != -1)
			{
				MSHTML::IHTMLElementPtr spanElem = Document()->createElement(L"<SPAN class=code>");
				MSHTML::IHTMLElementPtr selElem = rng->parentElement();

				MSHTML::IHTMLTxtRangePtr rngStart = rng->duplicate();
				MSHTML::IHTMLTxtRangePtr rngEnd = rng->duplicate();
				rngStart->collapse(true);
				rngEnd->collapse(false);

				MSHTML::IHTMLElementPtr elBegin = rngStart->parentElement(), elEnd = rngEnd->parentElement();
				while (U::scmp(elBegin->tagName, L"P"))
					elBegin = elBegin->parentElement;
				while (U::scmp(elEnd->tagName, L"P"))
					elEnd = elEnd->parentElement;

				MSHTML::IHTMLDOMNodePtr bNode = elBegin, eNode = elEnd;

				while (bNode)
				{
					CString elBeginHTML = elBegin->innerHTML;

					if (U::scmp(elBegin->tagName, L"P") == 0 && elBeginHTML.Find(L"<SPAN") < 0)
					{
						spanElem->innerHTML = elBegin->innerHTML;
						if (!(elBeginHTML.Find(L"<SPAN class=code>") == 0 &&
							  elBeginHTML.Find(L"</SPAN>") == elBeginHTML.GetLength() - 7))
						{
							elBegin->innerHTML = spanElem->outerHTML;
						}
					}
					// remove code tag
					else
					{
						elBeginHTML.Replace(L"<SPAN class=code>", L" ");
						elBeginHTML.Replace(L"</SPAN>", L" ");
						elBegin->innerHTML = elBeginHTML.AllocSysString();
					}

					if (bNode == eNode)
						break;

					bNode = bNode->nextSibling;
					elBegin = bNode;
				}
				// expand selection to the last paragraph
				rng->moveToElementText(elBegin);
				m_mk_srv->MovePointersToRange(rng, NULL, selEnd);
			}
			else if (rngHTML.Find(L"<SPAN class=code>") != -1 && rngHTML.Find(L"</SPAN>") != -1)
			{
				rngHTML.Replace(L"<SPAN class=code>", L" ");
				rngHTML.Replace(L"</SPAN>", L" ");
				rng->pasteHTML(rngHTML.AllocSysString());
			}
			else
			{
				if (iswspace(rngHTML[0]))
				{
					rng->moveStart(L"character", 1);
					rngHTML.SetString(rng->htmlText);
				}
				if (iswspace(rngHTML[rngHTML.GetLength() - 1]))
				{
					rng->moveEnd(L"character", -1);
					rngHTML.SetString(rng->htmlText);
				}
				rngHTML = L"<SPAN class=code>" + rngHTML + L"</SPAN>";
				rng->pasteHTML(rngHTML.AllocSysString());
			}

			// restore selection
			if (offset >= 0)
			{
				rng->move(L"word", -1);
				rng->move(L"character", offset);
				rng->select();
			}
			else
			{
				m_mk_srv->MoveRangeToPointers(selBegin, selEnd, rng);
				rng->select();
			}

			EndUndoUnit();
		}
		catch (_com_error &e)
		{
			U::ReportError(e);
			return -1;
		}

		return 0;
	}
}

MSHTML::IHTMLDOMRangePtr CFBEView::SetSelection(MSHTML::IHTMLElementPtr begin, MSHTML::IHTMLElementPtr end, int begin_pos, int end_pos)
{
	if (!(bool)begin)
	{
		return nullptr;
	}

	int b_pos = begin_pos;
	int e_pos = end_pos;
	MSHTML::IHTMLDOMNodePtr b_node = begin;
	MSHTML::IHTMLDOMNodePtr e_node;
	
	// Calculate real offset in text nodes
	GetRealCharPos1(b_node, b_pos);
	
	if (!(bool)end || (begin == end && begin_pos == end_pos))
	{
		e_node = b_node;
		e_pos = b_pos;
	}
	else
	{
		// if end node not equal start node
		e_node = end;
		GetRealCharPos1(e_node, e_pos);
	}

	MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(Document())->createRange();
	if (!range)
		return nullptr;
	MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(Document())->getSelection();
	if (!sel)
		return nullptr;
	//different offset for text and element node
	if (b_node->nodeType == NODE_TEXT)
		range->setStart(b_node, b_pos);
	else
	{
		// if offset =0 set to start of node, else to end of node
		if(b_pos == 0)
			range->setStartBefore(b_node);
		else
			range->setStartAfter(b_node);
	}
	if (e_node->nodeType == NODE_TEXT)
		range->setEnd(e_node, e_pos);
	else
	{
		// if offset =0 set to start of node, else to end of node
		if (e_pos == 0)
			range->setEndBefore(e_node);
		else
			range->setEndAfter(e_node);
	}
	sel->removeAllRanges();
	sel->addRange(range);
	return range;
/*


	begin_pos = this->GetRealCharPos(begin, begin_pos);
	end_pos = this->GetRealCharPos(end, end_pos);

	MSHTML::IHTMLTxtRangePtr rng(MSHTML::IHTMLBodyElementPtr(Document()->body)->createTextRange());
	if (!(bool)rng)
	{
		return 0;
	}

	// устанавливаем начало выделенной строки
	MSHTML::IHTMLTxtRangePtr rng_begin(rng->duplicate());
	HRESULT hr;
	hr = rng_begin->moveToElementText(begin);
	rng_begin->move(L"character", begin_pos);
	//rng_begin->moveStart(L"character", begin_pos);
	//hr = rng_begin->collapse(VARIANT_TRUE);
	
	if (begin == end)
	{
		hr = rng_begin->moveEnd(L"character", end_pos - begin_pos);
		hr = rng_begin->select();

		return rng_begin;
	}

	MSHTML::IHTMLTxtRangePtr rng_end(rng->duplicate());
	rng_end->moveToElementText(end);
	rng_end->moveStart(L"character", end_pos);

	// раздвигаем регион
	rng_begin->setEndPoint(L"EndToStart", rng_end);
	CString ss2 = rng_begin->text;

	rng_begin->select();

	return rng_begin;*/
}

/// <summary>Calculate real offset in text nodes</summary>
/// <params name="node">Source range</params>
/// <params name="pos">Source range</params>
/// <returns>true if successfully expandedselection
/// &rng - updated range
/// &begin - first element of range
/// &end - last element of range</returns>
void CFBEView::GetRealCharPos1(MSHTML::IHTMLDOMNodePtr& node, int& pos)
{
	if (!node)
		return;

//	int realpos = pos;
//	int cuttedchars = 0;
	MSHTML::IHTMLDOMNodePtr source_node = node;

	node = node->firstChild;
	while (node)
	{
		if (node->nodeType == NODE_TEXT)
		{
			if (pos <= MSHTML::IHTMLDOMTextNodePtr(node)->length)
				return;
			else
				pos -= MSHTML::IHTMLDOMTextNodePtr(node)->length;
		}
		node = node->nextSibling;
	}
	node = source_node;
}

///<summary>Count characters in text part of node including children nodes</summary>
///<params name="node">HTML Node</params>
///<returns>Number of characters</returns>
int CFBEView::CountNodeChars(MSHTML::IHTMLDOMNodePtr node)
{
	if (!node)
		return 0;

	int count = 0;

	node = node->firstChild;
	while (node)
	{
		MSHTML::IHTMLDOMTextNodePtr textNode(node);
		if (!textNode)
			count += CountNodeChars(node);
		else
			count += textNode->length;
		node = node->nextSibling;
	}

	return count;
}

bool CFBEView::CloseFindDialog(CFindDlgBase *dlg)
{
	if (!dlg || !dlg->IsValid())
		return false;

	dlg->DestroyWindow();
	return true;
}

bool CFBEView::CloseFindDialog(CReplaceDlgBase* dlg)
{
	if (!dlg || !dlg->IsValid())
		return false;

	dlg->DestroyWindow();
	return true;
}

bool CFBEView::CloseReplaceDialog(CReplaceDlgBase* dlg)
{
	if (!dlg || !dlg->IsValid())
		return false;

	dlg->DestroyWindow();
	return true;
}
/*
bool CFBEView::CloseFindReplaceDialog(CModelessDialog* dlg)
{
	if (!dlg || !dlg->IsValid())
		return false;

	dlg->DestroyWindow();
	return true;
}
*/

LRESULT CFBEView::OnCode(WORD wCode, WORD wID, HWND hWnd, BOOL &bHandled)
{
	return InsertCode();
}

/// <summary>Add image to document</summary>
/// <params name="filename">Filename</params>
/// <returns>Image id</returns>
CString CFBEView::AttachImage(const CString &filename)
{
	HRESULT hr;
	CComVariant args[4];
	CComVariant ret;
	CComDispatchDriver dd(Script());

	args[3] = filename;
	if (FAILED(hr = U::LoadImageFile(filename, &args[0])))
	{
		U::ReportError(hr);
		return L"";
	}
	args[2] = U::PrepareDefaultId(filename);
	args[1] = U::GetMimeType(filename);

	if (FAILED(hr = dd.InvokeN(L"apiAddBinary", args, 4, &ret)))
	{
		U::ReportError(hr);
		return L"";
	}

	if (FAILED(hr = dd.Invoke0(L"FillCoverList")))
	{
		U::ReportError(hr);
		return L"";
	}
	return ret;
}

// images
void CFBEView::AddImage(const CString &filename)
{
	HRESULT hr;
	CComVariant ret;

	// Stuff the thing into JavaScript
	try
	{
		CString s_id = AttachImage(filename);
		if (s_id.IsEmpty())
			return;

		CComDispatchDriver body(Script());
		_variant_t id = s_id;
	
		hr = body.Invoke1(L"InsImage", &id, &ret);
		
		if (FAILED(hr))
			U::ReportError(hr);

		//TO_DO catch return value!
/*		MSHTML::IHTMLDOMNodePtr node(ret);
		if (node)
			BubbleUp(node, L"DIV");*/
	}
	catch (_com_error &)
	{
	}
}

/// <summary>Check is HTML-body structure is changed
/// Count number of HTML-elements</summary>
/// <returns>true if changed</returns>
bool CFBEView::IsHTMLChanged()
{
	long newElementsNum = MSHTML::IHTMLElement2Ptr(Document()->body)->getElementsByTagName("*")->length;
	bool b = (newElementsNum != m_elementsNum);
	if (b)
		m_startMatch = m_endMatch = 0;
	m_elementsNum = newElementsNum;
	return b;
}
