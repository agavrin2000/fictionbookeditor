// Doc.cpp: implementation of the Doc class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "res1.h"

#include "utils.h"

#include "FBDoc.h"
#include "Scintilla.h"
#include "Settings.h"
#include "ElementDescriptor.h"

extern CElementDescMnr _EDMnr;

extern CSettings _Settings;

const CString NOBREAK_SPACE(L"\u00A0");

namespace FB
{

	// namespaces
	const _bstr_t NEWLINE(L"\n");
	const int INDENT_SIZE = 1;

	// document list
	Doc *FB::Doc::m_active_doc;
	//bool FB::Doc::m_fast_mode;

	//-------------- Functions to convert HTML BODY to XML  ---------------
	static bool tFindCloseTag(wchar_t* start, wchar_t* tag_name, wchar_t** close_tag);

	///<summary> Indent node</summary>
	/// <params name="node">XML Node</params>
	/// <params name="xml">XML-document</params>
	static void Indent(MSXML2::IXMLDOMNodePtr node, MSXML2::IXMLDOMDocument2Ptr xml, int len)
	{
		/*CString textIndent(L"\r\n");
		textIndent += CString(L' ', len);
		node->appendChild(xml->createTextNode(textIndent));*/
		// inefficient
		BSTR s = SysAllocStringLen(NULL, len + 2);
		if (s)
		{
			s[0] = L'\r';
			s[1] = L'\n';
			for (BSTR p = s + 2, q = s + 2 + len; p < q; ++p)
				*p = L' ';
			node->appendChild(xml->createTextNode(s));
			SysFreeString(s);
		}
	}
	
    ///<summary>Add attribute to element with namespace if value isn't empty</summary>
	/// <params name="xe">XML Node</params>
	/// <params name="name">Attribute name</params>
	/// <params name="ns">Attribute namespace</params>
	/// <params name="val">Attribute value</params>
	/// <params name="doc">XML-document</params>
	static void SetAttr2(MSXML2::IXMLDOMElementPtr xe, const wchar_t *name,	const wchar_t *ns, const _bstr_t &val, MSXML2::IXMLDOMDocument2Ptr doc)
	{
        if(val.length() == 0)
            return;
		MSXML2::IXMLDOMAttributePtr attr(doc->createNode(NODE_ATTRIBUTE, name, ns));
		attr->appendChild(doc->createTextNode(val));
		xe->setAttributeNode(attr);
	}

	///<summary>Build internal element (inside P)</summary>
	/// <params name="inl">Processed HTML element</params>
	/// <params name="doc">XML-document</params>
	///<returns>Builded IXMLDOMNodePtr with all internal nodes<returns>
	static MSXML2::IXMLDOMNodePtr ProcessInline(MSHTML::IHTMLDOMNodePtr inl, MSXML2::IXMLDOMDocument2Ptr doc)
	{
        CString name = inl->nodeName;
		MSHTML::IHTMLElementPtr einl(inl);
        CString cls = einl->className;

		CString xname = L"";

		// define name for xml-element
		if (name == L"STRONG")
			xname = L"strong";
		else if (name == L"EM")
			xname = L"emphasis";
		else if (name == L"STRIKE")
			xname = L"strikethrough";
		else if (name == L"SUB")
			xname = L"sub";
		else if (name == L"SUP")
			xname = L"sup";
		else if (name == L"A")
			xname = L"a";
		else if (name == L"SPAN")
		{
			if (cls == L"unknown_element")
				xname = einl->getAttribute(L"source_class", 2); //2=return BSTR
			else if (cls == L"image")
				xname = L"image";
			else if (cls == L"code")
				xname = L"code";
			else
				xname = L"style";
		}

		MSXML2::IXMLDOMElementPtr xinl(doc->createNode(NODE_ELEMENT, xname.GetString(), U::FBNS));
        
        // Process attributes
		if (xname == L"image" || xname == L"a")
			SetAttr2(xinl, L"l:href", U::XLINKNS, U::GetAttrB(einl, L"href"), doc);

		if (xname == L"a" && cls == L"note")
			SetAttr2(xinl, L"type", U::FBNS, cls.GetString(), doc);

		if (xname == L"style")
			SetAttr2(xinl, L"name", U::FBNS, cls.GetString(), doc);

		// Unknown element
		if (name == L"SPAN" && cls == L"unknown_element")
		{
			MSHTML::IHTMLAttributeCollectionPtr col = inl->attributes;
			for (int i = 0; i < col->length; ++i)
			{
				VARIANT index;
				V_VT(&index) = VT_INT;
				index.intVal = i;
				_bstr_t attr_name = MSHTML::IHTMLDOMAttributePtr(col->item(&index))->nodeName;
				_bstr_t attr_value;
				wchar_t *real_attr_name = 0;
				const wchar_t *prefix = L"unknown_attribute_";
				if (wcsncmp(attr_name, prefix, wcslen(prefix)))
				{
					continue;
				}
				else
				{
					real_attr_name = (wchar_t *)attr_name + wcslen(prefix);
					attr_value = MSHTML::IHTMLDOMAttributePtr(col->item(&index))->nodeValue;
				}
				MSXML2::IXMLDOMAttributePtr attr(doc->createNode(NODE_ATTRIBUTE, real_attr_name, U::FBNS));
				attr->appendChild(doc->createTextNode(attr_value));
				xinl->setAttributeNode(attr);
			}
		}

		// Process children recursively
		MSHTML::IHTMLDOMNodePtr cn(inl->firstChild);
		while (cn)
		{
			if (cn->nodeType == NODE_TEXT)
				MSXML2::IXMLDOMNodePtr(xinl)->appendChild(doc->createTextNode(cn->nodeValue.bstrVal));
			else if (cn->nodeType == NODE_ELEMENT && xname != L"image") // added by SeNS
				MSXML2::IXMLDOMNodePtr(xinl)->appendChild(ProcessInline(cn, doc));
			cn = cn->nextSibling;
		}

		return xinl;
	}

	///<summary>Build a XML-analog of P HTML-element with its subelements</summary>
	/// <params name="p">Processed P HTML-element</params>
	/// <params name="doc">XML-document</params>
	/// <params name="indent">Indent size</params>
	///<returns>Builded IXMLDOMNodePtr with all internal nodes<returns>
	static MSXML2::IXMLDOMNodePtr ProcessP(MSHTML::IHTMLElementPtr p, MSXML2::IXMLDOMDocument2Ptr doc, const wchar_t *baseName)
	{
		MSHTML::IHTMLDOMNodePtr hp(p);
        CString cls = p->className;
        CString xname = baseName;
		CString test = p->innerHTML;

		//define XML-element name
		if (cls == L"text-author")
			xname = L"text-author";
		else if (cls == L"subtitle")
			xname = L"subtitle";
		else if (cls == L"th")
			xname = L"th";
		else if (cls == L"td")
			xname = L"td";

		// process empty paragraph
		if (cls != L"td" && cls != L"th")
		{
			// no child or has just one empty text child
			if (!hp->hasChildNodes() ||
				(!hp->firstChild->nextSibling && hp->firstChild->nodeType == NODE_TEXT && U::is_whitespace(hp->firstChild->nodeValue.bstrVal)))
			{
				MSHTML::IHTMLElementPtr pParent = MSHTML::IHTMLElementPtr(p)->parentElement;
					// for stanza allow element with space symbol
					if (pParent && U::scmp(pParent->className, L"stanza") == 0) // if(U::scmp(baseName, L"v") == 0)
					{
						MSXML2::IXMLDOMNodePtr vNode = doc->createNode(NODE_ELEMENT, L"v", U::FBNS);
						vNode->appendChild(doc->createTextNode(L" "));
						return vNode;
					}
					else
						// create <empty-line> element
						return doc->createNode(NODE_ELEMENT, L"empty-line", U::FBNS);
			}
		}

		MSXML2::IXMLDOMElementPtr xp(doc->createNode(NODE_ELEMENT, xname.GetString(), U::FBNS));

		// setting attribute for table elements
		if (cls == L"td" || cls == L"th")
		{
			SetAttr2(xp, L"id", U::FBNS, p->id, doc);
			SetAttr2(xp, L"style", U::FBNS, U::GetAttrB(p, L"fbstyle"), doc);
			SetAttr2(xp, L"colspan", U::FBNS, U::GetAttrB(p, L"fbcolspan"), doc);
			SetAttr2(xp, L"rowspan", U::FBNS, U::GetAttrB(p, L"fbrowspan"), doc);
			SetAttr2(xp, L"align", U::FBNS, U::GetAttrB(p, L"fbalign"), doc);
			SetAttr2(xp, L"valign", U::FBNS, U::GetAttrB(p, L"fbvalign"), doc);
		}
		// process all children of <p>: strong, images etc
		hp = hp->firstChild;
		while (hp)
		{
			if (hp->nodeType == NODE_TEXT) // text segment
				MSXML2::IXMLDOMNodePtr(xp)->appendChild(doc->createTextNode(hp->nodeValue.bstrVal));
			else if (hp->nodeType == NODE_ELEMENT)
				MSXML2::IXMLDOMNodePtr(xp)->appendChild(ProcessInline(hp, doc));
			hp = hp->nextSibling;
		}

		return xp;
	}

	///<summary>Build a div element with subelements</summary>
	/// <params name="div">Processed DIV element</params>
	/// <params name="doc">XML-document</params>
	/// <params name="indent">Indent size</params>
	///<returns>Builded IXMLDOMNodePtr with internal nodes</returns>
	static MSXML2::IXMLDOMNodePtr ProcessDiv(MSHTML::IHTMLElementPtr div, MSXML2::IXMLDOMDocument2Ptr doc, int indent)
	{
        CString cls = div->className; //DIV class
        
		MSXML2::IXMLDOMElementPtr xdiv(doc->createNode(NODE_ELEMENT, cls.GetString(), U::FBNS));
		if (div->id.length()>0)
			SetAttr2(xdiv, L"id", U::FBNS, div->id, doc);

		// process image
		if (cls == L"image")
		{
			SetAttr2(xdiv, L"l:href", U::XLINKNS, U::GetAttrB(div, L"href"), doc);
			SetAttr2(xdiv, L"title", U::FBNS, U::GetAttrB(div, L"title"), doc);
			return xdiv;
		}

		// process table styles
		if (cls == L"table")
			SetAttr2(xdiv, L"style", U::FBNS, U::GetAttrB(div, L"fbstyle"), doc);
		if (cls == L"tr")
			SetAttr2(xdiv, L"align", U::FBNS, U::GetAttrB(div, L"fbalign"), doc);

		const wchar_t *bn = cls == L"stanza" ? L"v" : L"p";

		// do for all children
		MSHTML::IHTMLDOMNodePtr fc = MSHTML::IHTMLDOMNodePtr(div)->firstChild;
		while (fc)
		{
			CString name = fc->nodeName;

			// process div recursively
			if (name == L"DIV")
			{
				Indent(MSXML2::IXMLDOMNodePtr(xdiv), doc, indent + 1);
				MSXML2::IXMLDOMNodePtr(xdiv)->appendChild(ProcessDiv(MSHTML::IHTMLElementPtr(fc), doc, indent + 1));
			}
			// process p
			else if (name == L"P")
			{
                Indent(MSXML2::IXMLDOMNodePtr(xdiv), doc, indent + 1);
                MSXML2::IXMLDOMNodePtr(xdiv)->appendChild(ProcessP(MSHTML::IHTMLElementPtr(fc), doc, bn));
			}
			fc = fc->nextSibling;
		}
		// indent top node
		Indent(MSXML2::IXMLDOMNodePtr(xdiv), doc, indent);

		return xdiv;
	}

	///<summary>Find all BODY and process it adding to XML-root</summary>
	/// <params name="body">BODY element (fbw_body)</params>
	/// <params name="doc">XML-document</params>
	///<returns>IXMLDOMNodePtr with all added nodes or empty IXMLDOMNodePtr<returns>
	static void GetBodies(MSHTML::IHTMLElementPtr body, MSXML2::IXMLDOMDocument2Ptr doc)
	{
        MSHTML::IHTMLDOMChildrenCollectionPtr bodies = MSHTML::IElementSelectorPtr(body)->querySelectorAll(L"div.body");
		for (int i = 0; i < bodies->length; ++i)
		{
			MSHTML::IHTMLElementPtr div(bodies->item(i));

			// Create xml-element with "fbname" attribute, indent and add to root
			MSXML2::IXMLDOMElementPtr xb(ProcessDiv(div, doc, INDENT_SIZE));
			SetAttr2(xb, L"name", U::FBNS, U::GetAttrB(div, L"fbname"), doc);
			Indent(MSXML2::IXMLDOMNodePtr(doc->documentElement), doc, INDENT_SIZE);
			MSXML2::IXMLDOMNodePtr(doc->documentElement)->appendChild(MSXML2::IXMLDOMNodePtr(xb));
		}
	}

	///<summary>Find first DIV with particular class and process it</summary>
	/// <params name="body">BODY element</params>
	/// <params name="doc">XML-document</params>
	/// <params name="name">class name</params>
	/// <params name="indent">Indent size</params>
	///<returns>IXMLDOMNodePtr with all added nodes or empty IXMLDOMNodePtr<returns>
	static MSXML2::IXMLDOMNodePtr GetDiv(MSHTML::IHTMLElementPtr body, MSXML2::IXMLDOMDocument2Ptr doc, const wchar_t *name, int indent)
	{
        MSHTML::IHTMLElementPtr div = MSHTML::IElementSelectorPtr(body)->querySelector(L"div." + _bstr_t(name));
        if(div)
			return ProcessDiv(div, doc, indent);
        else
			return nullptr;
	}
	///<summary>Get opening& closing xml tags</summary>
///<param name="xml">String for searching</param>
///<param name="open_tag_begin">Pointer to begin of opening tag (out)</param>
///<param name="open_tag_end">Pointer to end of opening tag (out)</param>
///<param name="close_tag_begin">Pointer to begin of closing tag (out)</param>
///<param name="close_tag_end">Pointer to end of closing tag (out)</param>
///<param name="xml">String for searching</param>
///<returns>true-if tag found, false -if not found</returns>
	static bool tGetFirstXMLNodeParams(const wchar_t* xml, wchar_t** open_tag_begin, wchar_t** open_tag_end, wchar_t** close_tag_begin, wchar_t** close_tag_end)
	{
		if (!xml)
			return false;

		// serach first xml opening tag

		//search opening bracket
		wchar_t* optb = wcsstr(const_cast<wchar_t*>(xml), L"<");
		//skip closing tags & xml-header
		while (optb && (*(optb + 1) == L'/' || *(optb + 1) == L'?'))
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

		// We know now name. So will search closing tag
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
	static bool tFindCloseTag(wchar_t* start, wchar_t* tag_name, wchar_t** close_tag)
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

//---------------------- DomPath--------------------------------------------
// Class to construct HTML-XML DOM path to transfer selection between BODY<->SOURCE views
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
	int BuildPathForXMLElement(MSXML2::IXMLDOMNodePtr xml_root);
	int BuildPathFromHTMLDOM(MSHTML::IHTMLElementPtr html_root, MSHTML::IHTMLDOMNodePtr selNode);

	MSXML2::IXMLDOMNodePtr GetNodeFromXMLDOM(MSXML2::IXMLDOMNodePtr root);
	MSHTML::IHTMLDOMNodePtr GetNodeFromHTMLDOM(MSHTML::IHTMLDOMNodePtr root);
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
		//MSHTML::IHTMLDOMNodePtr parentNode = endNode->parentNode;
		MSHTML::IHTMLDOMNodePtr currentNode = endNode;

        if(currentNode->nodeType == NODE_TEXT)
			currentNode = currentNode->parentNode;
            
//		if (U::scmp(currentNode->nodeName, L"#text") == 0)
//			currentNode = currentNode->previousSibling;
		
        // climbing up
		while (currentNode)
		{
			CString ss = currentNode->nodeName;

			if (currentNode == root)
				return true;
			
			MSHTML::IHTMLDOMNodePtr prev = currentNode->previousSibling;
			// count siblings
            while (prev)
			{
                if(prev->nodeType != NODE_TEXT)
                    ++id;
//				if (U::scmp(previousSibling->nodeName, L"#text") != 0)
//					++id;
				prev = prev->previousSibling;
			}
			m_path.push_front(id);
			id = 0;
			currentNode = currentNode->parentNode;
		}
        // not found
		m_path.clear();
		return false;
	}

    ///<summary>Fill path queue and define root path element</summary>
///<param name="root">Topmost HTML node</param>
///<param name="endNode">Start HTML node</param>
///<returns>true-if success, false - if error</returns>
	int DomPath::BuildPathFromHTMLDOM(MSHTML::IHTMLElementPtr html_root, MSHTML::IHTMLDOMNodePtr selNode)
	{
        const int OTHER_ROOT = -3;
        const int ANNOTATION_ROOT = -2;
        const int HISTORY_ROOT = -1;
        int body_index = OTHER_ROOT;

        // define elements for contains checking
        MSHTML::IHTMLElementPtr cur_element = nullptr;
        if (selNode->nodeType == NODE_TEXT)
            cur_element = selNode->parentNode;
        else
            cur_element = selNode;
        
        MSHTML::IHTMLElementPtr ann = MSHTML::IElementSelectorPtr(html_root)->querySelector(L"div#fbw_body>div.annotation");
        MSHTML::IHTMLElementPtr hist = MSHTML::IElementSelectorPtr(html_root)->querySelector(L"div#fbw_body>div.history");

        MSHTML::IHTMLElementPtr root = nullptr;
        if (ann->contains(cur_element))
        {
            body_index = ANNOTATION_ROOT;
            root = ann;
        }
        else if (hist->contains(cur_element))
        {
            body_index = HISTORY_ROOT;
            root = hist;
        }
        else
        {
            MSHTML::IHTMLDOMChildrenCollectionPtr bodies = MSHTML::IElementSelectorPtr(html_root)->querySelectorAll(L"div#fbw_body>div.body");
            for (int i = 0; i < bodies->length; ++i)
            {
                // define body number for start element
                if (MSHTML::IHTMLElementPtr(bodies->item(i))->contains(cur_element))
                {
                    body_index = i;
                    root = bodies->item(i);
                    break;
                }
            }
        }
        if(CreatePathFromHTMLDOM(root, cur_element))
            return body_index;
        else
            return OTHER_ROOT;
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

		// climb down, skipping text nodes
		for (unsigned int i = 0; i < m_path.size(); ++i) //skip root
		{
			currentNode = currentNode->firstChild;
			if (!currentNode)
                return nullptr;

			if (U::scmp(currentNode->nodeName, L"#text") == 0)
			{
				currentNode = currentNode->nextSibling;
				if (!currentNode)
					return nullptr;
			}

			// move in siblings, skipping text nodes
			for (int j = 0; j < m_path[i]; ++j)
			{
				currentNode = currentNode->nextSibling;
				if (U::scmp(currentNode->nodeName, L"#text") == 0)
					currentNode = currentNode->nextSibling;

				if (!currentNode)
					return nullptr;
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
			//m_path.push_front(0); // add root to path
			return true;
		}
		int id = 0;

		// climb up from target node until root
		MSXML2::IXMLDOMNodePtr currentNode = endNode;

		if (currentNode->nodeType == NODE_TEXT)
			currentNode = currentNode->previousSibling;

		while (currentNode)
		{
			if (currentNode == root)
			{
				//m_path.push_front(0); // add root to path
				return true;
			}

			//calculate order in siblings and store in path
			MSXML2::IXMLDOMNodePtr prev = currentNode->previousSibling;
			while (prev)
			{
				if (prev->nodeType != NODE_TEXT)
					++id;
				prev = prev->previousSibling;
			}
			m_path.push_front(id);
			
			id = 0;
			currentNode = currentNode->parentNode;
		}
		// clear path on error
		m_path.clear();
		return false;
	}
	
///<summary>Fill path queue and define root path element</summary>
///<param name="xml">XML document</param>
///<param name="selectedElement">XML-element which path is builded</param>
///<returns>index of root path element (dody, annotation, etc.)</returns>
	int DomPath::BuildPathForXMLElement(MSXML2::IXMLDOMNodePtr xml_root)
	{
		const int OTHER_ROOT = -3;
		const int ANNOTATION_ROOT = -2;
		const int HISTORY_ROOT = -1;

        MSXML2::IXMLDOMElementPtr selectedElement = GetNodeFromXMLDOM(xml_root);
        
        if(!selectedElement)
            return OTHER_ROOT;

		// define root element to build path

		int body_index = OTHER_ROOT; // index of root element with selected begin element

		MSXML2::IXMLDOMNodePtr root_element = nullptr;
		MSXML2::IXMLDOMNodePtr descr = xml_root->selectSingleNode(L"fb:description");
		MSXML2::IXMLDOMNodePtr ann = descr->selectSingleNode(L"fb:title-info/fb:annotation");
		MSXML2::IXMLDOMNodePtr hist = descr->selectSingleNode(L"fb:document-info/fb:history");

		if (ann && U::IsParentElement(selectedElement, ann))
		{
			// annotation
			root_element = ann;
			body_index = ANNOTATION_ROOT;
		}
		else if (hist && U::IsParentElement(selectedElement, hist))
		{
			// history
			root_element = hist;
			body_index = HISTORY_ROOT;
		}
		else if (descr && U::IsParentElement(selectedElement, descr))
		{
			// somewhere in description
			body_index = OTHER_ROOT;
		}
		else
		{
			// define body index
			MSXML2::IXMLDOMNodeListPtr pXMLDomNodeList = xml_root->selectNodes("//fb:body");
			for (int i = 0; i < pXMLDomNodeList->length; i++)
			{
				if (U::IsParentElement(selectedElement, pXMLDomNodeList->item[i]))
				{
					// found right body
					root_element = pXMLDomNodeList->item[i];
					body_index = i;
					break;
				}
			}
		}
        // create path
		if(CreatePathFromXMLDOM(root_element, selectedElement))
			return body_index;
		else
			return OTHER_ROOT;
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

///<summary>Recursively build path queue</summary>
///<param name="xml">String for searching</param>
///<param name="open_tag_begin">Select position in string</param>
///<param name="char_pos">Position (out)</param>
///<returns>true-if success, false - if xml=NULL</returns>
	bool DomPath::CPFT(const wchar_t* xml, int pos, int* char_pos)
	{
		if (!xml)
			return false;

		const wchar_t* selpos = xml + pos;
		int virtual_pos = pos;
        
        // если попали в Indent (пустой тег) надо сместиться назад до первого непустого символа
        // если не стоим на пробеле - все Ok
        if(iswspace(*selpos))
        {
            //идем вперед до первого <, пока встречаются пустые символы (если встречаем непустой - все Ok)
            const wchar_t* e_curpos = selpos;
            while (iswspace(*e_curpos) && *e_curpos != L'\0')
               e_curpos++;
            
            // идем назад до первого >, пока встречаются  пустые символы (если встречаем непустой - все Ok)
            const wchar_t* b_curpos = selpos;
            while (iswspace(*b_curpos) && b_curpos != xml)
               b_curpos--;

            // становимся на предыдущий перед > символ
            if(*e_curpos == L'<' && *b_curpos == L'>')
            {
                selpos = b_curpos-1;
            }
        }
		// TO_DO Надо сдвинуться до первого непустого символа в текстовом узле

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
			if (selpos >= open_tag_begin && selpos < open_tag_end)
			{
				*char_pos = 0;
				m_path.push_back(id);
				return true;
			}
			if (selpos > close_tag_begin && selpos <= close_tag_end)
			{
				*char_pos = close_tag_begin - open_tag_end;
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
		if (this->CPFT(xml, pos, char_pos))
		{
			if(m_path.size()>1)
				m_path.pop_front(); // remove root
			return true;
		}
		else
			return false;
	}

///<summary>Get absolute position in XML-string
/// using prebuilded path</summary>
///<param name="xml">XML-string</param>
///<param name="char_pos">Offset in text content of node</param>
///<returns>Absolute positionin xml-string</returns>
	int DomPath::GetNodeFromText(wchar_t* xml, int char_pos)
	{
		m_path.push_front(0); // add root to path

		// define target node from prebuilded path
		wchar_t* curpos = nullptr;
		wchar_t* nested_node = xml;
		for (unsigned int i = 0; i < m_path.size(); ++i) 
		{
			curpos = nested_node;
			int node_id = m_path[i];

			for (int j = 0; j < node_id; ++j)
			{
				if (!tGetFirstXMLNodeParams(curpos, nullptr, nullptr, nullptr, &curpos))
					return 0;
			}
			if (!tGetFirstXMLNodeParams(curpos, nullptr, &nested_node, nullptr, &curpos))
				return 0;
		}

		// calculate position in node text skipping nested nodes
		wchar_t* open_tag_begin = nullptr;    // start position of opening tag
		wchar_t* open_tag_end = nullptr;      // end position of opening tag
		wchar_t* close_tag_begin = nullptr;   // start position of closing tag
		wchar_t* close_tag_end = nullptr;     // end position of closing tag
		int virtualcount = 0;

		int realcount = char_pos;
		curpos = nested_node; // start position of node text content
		while (tGetFirstXMLNodeParams(curpos, &open_tag_begin, &open_tag_end, &close_tag_begin, &close_tag_end))
		{
			virtualcount += open_tag_begin - curpos;
			if (virtualcount > char_pos)
				return nested_node - xml + realcount;

			realcount += close_tag_end - open_tag_begin; // add length of nested node
			curpos = close_tag_end;
		}

		return nested_node - xml + realcount;
	}
    
    //----------------------File functions-----------------------------------------
	/// <summary>Load fb2-file into buffer.
	/// Buffer allocate inside of function! Please, don't forget to deallocate</summary>
	/// <param name="filename">fb2 filename</param>
	/// <param name="buffer">buffer ptr</param>
	/// <returns>buffer size-success, 0 - error</returns>
	static size_t ReadPlainTextFile(const CString &filename, char **buffer)
	{
		size_t filesize;

		FILE *f = _wfopen(filename, L"rbS");

		// Failed to open file
		if (f == NULL)
			throw(L"Can't open file " + filename);

		struct _stat fileinfo;
		_wstat(filename, &fileinfo);
		filesize = fileinfo.st_size;

		// Read entire file contents in to memory
		*buffer = (char *)malloc(sizeof(char) * filesize + 2);
		fread(*buffer, sizeof(char), filesize, f);

		// set end of line for unicode strings
		(*buffer)[filesize] = '\0';
		(*buffer)[filesize + 1] = '\0';

		fclose(f);
		return filesize + 2;
	}

	/// <summary>Prepare ZIP-error message</summary>
	/// <returns>Error string</returns>
	static CString HandleZIPLibError(zip_error_t &error)
	{
		CString ziperror(zip_error_strerror(&error));
		zip_error_fini(&error);
		return ziperror;
	}

	/// <summary>Decrypt zip file and load into buffer</summary>
	/// <summary>Buffer allocate inside of function! Please, don't forget to deallocate</summary>
	/// <param name="filename">zip filename</param>
	/// <param name="buffer">buffer ptr</param>
	/// <returns>buffer size-success, throw exeptions</returns>
	static size_t ReadZIPFile(const CString &filename, char **buffer)
	{
		// open zip and read to buffer
		size_t filesize = 0;
		zip_source_t *src;
		zip_t *z;
		zip_error_t error;

		zip_error_init(&error);
		// create source from name
		if ((src = zip_source_win32w_create(filename, 0, -1, &error)) == NULL)
		{
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		// open zip archive from source
		if ((z = zip_open_from_source(src, 0, &error)) == NULL)
		{
			zip_source_free(src);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		// Check number of files in archive
		if (zip_get_num_files(z) > 1)
		{
			zip_close(z);
			throw(L"There is more than one file in zip archive.\nFile:" + filename);
		}

		struct zip_stat file_info; // file info
		zip_stat_index(z, 0, 0, &file_info);
		filesize = size_t(file_info.size);
		*buffer = (char *)malloc(sizeof(char) * filesize + 2);

		struct zip_file *file_in_zip; // file desciptor in archive
		file_in_zip = zip_fopen_index(z, 0, 0);
		if (!file_in_zip)
		{
			zip_close(z);
			free(buffer);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		// read zip-file
		if (zip_fread(file_in_zip, *buffer, sizeof(char) * file_info.size + 2) == -1)
		{
			zip_close(z);
			free(buffer);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}
		zip_close(z);
		// set end of line for unicode strings
		(*buffer)[filesize] = '\0';
		(*buffer)[filesize + 1] = '\0';

		return filesize + 2;
	}

	///<summary> Save buffer as zip-file</summary>
	///filename - zip-archive filename. Filename for internal file - zip-archive filename +."fb2"
	/// buffer - buffer with data
	/// bufsize - buffer size
	/// return true on SUCCESS
	static bool SaveZIPFile(const CString &filename, void *buffer, size_t bufsize)
	{
		zip_source_t *src;
		zip_t *z;
		zip_error_t error;
		zip_source *source;

		// define name for internal file& convert to UTF-8
		CString ff = filename.Left(filename.GetLength() - 4);
		if (ff.Right(4) != L".fb2")
			ff += L".fb2";
		wchar_t *filename_noext = PathFindFileName(ff);

		DWORD ulen = ::WideCharToMultiByte(CP_UTF8, 0, filename_noext, wcslen(filename_noext), NULL, 0, NULL, NULL);
		std::string FB2filename(ulen, 0);
		WideCharToMultiByte(CP_UTF8, 0, filename_noext, wcslen(filename_noext), &FB2filename[0], ulen, NULL, NULL);

		zip_error_init(&error);

		// create source from name
		if ((src = zip_source_win32w_create(filename, 0, -1, &error)) == NULL)
		{
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		// open zip archive from source
		if ((z = zip_open_from_source(src, ZIP_CREATE, &error)) == NULL)
		{
			zip_source_free(src);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		// fill zip-buffer
		if ((source = zip_source_buffer(z, buffer, bufsize, 0)) == NULL)
		{
			zip_source_free(src);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		//create internal file
		int index;
		if ((index = (int)zip_file_add(z, FB2filename.c_str(), source, ZIP_FL_OVERWRITE)) == -1)
		{
			zip_source_free(src);
			throw(HandleZIPLibError(error) + L"\nFile:" + filename);
		}

		zip_close(z);

		// Trick! Update datetime stamp, as libzip doesn't do it for update operation
		SYSTEMTIME st;
		FILETIME ft;
		FILE *f = _wfopen(filename, L"r+");

		SystemTimeToFileTime(&st, &ft);
		SetFileTime(f, &ft, NULL, NULL);
		fclose(f);

		return true;
	}

	//----------------------Doc class----------------------------------------------
	Doc::Doc(HWND hWndFrame) : 
        m_filename(L"Untitled.fb2"), 
        m_namevalid(false),
		m_body(hWndFrame, true),
		m_doc_changed(false),
		//m_trace_changes(true),
        m_saved_xml(nullptr),
		m_file_encoding(L"utf-8")
	{
	}

	Doc::~Doc()
	{
		// destroy windows explicitly
		if (m_body.IsWindow())
			m_body.DestroyWindow();
       	if (m_saved_xml)
            m_saved_xml.Release();
	}

	/// <summary>Get image data as base64 string</summary>
	/// <param name="id">image id</param>
	/// <param name="vt">Image data (OUT)</param>
	/// <returns>true if SUCCESS</returns>
    bool Doc::GetBinary(const wchar_t *id, _variant_t &vt)
	{
		if (id && *id == L'#')
		{
			try
			{
				CComVariant params[1];
				params[0] = id + 1; // miss started ":" character
				CComVariant res;
				InvokeFunc(L"GetImageData", params, 1, res);
				vt = res;
				return true;
			}
			catch (_com_error(&e))
			{
				U::ReportError(e);
			}
		}
		return false;
	}

	/// <summary>Invoke function from js script</summary>
	/// <param name="FuncName">Function name</param>
	/// <param name="params">Function parameters</param>
	/// <param name="count">Parameters count</param>
	/// <param name="vtResult">Function result</param>
	HRESULT Doc::InvokeFunc(BSTR FuncName, CComVariant *params, int count, CComVariant &vtResult)
	{
		LPDISPATCH pScript;
		IHTMLDocument2Ptr doc = m_body.Browser()->Document;
        
		doc->get_Script(&pScript);
		pScript->AddRef();

		BSTR szMember = FuncName;
		DISPID dispid;

		HRESULT hr = pScript->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
        //SysFreeString(szMember);
        
		if (SUCCEEDED(hr))
		{
			CComDispatchDriver dispDriver(pScript);
			hr = dispDriver.InvokeN(dispid, params, count, &vtResult);
		}

		return hr;
	}

	/// <summary>Add js script body from file into html</summary>
	/// <param name="filePath">Path to file</param>
	bool Doc::LoadScript(BSTR filePath)
	{
		int offset = 0;
		char *buffer = nullptr;
		size_t buffersize = 0;
		BSTR ustr = nullptr;

		buffersize = ReadPlainTextFile(filePath, &buffer);
		if (buffersize == 0 && !buffer)
		{
			free(buffer);
			return false;
		}

		// convert to widechar
		if (buffer[0] == '\xEF' && buffer[1] == '\xBB' && buffer[2] == '\xBF')
		{
			offset = 3;
			DWORD ulen = ::MultiByteToWideChar(CP_UTF8, 0, buffer + offset, buffersize - offset - 1, NULL, 0);
			ustr = ::SysAllocStringLen(NULL, ulen);
			::MultiByteToWideChar(CP_UTF8, 0, buffer + offset, buffersize - offset - 1, ustr, ulen);
		}
		else if ((buffer[0] == '\xFF' && buffer[1] == '\xFE') || (buffer[0] == '\xFF' && buffer[1] == '\xFE'))
		{
			offset = 2;
			ustr = ::SysAllocStringLen(NULL, buffersize - offset);
			ustr = (wchar_t *)(buffer + offset);
		}
		else
		{
			DWORD ulen = ::MultiByteToWideChar(CP_ACP, 0, buffer, buffersize - 1, NULL, 0);
			ustr = ::SysAllocStringLen(NULL, ulen);
			::MultiByteToWideChar(CP_ACP, 0, buffer, buffersize - 1, ustr, ulen);
		}
		free(buffer);

		// add script text

		// remove previous script element
		MSHTML::IHTMLDocument3Ptr doc = m_body.Browser()->Document;
		MSHTML::IHTMLElementPtr Script = doc->getElementById("userCmd");
		if (Script)
			MSHTML::IHTMLDOMNodePtr(Script)->removeNode(VARIANT_TRUE);

		// add new script element
		MSHTML::IHTMLElementPtr headElement = doc->getElementsByTagName("head")->item(0);
		MSHTML::IHTMLScriptElementPtr scriptElement = MSHTML::IHTMLDocument2Ptr(doc)->createElement("script");
		MSHTML::IHTMLElementPtr(scriptElement)->id = "userCmd";
		scriptElement->put_text(ustr);
		MSHTML::IHTMLDOMNodePtr(headElement)->appendChild(MSHTML::IHTMLDOMNodePtr(scriptElement));

		::SysFreeString(ustr);
		return true;
	}
	/// <summary>Run js script from file</summary>
	/// <param name="filePath">Path to file</param>
	void Doc::RunScript(BSTR filePath)
	{
		if (LoadScript(filePath))
		{
			CComVariant vtResult;
			CComVariant params(filePath);
			InvokeFunc(L"Run", &params, 0, vtResult);
		}
	}

	///<summary>Get current doc filename</summary>
	///<returns>file name<returns>
	CString Doc::GetDocFileName() const
	{
		if (m_namevalid)
			return m_filename;
		else
			return L"";
	}

	/// <summary>Load file to Body Editor</summary>
	/// <param name="hWndParent">Main editor window</param>
	/// <param name="filename">filename</param>
	/// <returns>true-success, false - error</returns>
	bool Doc::Load(HWND hWndParent, const CString &filename)
	{

		char *buffer = nullptr;
		size_t buffersize = 0;
		U::CPersistentWaitCursor wc;

		try
		{
			if (U::GetFileExtension(filename) == L"zip")
			{
				buffersize = ReadZIPFile(filename, &buffer);
			}
			else
			{
				buffersize = ReadPlainTextFile(filename, &buffer);
			}

			if (buffersize == 0)
			{
				free(buffer);
				return false;
			}
			CString enc = U::GetStringEncoding(buffer);

			CComVariant params[3];
			params[2] = filename;
			params[0] = _Settings.GetInterfaceLanguageName();
			CComVariant res;

			//Encode xml-string
			int offset = 0;
			BSTR ustr;

			if (enc == L"utf-8" || enc == L"")
			{
				if (buffer[0] == '\xEF' && buffer[1] == '\xBB' && buffer[2] == '\xBF')
				{
					offset = 3;
				}
				DWORD ulen = ::MultiByteToWideChar(CP_UTF8, 0, buffer + offset, buffersize - offset - 1, NULL, 0);
				ustr = ::SysAllocStringLen(NULL, ulen);
				::MultiByteToWideChar(CP_UTF8, 0, buffer + offset, buffersize - offset - 1, ustr, ulen);
				params[1] = ustr;
			}
			else if (enc == L"utf-16")
			{
				if ((buffer[0] == '\xFF' && buffer[1] == '\xFE') || (buffer[0] == '\xFF' && buffer[1] == '\xFE'))
				{
					offset = 2;
				}
				ustr = ::SysAllocStringLen(NULL, buffersize - offset);
				ustr = (wchar_t *)(buffer + offset);
				params[1] = ustr;
			}
			else
			{
				params[1] = buffer;
			}

			// Create ActiveX object = WebBrowser and load main.html to call jscript function
			CRect rr(0, 0, 500, 500);
			m_body.Create(hWndParent, &rr, _T("{8856F961-340A-11D0-A96B-00C04FD705A2}"));

			HRESULT hr; 
			hr = m_body.Browser()->Navigate(U::GetProgDirFile(L"main.html").GetString());
            
            // wait for complete load page
			MSG msg;
			while (!m_body.Loaded() && ::GetMessage(&msg, NULL, 0, 0))
			{
					::TranslateMessage(&msg);
					::DispatchMessage(&msg);
			}
			m_body.SetExternalDispatch(m_body.CreateHelper());
            
			// Load fb2-file to body wih XSLT transform
			InvokeFunc(L"apiLoadFB2_new", params, 3, res);

			if (buffer)
				free(buffer);

			// process return value
			if (res.vt == VT_BOOL && res.vt == false)
			{
				return false;
			}

			if (res.vt == VT_BSTR)
			{
                U::ReportError(res.bstrVal);
				return false;
			}
			m_file_encoding = enc;
			m_body.Init();
			m_body.ApplyConfChanges();
			// mark unchanged
			Changed(false);
            m_namevalid = true;

			if (filename != L"blank.fb2")
				m_filename = filename;
            else 
                m_filename = L"Untitled.fb2";
		}
		catch (_com_error &e)
		{
			if (buffer)
				free(buffer);
			U::ReportError(e);
			return false;
		}
		catch (CString &e)
		{
			if (buffer)
				free(buffer);
            U::ReportError(e);
			return false;
		}

		return true;
	}

	/// <summary>Load blank fb2-file "blank.fb2" to Body Editor</summary>
	/// <param name="hWndParent">Main editor window</param>
	void Doc::CreateBlank(HWND hWndParent)
	{
		Load(hWndParent, L"blank.fb2");
        m_namevalid = false;
	}

	/// <summary>Create xml document from html</summary>
	/// <param name="encoding">HTML encoding</param>
	/// <returns>xml document</returns>
	void Doc::CreateDOM()
	{
		CString encoding = _Settings.m_keep_encoding ? m_file_encoding : _Settings.GetDefaultEncoding();
		// clean classes setted by TreeCmd scripts
		_EDMnr.CleanUpAll(); //??

		// normalize body first
		//m_body.TraceChanges(false);
		m_body.Normalize();
		try
		{
            if(m_saved_xml)
                m_saved_xml.Release();
			// create xml-document
            
			m_saved_xml = U::CreateDocument(false);
			m_saved_xml->async = VARIANT_FALSE;

			// set encoding
			if (!encoding.IsEmpty())
				MSXML2::IXMLDOMNodePtr(m_saved_xml)->appendChild(m_saved_xml->createProcessingInstruction(L"xml", (const wchar_t *)(L"version=\"1.0\" encoding=\"" + encoding + L"\"")));

			// create root element
			MSXML2::IXMLDOMElementPtr root = m_saved_xml->createNode(NODE_ELEMENT, L"FictionBook", U::FBNS);
			root->setAttribute(L"xmlns:l", U::XLINKNS);
			m_saved_xml->documentElement = MSXML2::IXMLDOMElementPtr(root);


			// detect body
            MSHTML::IHTMLElementPtr fbw_body = MSHTML::IHTMLDocument3Ptr(m_body.Document())->getElementById(L"fbw_body");
          
			// add annotation
			MSXML2::IXMLDOMNodePtr ann(GetDiv(fbw_body, m_saved_xml, L"annotation", INDENT_SIZE + 2));

			// add history
			MSXML2::IXMLDOMNodePtr hist(GetDiv(fbw_body, m_saved_xml, L"history", INDENT_SIZE + 2));

			// add description
			CComVariant args[3];
			CComVariant res;
			if (hist)
				args[0] = hist.GetInterfacePtr();
			if (ann)
				args[1] = ann.GetInterfacePtr();
			args[2] = m_saved_xml.GetInterfacePtr();
            InvokeFunc(L"GetDesc", &args[0], 3, res);

			// add body elements
			GetBodies(fbw_body, m_saved_xml);
			CString ss = MSXML2::IXMLDOMNodePtr(m_saved_xml)->xml;
			// add binaries
            InvokeFunc(L"GetBinaries", &args[2], 1, res);
                              
            // indent root
			Indent(MSXML2::IXMLDOMNodePtr(root), m_saved_xml, 0);
		}
		catch (_com_error &e)
		{
			U::ReportError(e);
		}
		//m_body.TraceChanges(true);
		return;
	}

	/// <summary>Transfer selection from XML-document to HTML-doc</summary>
	/// <param name="selectedPosBegin">XML-selection begin  (char offset)</param>
	/// <param name="selectedPosBegin">XML-selection end (char offset)</param>
	/// <param name="ustr">XML-string</param>
	/// <returns>true on success</returns>
	bool Doc::SetHTMLSelectionFromXML(int selectedPosBegin, int selectedPosEnd, BSTR ustr)
	{
		int begin_char = 0; // offset selection begin
		int end_char = 0;   // offset selection end
		CString dd = ustr;
		
		int start_pos = 0;
		while ((start_pos = dd.Find(L"</body>", start_pos)) !=-1)
		{
			end_char = start_pos++;
		}

		bool one_pos = selectedPosEnd == selectedPosBegin; //no selection, just carret

		//	Build DOMpath for selection
		DomPath path_begin;
		DomPath path_end;

		path_begin.CreatePathFromText(ustr, selectedPosBegin, &begin_char);

		if (one_pos)
		{
			path_end = path_begin;
			end_char = begin_char;
		}
		else
		{
			path_end.CreatePathFromText(ustr, selectedPosEnd, &end_char);
		}

		//Build XML DOM paths

		// define root element to build path
		const int OTHER_ROOT = -3;
		const int ANNOTATION_ROOT = -2;
		const int HISTORY_ROOT = -1;

		int body_index1 = OTHER_ROOT; // index of root element with selected begin element
		int body_index2 = OTHER_ROOT; // index of root element with selected end element

		body_index1 = path_begin.BuildPathForXMLElement(m_saved_xml->documentElement);
		body_index2 = one_pos ? body_index1 : path_end.BuildPathForXMLElement(m_saved_xml->documentElement);


		//	Find element in HTML using path
		MSHTML::IHTMLElementPtr selectedHTMLElementBegin = nullptr;
		MSHTML::IHTMLElementPtr selectedHTMLElementEnd = nullptr;;

		MSHTML::IHTMLElementPtr root1 = nullptr; //root for begin element
		MSHTML::IHTMLElementPtr root2 = nullptr; //root for end element
		if (body_index1 == ANNOTATION_ROOT)
			root1 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_body>div.annotation");
		else if (body_index1 == HISTORY_ROOT)
			root1 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_body>div.history");
		else if (body_index1 == OTHER_ROOT)
			root1 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_descr");
		else if (body_index1 >= 0)
			root1 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelectorAll(L"div#fbw_body>div.body")->item(body_index1);

		selectedHTMLElementBegin = path_begin.GetNodeFromHTMLDOM(root1);
		if (one_pos)
		{
			selectedHTMLElementEnd = selectedHTMLElementBegin;
			end_char = begin_char;
		}
		else
		{
			if (body_index2 == ANNOTATION_ROOT)
				root2 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_body>div.annotation");
			else if (body_index2 == HISTORY_ROOT)
				root2 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_body>div.history");
			else if (body_index2 == OTHER_ROOT)
				root2 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_descr");
			else if (body_index2 >= 0)
				root2 = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelectorAll(L"div#fbw_body>div.body")->item(body_index2);
			selectedHTMLElementEnd = path_end.GetNodeFromHTMLDOM(root2);
		}

		if (!selectedHTMLElementBegin && selectedHTMLElementEnd)
		{
			selectedHTMLElementBegin = selectedHTMLElementEnd;
			begin_char = end_char;
		}

		if (selectedHTMLElementBegin && !selectedHTMLElementEnd)
		{
			selectedHTMLElementEnd = selectedHTMLElementBegin;
			end_char = begin_char;
		}

		if (!selectedHTMLElementBegin && !selectedHTMLElementEnd)
		{
			selectedHTMLElementEnd = selectedHTMLElementBegin = MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelector(L"div#fbw_body>div.annotation");
			end_char = begin_char = 0;
		}
			   
		ClearHTMLSelection();
		m_body_sel = m_body.SetSelection(selectedHTMLElementBegin, selectedHTMLElementEnd, begin_char, end_char);
		// scroll to the center of view
		MSHTML::IHTMLRectPtr rect = MSHTML::IHTMLElement2Ptr(selectedHTMLElementBegin)->getBoundingClientRect();
		MSHTML::IHTMLWindow2Ptr window(MSHTML::IHTMLDocument2Ptr(m_body.Document())->parentWindow);
		if (rect && window)
		{
			if (rect->bottom - rect->top <= _Settings.GetViewHeight())
				window->scrollBy(0, (rect->top + rect->bottom - _Settings.GetViewHeight()) / 2);
			else
				window->scrollBy(0, rect->top);
		}
		m_body.Synched(true); // document is in sync with source
		return true;
	}
	
	/// <summary>Get selection from HTML-document to XML-doc</summary>
	/// <param name="selectedPosBegin">XML-selection begin  (char offset)(OUT)</param>
	/// <param name="selectedPosBegin">XML-selection end (char offset)(OUT)</param>
	/// <param name="ustr">XML-string(OUT)</param>
	/// <returns>true on success</returns>
	bool Doc::GetXMLSelectionFromHTML(int& selectedPosBegin, int& selectedPosEnd, _bstr_t& ustr)
	{
		const int OTHER_ROOT = -3;
		const int ANNOTATION_ROOT = -2;
		const int HISTORY_ROOT = -1;

		DomPath selection_begin_path;
		DomPath selection_end_path;

		int selection_begin_char = 0; // offset selection begin
		int selection_end_char = 0; // offset selection end
		bool one_element = false; // flag for same element (if end and begin elements equal)
		//bstr_t path;
		int body_index1 = 0;
		int body_index2 = 0;

		MSHTML::IHTMLDOMNodePtr selectedBeginElement = nullptr;
		MSHTML::IHTMLDOMNodePtr selectedEndElement = nullptr;

		// get start&end elements & offsets
			// get start&end elements from selection
		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_body.Document())->getSelection();
		if (sel && sel->rangeCount > 0)
		{
			MSHTML::IHTMLDOMRangePtr range = sel->getRangeAt(0);
			selectedBeginElement = range->startContainer;
			selection_begin_char = range->startOffset;

			selectedEndElement = range->endContainer;
			selection_end_char = range->endOffset;
		}
		// get start&end elements from stored selection info
		else if (m_body_sel)
		{
			selectedBeginElement = m_body_sel->startContainer;
			selection_begin_char = m_body_sel->startOffset;
			selectedEndElement = m_body_sel->endContainer;
			selection_end_char = m_body_sel->endOffset;
		}
		// no selection at all - go to beginning of first body
		else
		{
			selectedBeginElement = selectedEndElement = MSHTML::IHTMLDOMNodePtr(MSHTML::IElementSelectorPtr(m_body.Document()->body)->querySelectorAll(L"div#fbw_body>div.body")->item(0))->firstChild;
			selection_begin_char = selection_end_char = 0;
		}
		one_element = selectedBeginElement == selectedEndElement;

		body_index1 = selection_begin_path.BuildPathFromHTMLDOM(m_body.Document()->body, selectedBeginElement);
		if (one_element)
		{
			selection_end_path = selection_begin_path;
			body_index2 = body_index1;
		}
		else
			body_index2 = selection_end_path.BuildPathFromHTMLDOM(m_body.Document()->body, selectedEndElement);

		// if document was changed rebuild XML DOM
		if (!m_body.IsSynched())
		{
			CreateDOM();

			if (!m_saved_xml)
				return false;
		}

		//Build XML DOM paths
		CString body_path;
		MSXML2::IXMLDOMNodePtr root_element1 = nullptr;
		MSXML2::IXMLDOMNodePtr root_element2 = nullptr;
		MSXML2::IXMLDOMNodePtr xml_root = m_saved_xml->documentElement;

		if (body_index1 == ANNOTATION_ROOT)
		{
			body_path.SetString(L"fb:description/fb:title-info/fb:annotation");
		}
		else if (body_index1 == HISTORY_ROOT)
		{
			body_path.SetString(L"fb:description/fb:document-info/fb:history");
		}
		else
		{
			body_path.Format(L"//fb:body[%i]", body_index1 + 1);
		}
		root_element1 = xml_root->selectSingleNode(body_path.GetString());

		if (one_element)
		{
			root_element2 = root_element1;
		}
		else
		{
			if (body_index2 == ANNOTATION_ROOT)
			{
				body_path.SetString(L"fb:description/fb:title-info/fb:annotation");
			}
			else if (body_index2 == HISTORY_ROOT)
			{
				body_path.SetString(L"fb:description/fb:document-info/fb:history");
			}
			else
			{
				body_path.Format(L"//fb:body[%i]", body_index2 + 1);
			}
			root_element2 = xml_root->selectSingleNode(body_path.GetString());
		}

		//body_path.Format(L"//fb:body[%i]", body_index2 + 1);
		//MSXML2::IXMLDOMNodePtr xml_body2 = MSXML2::IXMLDOMNodePtr(m_saved_xml->documentElement)->selectSingleNode(body_path.GetString());

		if (root_element1 && root_element2)
		{
			MSXML2::IXMLDOMNodePtr xml_selected_begin = nullptr;
			MSXML2::IXMLDOMNodePtr xml_selected_end = nullptr;

			xml_selected_begin = selection_begin_path.GetNodeFromXMLDOM(root_element1);
			selection_begin_path.CreatePathFromXMLDOM(m_saved_xml->documentElement, xml_selected_begin);
			if (one_element)
			{
				selection_end_path = selection_begin_path;
			}
			else
			{
				xml_selected_end = selection_end_path.GetNodeFromXMLDOM(root_element2);
				selection_end_path.CreatePathFromXMLDOM(m_saved_xml->documentElement, xml_selected_end);
			}
		}

		// get XML as string
		ustr = MSXML2::IXMLDOMNodePtr(m_saved_xml)->xml;

		//define select range in string
		selectedPosBegin = selection_begin_path.GetNodeFromText(ustr, selection_begin_char);
		selectedPosEnd = selection_end_path.GetNodeFromText(ustr, selection_end_char);

		return true;
	}

	// Selection storing operations
	///<summary>Restore selection in BODY view from selection info</summary>
	void Doc::RestoreHTMLSelection()
	{
		if (m_body_sel)
		{
			MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_body.Document())->getSelection();
			sel->removeAllRanges();
			sel->addRange(m_body_sel);
		}
	}

	///<summary>Reset selection info</summary>
	void Doc::ClearHTMLSelection()
	{
		if (m_body_sel)
			m_body_sel.Detach();
	}

	///<summary>Save BODY view selection in selection info</summary>
	void Doc::SaveHTMLSelection()
	{
		ClearHTMLSelection();
		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_body.Document())->getSelection();
		if (sel && sel->rangeCount > 0)
			m_body_sel = sel->getRangeAt(0);
	}

	///<summary>Save to file/ Internal function</summary>
	/// filename - filename.
	/// return true on SUCCESS
	bool Doc::SaveToFile(const CString &filename)
	{
        // construct the document
        //MSXML2::IXMLDOMDocument2Ptr ndoc(CreateDOM());
        
		// replace all nbsp on standard non-breaking space
        if (_Settings.GetNBSPChar().Compare(NOBREAK_SPACE) != 0)
        {
            MSXML2::IXMLDOMNodePtr node = MSXML2::IXMLDOMNodePtr(m_saved_xml)->firstChild;
            // CString nbsp = _Settings.GetNBSPChar();
            while (node && node != m_saved_xml)
            {
                if (node->nodeType == NODE_TEXT)
                {
                    CString s = node->nodeValue;
                    int n = s.Replace(_Settings.GetNBSPChar(), NOBREAK_SPACE);
                    int k = s.Replace(L"<p>" + NOBREAK_SPACE + L"</p>", L"<empty-line/>");
                    if (n || k)
                    {
                        node->nodeValue = s.GetString();
                    }
                }
                if (node->firstChild)
                    node = node->firstChild;
                else
                {
                    while (node && node != m_saved_xml && node->nextSibling == NULL)
                        node = node->parentNode;
                    
                    if (node && node != m_saved_xml)
                        node = node->nextSibling;
                }
            }
        }
        
        // save file
        char *pszBuf = nullptr;
        FILE *f = nullptr;
        try
        {
            DWORD dwBufSize;
            ULONG ulBytesRead;
            STATSTG stats;
			HRESULT hr;

            IStreamPtr pStream(m_saved_xml);
            //hr = ndoc->QueryInterface(&pStream);
            hr = pStream->Stat(&stats, 0); // get stream status

            dwBufSize = stats.cbSize.LowPart;
            pszBuf = new char[dwBufSize + 1];
            pszBuf[dwBufSize] = '\0';
            hr = pStream->Read((void *)pszBuf, dwBufSize, &ulBytesRead);
            if (U::GetFileExtension(filename) == L"zip")
            {
                //Save ZIP
                SaveZIPFile(filename, (void *)pszBuf, dwBufSize);
            }
            else
            {
                //Save fb2
                f = _wfopen(filename, L"wbS");
                fwrite(pszBuf, sizeof(char), dwBufSize, f);
                fclose(f);
            }
        }

        catch (CString &e)
        {
            // Clean up.
            if (pszBuf)
            {
                delete pszBuf;
            }
            if (f)
                fclose(f);
            U::ReportError(e.GetString());
            //				AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, e.GetString(), (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_ERROR_ICON);
            return false;
        }

        // Clean up.
        if (pszBuf)
        {
            delete pszBuf;
        }
        if (f)
            fclose(f);
        return true;
	}

  	///<summary> Save file</summary>
	/// filename - filename.
	/// return true on SUCCESS
	bool Doc::Save(const CString &filename)
	{
		U::CPersistentWaitCursor wc;
        CString fn = m_filename;
        if(filename && filename.GetLength()>0)
            fn = filename;
        if (SaveToFile(fn))
		{
			Changed(false);
            m_namevalid = true;
			m_filename = filename;
			
			return true;
		}
		return false;
	}

	// create ID for image from filename
  	///<summary>Create image ID from filename</summary>
	/// filename - filename.
	/// return ID
	static CString PrepareDefaultId(const CString &filename)
	{
		CString _filename = U::Transliterate(filename);
		// prepare a default id
		int cp = _filename.ReverseFind(L'\\');
		if (cp < 0)
			cp = 0;
		else
			++cp;
		CString newid(L"");
		wchar_t *ncp = newid.GetBuffer(_filename.GetLength() - cp);
		int newlen = 0;

		while (cp < _filename.GetLength())
		{
			wchar_t c = _filename[cp];
			if (iswalnum(c) || c == L'_' || c == L'.')
				ncp[newlen++] = c;
			++cp;
		}
		newid.ReleaseBuffer(newlen);
		if (!newid.IsEmpty() && !(iswalpha(newid.GetAt(0)) || newid.GetAt(0) == L'_'))
			newid.Insert(0, L'_');
		return newid;
	}

	// XML SAX error handler class
	class SAXErrorHandler : public CComObjectRoot, public MSXML2::ISAXErrorHandler
	{
	public:
		CString m_msg;
		int m_line, m_col;

		SAXErrorHandler() : m_line(0), m_col(0) {}
        
        // Fill message, line, col members
		void SetMsg(MSXML2::ISAXLocator *loc, const wchar_t *msg, HRESULT hr)
		{
			if (!m_msg.IsEmpty())
				return;
			m_msg = msg;
            // remove namespaces
			CString ns;
			ns.Format(L"{%s}", (const TCHAR *)U::FBNS);
			m_msg.Replace(ns, L"");
			ns.Format(L"{%s}", (const TCHAR *)U::XLINKNS);
			m_msg.Replace(ns, L"xlink");
            
			m_line = loc->getLineNumber();
			m_col = loc->getColumnNumber();
		}

		BEGIN_COM_MAP(SAXErrorHandler)
			COM_INTERFACE_ENTRY(MSXML2::ISAXErrorHandler)
		END_COM_MAP()

		STDMETHOD(raw_error)
			(MSXML2::ISAXLocator *loc, wchar_t *msg, HRESULT hr)
		{
			SetMsg(loc, msg, hr);
			return E_FAIL;
		}
		STDMETHOD(raw_fatalError)
			(MSXML2::ISAXLocator *loc, wchar_t *msg, HRESULT hr)
		{
			SetMsg(loc, msg, hr);
			return E_FAIL;
		}
		STDMETHOD(raw_ignorableWarning)
			(MSXML2::ISAXLocator *loc, wchar_t *msg, HRESULT hr)
		{
			SetMsg(loc, msg, hr);
			return E_FAIL;
		}
	};

/// <summary>Validate xml document</summary>
/// <param name="strxml">XML string</param>
/// <param name="errcol">Column of error (out)</param>
/// <param name="errline">Line of error (out)</param>
/// <param name="errmessage">Error message (out)</param>
/// <returns>true on success</returns>
	bool Doc::Validate(int& errcol, int& errline, CString& errmessage)
	{
		try
		{
			// create a schema collection
			MSXML2::IXMLDOMSchemaCollection2Ptr scol;
			CheckError(scol.CreateInstance(L"Msxml2.XMLSchemaCache.6.0"));

			// load fictionbook schema
			scol->add(U::FBNS, (const wchar_t *)U::GetProgDirFile(L"FictionBook.xsd"));

			// create a SAX reader
			MSXML2::ISAXXMLReaderPtr rdr;
			CheckError(rdr.CreateInstance(L"Msxml2.SAXXMLReader.6.0"));

			// attach a schema
			rdr->putFeature(L"schema-validation", VARIANT_TRUE);
			rdr->putProperty(L"schemas", scol.GetInterfacePtr());
			rdr->putFeature(L"exhaustive-errors", VARIANT_TRUE);

			// create an error handler
			CComObject<SAXErrorHandler> *ehp;
			CheckError(CComObject<SAXErrorHandler>::CreateInstance(&ehp));
			CComPtr<CComObject<SAXErrorHandler>> eh(ehp);
			rdr->putErrorHandler(eh);

			IStreamPtr	isp(m_saved_xml);
			HRESULT hr = rdr->raw_parse(_variant_t((IUnknown *)isp));

			if (FAILED(hr))
			{
				errcol = eh->m_col;
				errline = eh->m_line;

				CString strLine;
				CString strCol;
				strLine.Format(L"%d", errline);
				strCol.Format(L"%d", errcol);
				errmessage = L"ln:" + strLine + L", col:" + strCol + L"\n" + eh->m_msg;
				return false;
			}
		}
		catch (_com_error &e)
		{
			errmessage.SetString(e.Description());
			errcol = 0;
			errline = 0;
			return false;
		}
		return true;
	}

	// binaries
	void Doc::AddBinary(const CString &filename)
	{
		m_body.AttachImage(filename);
	}
	 
	//?????? 
	MSHTML::IHTMLDOMNodePtr Doc::MoveNode(MSHTML::IHTMLDOMNodePtr from, MSHTML::IHTMLDOMNodePtr to, MSHTML::IHTMLDOMNodePtr insertBefore)
	{
		VARIANT disp;
		// MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElementPtr)to;
		// bstr_t text = elem->innerHTML;

		//  title
		if (insertBefore)
		{
			while (1)
			{
				MSHTML::IHTMLElementPtr elem = (MSHTML::IHTMLElementPtr)insertBefore;
                CString class_name = elem->className;
				// _bstr_t class_name(elem->className);
				if (class_name == L"title" || class_name == L"epigraph" || class_name == L"annotation" || class_name == L"image")
				{
					insertBefore = insertBefore->nextSibling;
					continue;
				}
				break;
			}
		}

		MSHTML::IHTMLDOMNodePtr ret;
		if (insertBefore)
		{
			disp.pdispVal = insertBefore;
			disp.vt = VT_DISPATCH;
			ret = to->insertBefore(from, disp);
		}
		else
		{
			ret = to->appendChild(from);
		}

		return ret;
	}

	//---------Words list ----------------
	static int compare_nocase(const void *v1, const void *v2)
	{
		CString *s1 = (CString *)v1;
		CString *s2 = (CString *)v2;

		int cv = s1->CompareNoCase(*s2);
		if (cv != 0)
			return cv;

		return s1->Compare(*s2);
	}

	static int compare_counts(const void *v1, const void *v2)
	{
		const Doc::Word *w1 = (const Doc::Word *)v1;
		const Doc::Word *w2 = (const Doc::Word *)v2;
		int diff = w1->count - w2->count;
		return diff ? diff : w1->word.CompareNoCase(w2->word);
	}

	void Doc::GetWordList(int flags, CSimpleArray<Word> &words, CString tagName)
	{
		CWaitCursor hourglass;

		MSHTML::IHTMLElementPtr fbw_body = MSHTML::IHTMLDocument3Ptr(m_body.Document())->getElementById(L"fbw_body");
		MSHTML::IHTMLElementCollectionPtr paras = MSHTML::IHTMLElement2Ptr(fbw_body)->getElementsByTagName(L"P");
		if (!paras->length)
			return;

		int iNextElem = 0;

		// Construct a word list
		CSimpleArray<CString> wl;

		while (iNextElem < paras->length)
		{
			MSHTML::IHTMLElementPtr currElem(paras->item(iNextElem));
			CString innerText = currElem->innerText;

			MSHTML::IHTMLDOMNodePtr currNode(currElem);
			if (MSHTML::IHTMLElementPtr siblElem = currNode->nextSibling)
			{
				int jNextElem = iNextElem + 1;
				for (int i = jNextElem; i < paras->length; ++i)
				{
					MSHTML::IHTMLElementPtr nextElem = paras->item(i);
					if (siblElem == nextElem)
					{
						innerText += CString(L"\r\n") + siblElem->innerText.GetBSTR();
						iNextElem++;
						siblElem = MSHTML::IHTMLDOMNodePtr(nextElem)->nextSibling;
					}
					else
					{
						break;
					}
				}
			}

			_bstr_t bb(innerText.AllocSysString());

			if (bb.length() == 0)
			{
				iNextElem++;
				continue;
			}

			// iterate over bb using a primitive fsm
			wchar_t *p = bb, *e = p + bb.length() + 1; // include trailing 0!
			wchar_t *wstart = nullptr;
			wchar_t *wend = nullptr;

			enum
			{
				INITIAL,
				INWORD1,
				INWORD2,
				HYPH1,
				HYPH2
			} state = INITIAL;

			while (p < e)
			{
				int letter = iswalpha(*p);
				switch (state)
				{
				case INITIAL:
				initial:
					if (letter)
					{
						wstart = p;
						state = INWORD1;
					}
					break;
				case INWORD1:
					if (!letter)
					{
						if (flags & GW_INCLUDE_HYPHENS)
						{
							if (iswspace(*p))
							{
								wend = p;
								state = HYPH1;
								break;
							}
							else if (*p == L'-')
							{
								wend = p;
								state = HYPH2;
								break;
							}
						}
						if (!(flags & GW_HYPHENS_ONLY))
						{
							*p = L'\0';
							wl.Add(wstart);
						}
						state = INITIAL;
					}
					break;
				case INWORD2:
					if (!letter)
					{
						*p = L'\0';
						U::RemoveSpaces(wstart);
						wl.Add(wstart);
						state = INITIAL;
					}
					break;
				case HYPH1:
					if (*p == L'-')
						state = HYPH2;
					else if (!iswspace(*p))
					{
						if (!(flags & GW_HYPHENS_ONLY))
						{
							*wend = L'\0';
							wl.Add(wstart);
						}
						state = INITIAL;
						goto initial;
					}
					break;
				case HYPH2:
					if (letter)
						state = INWORD2;
					else if (!iswspace(*p))
					{
						if (!(flags & GW_HYPHENS_ONLY))
						{
							*wend = L'\0';
							wl.Add(wstart);
						}
						state = INITIAL;
						goto initial;
					}
					break;
				}
				++p;
			}

			iNextElem++;
		}

		if (wl.GetSize() == 0)
			return;

		// now sort the list
		qsort(wl.GetData(), wl.GetSize(), sizeof(CString), compare_nocase);

		int wlSize = wl.GetSize();
		for (int i = 0; i < wlSize; ++i)
		{
			int count = 1, k = 0;
			for (int j = i + 1; j < wlSize; ++j)
			{
				if (wl[i] == wl[j])
					count++;
				else
				{
					k = --j;
					break;
				}

				k = j;
			}
			words.Add(Word(wl[i], count));
			if (k)
				i = k;
		}

		// Sort by count now
		if (flags & GW_SORT_BY_COUNT)
			qsort(words.GetData(), words.GetSize(), sizeof(Word), compare_counts);
	}

	bool Doc::LoadXMLFromString(BSTR text)
	{
		if (MSXML2::IXMLDOMDocumentPtr(m_saved_xml)->loadXML(text) == FALSE)
			return false;

		// add properties
		m_saved_xml->setProperty(L"SelectionLanguage", L"XPath");
		m_saved_xml->setProperty(L"SelectionNamespaces", L"xmlns:fb='" + U::FBNS + L"' xmlns:l='" + U::XLINKNS + L"'");
		return true;
	}

	bool Doc::BuildHTML()
	{
		try
		{
			CString dd = MSXML2::IXMLDOMNodePtr(m_saved_xml)->xml;
			CComDispatchDriver body1(m_body.Script());
			CComVariant args[2];
			args[1] = m_saved_xml.GetInterfacePtr();
			args[0] = _Settings.GetInterfaceLanguageName();
			CheckError(body1.InvokeN(L"LoadFromDOM", args, 2));
			m_body.Init();
		}
		catch (_com_error(&e))
		{
			U::ReportError(e);
			return false;
		}
		return true;
	}

    /*	void Doc::GoTo(int pos)
	{
		MSHTML::IHTMLElementPtr fbw_body = MSHTML::IHTMLDocument3Ptr(m_body.Document())->getElementById(L"fbw_body");
		MSHTML::IHTMLDOMRangePtr range = MSHTML::IDocumentRangePtr(m_body.Document())->createRange();
		if (!range)
			return;
		MSHTML::IHTMLSelectionPtr sel = MSHTML::IHTMLDocument7Ptr(m_body.Document())->getSelection();
		if (!sel)
			return;

		int real_pos = 0;

		//int real_pos = pos < 0 ? 0: pos;
		
		range->setStart(fbw_body, real_pos);
		range->collapse(VARIANT_TRUE);
		sel->removeAllRanges();
		sel->addRange(range);
	}
*/
} // namespace FB
