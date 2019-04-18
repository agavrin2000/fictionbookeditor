// FBDoc.h: interface for the FBDoc class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_FBDOC_H__205AB072_350A_4D2F_B58B_61BBC255FFE5__INCLUDED_)
#define AFX_FBDOC_H__205AB072_350A_4D2F_B58B_61BBC255FFE5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "FBEView.h"
#include "zip.h"

namespace FB // put all FB2 related stuff into its own namespace
{
// an fb2 document in-memory representation
class Doc
{
public:
	Doc(HWND hWndFrame);
	~Doc();
  
  // document text is stored here
  CFBEView		m_body;
  MSXML2::IXMLDOMDocument2Ptr m_saved_xml;
    
  CString		m_filename; // filename
  CString		m_file_encoding; // file encoding
  
  // loading and saving
  void	  CreateBlank(HWND hWndParent);
  bool	  Load(HWND hWndParent,const CString& filename);
  bool	  Save(const CString& filename);
  CString GetDocFileName() const;

  void CreateDOM();
  bool LoadXMLFromString(BSTR text);
  bool BuildHTML();
  bool Validate(int& errcol, int& errline, CString& errmessage);
  
  // transfer selection between views
  bool SetHTMLSelectionFromXML(int selectedPosBegin, int selectedPosEnd, BSTR ustr);
  bool GetXMLSelectionFromHTML(int& selectedPosBegin, int& selectedPosEnd, _bstr_t& ustr);
  void ClearHTMLSelection();
  void SaveHTMLSelection();
  void RestoreHTMLSelection();
  
  HRESULT InvokeFunc(BSTR FuncName, CComVariant *params, int count, CComVariant &vtResult);
  void	  RunScript(BSTR filePath);
  bool LoadScript(BSTR filePath);

  // changes
  bool IsChanged()
  {
	  return m_doc_changed;
  }

  void Changed(bool state)
  {
	  m_doc_changed = state;
  }

  bool IsTraceChangesEnabled() 
  {
	  return m_trace_changes;
  }

  void TraceChanges(bool state)
  {
	  m_trace_changes = state;
  }

  // binary objects
  //BSTR PrepareDefaultId(const CString& filename);
  void AddBinary(const CString& filename);
  bool GetBinary(const wchar_t *id,_variant_t& vt);

  // config
  //void	  ApplyConfChanges();

  // active document table
  static Doc* m_active_doc;


	// Word lists
	struct Word
	{
		CString word;
		CString replacement;
		int count;
		int flags;

		Word() : count(0), flags(0) { }
		Word(CString& word, int count) : word(word), count(count), flags(0) { }
	};
	enum
	{
		GW_INCLUDE_HYPHENS	= 1,
		GW_HYPHENS_ONLY		= 2,
		GW_SORT_BY_COUNT	= 4
	};
	void GetWordList(int flags, CSimpleArray<Word>& words, CString tagName);
	MSHTML::IHTMLDOMNodePtr MoveNode(MSHTML::IHTMLDOMNodePtr from, MSHTML::IHTMLDOMNodePtr to, MSHTML::IHTMLDOMNodePtr insertBefore);
	//void GoTo(int pos);
private:
	MSHTML::IHTMLDOMRangePtr m_body_sel;
	bool m_namevalid;
	bool m_trace_changes;
	bool m_doc_changed;

	CString MyID() 
	{ 
		CString ret; 
		ret.Format(_T("%lu"), (unsigned long)this); 
		return ret; 
	}
	CString MyURL(const wchar_t *part) 
    { 
        CString ret; 
        ret.Format(_T("fbw-internal:%lu:%s"), (unsigned long)this, part); 
        return ret; 
    }

	static CSimpleMap<Doc*, Doc*>	m_active_docs;
	bool	  SaveToFile(const CString& filename);
};

} // namespace FB
#endif // !defined(AFX_FBDOC_H__205AB072_350A_4D2F_B58B_61BBC255FFE5__INCLUDED_)
