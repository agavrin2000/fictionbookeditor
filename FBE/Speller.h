#ifndef SPELLER_H
#define SPELLER_H

#include <map>
#include <set>
#include <atlutil.h>
#include "utils.h"
#include "ModelessDialog.h"

// Hunspell API function prototypes
extern "C"
{
#include <hunspell/hunspell.h>
}

typedef CSimpleArray<CString> CStrings;
class CSpeller;

class CSpellDialog : public CModelessDialogImpl<CSpellDialog>
{
  public:
	enum
	{
		IDD = IDD_SPELL_CHECK
	};

	CString m_sBadWord;
	CString m_sReplacement;
	CStrings *m_strSuggestions;

	CSpellDialog(CSpeller *parent) : m_Speller(parent), m_strSuggestions(NULL) {}

	BEGIN_MSG_MAP(CSpellDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_ACTIVATE, OnActivate)
	COMMAND_ID_HANDLER(IDC_SPELL_IGNORE, OnIgnore)
	COMMAND_ID_HANDLER(IDC_SPELL_IGNOREALL, OnIgnoreAll)
	COMMAND_ID_HANDLER(IDC_SPELL_CHANGE, OnChange)
	COMMAND_ID_HANDLER(IDC_SPELL_CHANGEALL, OnChangeAll)
	COMMAND_ID_HANDLER(IDC_SPELL_ADD, OnAdd)
	COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
	COMMAND_HANDLER(IDC_SPELL_REPLACEMENT, EN_CHANGE, OnEditChange)
	COMMAND_HANDLER(IDC_SPELL_SUGG_LIST, LBN_SELCHANGE, OnSelChange)
	COMMAND_HANDLER(IDC_SPELL_SUGG_LIST, LBN_DBLCLK, OnSelDblClick)
	END_MSG_MAP()

	LRESULT UpdateData();

  private:
	CSpeller *m_Speller;
	//int m_RetCode;

	bool m_WasSuspended;
	// controls
	CEdit m_BadWord;
	CEdit m_Replacement;
	CListBox m_Suggestions;
	CButton m_IgnoreContinue;

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnActivate(UINT, WPARAM, LPARAM, BOOL &);
	LRESULT OnIgnore(WORD, WORD wID, HWND, BOOL &);
	LRESULT OnIgnoreAll(WORD, WORD wID, HWND, BOOL &);
	LRESULT OnChange(WORD, WORD wID, HWND, BOOL &);
	LRESULT OnChangeAll(WORD, WORD wID, HWND, BOOL &);
	LRESULT OnAdd(WORD, WORD wID, HWND, BOOL &);
	LRESULT OnSelChange(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled);
	LRESULT OnSelDblClick(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled);
	LRESULT OnEditChange(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled);
	LRESULT OnCancel(WORD, WORD wID, HWND, BOOL &);
};

typedef std::pair<long, MSHTML::IHighlightSegmentPtr> HIGHLIGHT;
typedef std::multimap<long, MSHTML::IHighlightSegmentPtr> HIGHLIGHTS;

class CSpeller
{
	// Dictionary info
	struct DICT
	{
		Hunhandle *handle;
		CString lang;
		CString filename;
	};

  public:
	CSpeller(CString);
	~CSpeller();
	void AttachDocument(MSHTML::IHTMLDocumentPtr doc);
	bool Enabled() 
    { 
        return m_Enabled; 
    }
	void Enable(bool state);
	void SetDocumentLanguage(CString lang);
	void SetCustomDictionary(CString pathToDictionary);
	void StartDocumentCheck(MSHTML::IMarkupServices2Ptr undoSrv = nullptr);
	void EndDocumentCheck(bool bCancel = true);
	void ContinueDocumentCheck();

	void Undo();
	bool GetUndoState();
	void BeginUndoUnit(const wchar_t *name);
	void EndUndoUnit();

	void CheckScroll();
	void CheckElement(MSHTML::IHTMLElementPtr elem, bool HTMLChanged);
	void doSpellCheck();

	CStrings *GetSuggestions();
	void Replace(CString word);
	void ReplaceAll(CString badWord, CString goodWord);
	void IgnoreAll(CString word = L"");
	void AddToDictionary(CString word = L"");

  private:
	SHD::IWebBrowser2Ptr m_browser;
	MSHTML::IHTMLDocument2Ptr m_doc2;
	MSHTML::IHTMLElement2Ptr m_scrollElement;
	MSHTML::IHTMLRenderStylePtr m_irs;
	MSHTML::IMarkupServicesPtr m_ims;
	MSHTML::IDisplayServicesPtr m_ids;
	MSHTML::IHighlightRenderingServicesPtr m_ihrs;

	// for document checking
	MSHTML::IHTMLDOMRangePtr m_prevSelection;
	MSHTML::IHTMLDOMRangePtr m_curSelection;

	MSHTML::IMarkupServices2Ptr m_undoSrv;
	std::set<long> m_uniqIDs;

	CSpellDialog *m_spell_dlg;

	bool m_Enabled;
	CString m_Lang;
	CString m_DictPath;
	CString m_CustomDictPath;

	int m_prevY;
	int m_numAphChanged;

	CSimpleArray<CSpeller::DICT> m_Dictionaries;
	CStrings m_CustomDict;

	CStrings m_IgnoreWords;
	CStrings *m_menuSuggestions;
	HIGHLIGHTS m_ElementHighlights;

	Hunhandle *LoadDictionary(CString dictName);
	void getAvailDictionary();

	MSHTML::IHTMLDOMRangePtr DetectSelWordRange();
	bool SelectNextWord();
	CString GetSelWord();

	bool SpellCheck(CString word);
	CStrings *GetSuggestions(CString word);
	CString detectWordLanguage(const CString &word);

	void ClearMark(int elemID);
	void ClearAllMarks();
	void MarkWord(MSHTML::IHTMLElementPtr elem, long uniqID, CString word, int pos);
	void MarkElement(MSHTML::IHTMLElementPtr elem, long uniqID);

};

#endif
