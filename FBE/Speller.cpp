#include "stdafx.h"
#include "resource.h"
#include "res1.h"
#include <iostream>
#include <fstream>
#include "FBE.h"
#include "Speller.h"

// spell check dialog initialisation
LRESULT CSpellDialog::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	m_BadWord = GetDlgItem(IDC_SPELL_BEDWORD);
	m_Replacement = GetDlgItem(IDC_SPELL_REPLACEMENT);
	m_Suggestions = GetDlgItem(IDC_SPELL_SUGG_LIST);
	m_IgnoreContinue = GetDlgItem(IDC_SPELL_IGNORE);
	m_WasSuspended = false;

	UpdateData();
	CenterWindow();
	return 1;
}

LRESULT CSpellDialog::OnActivate(UINT, WPARAM wParam, LPARAM, BOOL&)
{
	if (wParam == WA_INACTIVE)
	{
		CString txt;
		txt.LoadString(IDS_SPELL_CONTINUE);
		m_IgnoreContinue.SetWindowText(txt);

		GetDlgItem(IDC_SPELL_IGNOREALL).EnableWindow(FALSE);
		GetDlgItem(IDC_SPELL_CHANGE).EnableWindow(FALSE);
		GetDlgItem(IDC_SPELL_CHANGEALL).EnableWindow(FALSE);
		GetDlgItem(IDC_SPELL_ADD).EnableWindow(FALSE);
		m_Replacement.EnableWindow(FALSE);
		m_Suggestions.EnableWindow(FALSE);
		m_WasSuspended = true;

		UpdateWindow();
	}
	return 0;
}

LRESULT CSpellDialog::UpdateData()
{
	m_BadWord.SetWindowText(m_sBadWord);
	m_Suggestions.ResetContent();
	if (m_strSuggestions)
		for (int i = 0; i < m_strSuggestions->GetSize(); i++)
		{
			m_Suggestions.AddString((*m_strSuggestions)[i]);
			if (!i) m_Replacement.SetWindowText((*m_strSuggestions)[i]);
		}

	if (m_Speller->GetUndoState())
		GetDlgItem(IDC_SPELL_UNDO).EnableWindow(TRUE);
	else
		GetDlgItem(IDC_SPELL_UNDO).EnableWindow(FALSE);

	return 1;
}

/// <summary>Select suggestion in list</summary>
LRESULT CSpellDialog::OnSelChange(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CString strText;
	if (m_Suggestions.GetText(m_Suggestions.GetCurSel(), strText))
		m_Replacement.SetWindowText(strText);
	return 0;
}

// change text to suggested word on doubleclick
LRESULT CSpellDialog::OnSelDblClick(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CString strText;
	if (m_Suggestions.GetText(m_Suggestions.GetCurSel(), strText))
	{
		m_Replacement.SetWindowText(strText);
		OnChange(0, 0, 0, bHandled);
	}
	return 0;
}

LRESULT CSpellDialog::OnEditChange(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	m_Replacement.GetWindowText(m_sReplacement);
	return 0;
}

/// <summary>Cancel Spell Dialog</summary>
LRESULT CSpellDialog::OnCancel(WORD, WORD wID, HWND, BOOL&)
{
	m_Speller->EndDocumentCheck();
	return 0;
}

/// <summary>Skip misspelled word</summary>
LRESULT CSpellDialog::OnIgnore(WORD, WORD wID, HWND, BOOL&)
{
	if (m_WasSuspended)
	{
		CString txt;
		txt.LoadString(IDC_SPELL_IGNORE);
		m_IgnoreContinue.SetWindowText(txt);

		GetDlgItem(IDC_SPELL_IGNOREALL).EnableWindow(TRUE);
		GetDlgItem(IDC_SPELL_CHANGE).EnableWindow(TRUE);
		GetDlgItem(IDC_SPELL_CHANGEALL).EnableWindow(TRUE);
		GetDlgItem(IDC_SPELL_ADD).EnableWindow(TRUE);
		m_Replacement.EnableWindow(TRUE);
		m_Suggestions.EnableWindow(TRUE);
		m_WasSuspended = false;

		UpdateWindow();
		m_Speller->StartDocumentCheck();
	}
	else
		m_Speller->ContinueDocumentCheck();

	return 0;
}

/// <summary>Ignore all such misspelled words</summary>
LRESULT CSpellDialog::OnIgnoreAll(WORD, WORD wID, HWND, BOOL&)
{
	m_Speller->IgnoreAll(m_sBadWord);
	m_Speller->ContinueDocumentCheck();
	return 0;
}

/// <summary>Replace misspelled words once</summary>
LRESULT CSpellDialog::OnChange(WORD, WORD wID, HWND, BOOL&)
{
	m_Speller->BeginUndoUnit(L"replace word");
	m_Speller->Replace(m_sReplacement);
	m_Speller->EndUndoUnit();
	m_Speller->ContinueDocumentCheck();
	return 0;
}

/// <summary>Replace all such misspelled words</summary>
LRESULT CSpellDialog::OnChangeAll(WORD, WORD wID, HWND, BOOL&)
{
	
	m_Speller->BeginUndoUnit(L"replace word");
	m_Speller->ReplaceAll(m_sBadWord, m_sReplacement);
	m_Speller->EndUndoUnit();
	m_Speller->ContinueDocumentCheck();
	return 0;
}

/// <summary>Add word to custom dictionary</summary>
LRESULT CSpellDialog::OnAdd(WORD, WORD wID, HWND, BOOL&)
{
	m_Speller->AddToDictionary(m_sBadWord);
	m_Speller->ContinueDocumentCheck();
	return 0;
}


/////////////////////////////////////////////////////////////////////////////////////////////
//
// CSpeller methods
//
////////////////////////////////////////////////////////////////////////////////////////////

//
// CSpeller constructor
//
CSpeller::CSpeller(CString& dictPath) :
	m_prevY(0), m_Lang(""), m_Enabled(false), m_spell_dlg(0), m_prevSelRange(0), m_CustomDictPath(L"")
{
	m_DictPath = dictPath;

	// get all available dictionary
	getAvailDictionary();
	// but load only English dictionary 
	for (int i = 0; i<m_Dictionaries.GetSize(); i++) 
	{
		if (m_Dictionaries[i].lang == L"en")
			m_Dictionaries[i].handle = LoadDictionary(m_Dictionaries[i].filename);
        HRESULT GetCodePageInfo(
  [in]  UINT        uiCodePage,
  [out] PMIMECPINFO pCodePageInfo
);
	}
}

/// <summary>Init all dictionaries in folder</summary>
void CSpeller::getAvailDictionary()
{
	WIN32_FIND_DATA ffd;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	
	hFind = FindFirstFile(m_DictPath+L"*.dic", &ffd);
	if (INVALID_HANDLE_VALUE == hFind)
		return;

	do
	{
		CString fname = ffd.cFileName;
		int idx = fname.ReverseFind(L'.');
		if (idx != -1) {
			fname = fname.Left(idx);
		}
		m_Dictionaries.Add({0, fname.Left(2), fname});
	} while (FindNextFile(hFind, &ffd) != 0);

	FindClose(hFind);
	return;
}
// CSpeller destructor
CSpeller::~CSpeller()
{
	EndDocumentCheck();
	// unload dictionaries
/*	for (int i = 0; i < m_Dictionaries.GetSize(); i++)
		if (m_Dictionaries[i].handle)
			Hunspell_destroy(m_Dictionaries[i].handle);
			*/
	for (int i = 0; i < m_Dictionaries.GetSize(); i++) 
	{
		if (m_Dictionaries[i].handle)
			Hunspell_destroy(m_Dictionaries[i].handle);
	}
}

/// <summary>Load Dictionary from file</summary>
/// <param name="dictName">Dictionary filename</param>  
/// <returns>Dic handle or NULL</returns> 
Hunhandle* CSpeller::LoadDictionary(CString dictName)
{
	Hunhandle* dict = nullptr;

	if (PathFileExistsW(m_DictPath + dictName + L".aff") && PathFileExistsW(m_DictPath + dictName + L".dic"))
	{
		// create dictionary from file
		CW2A affpath(m_DictPath + dictName + L".aff", CP_UTF8);
		CW2A dpath(m_DictPath + dictName + L".dic", CP_UTF8);
		dict = Hunspell_create(affpath, dpath);
        CString cp = Hunspell_get_dic_encoding(dict);
        if (cp != L"UTF-8")
        {
            //load only utf-8 dictionary
            Hunspell_destroy(dict);
            return nullptr;
        }
            
	}
	return dict;
}

/// <summary>Init Custom Dictionary</summary>
/// <param name="pathToDictionary">Path to dictionary</param>  
/// <returns>Dic handle or NULL</returns> 
void CSpeller::SetCustomDictionary(CString pathToDictionary)
{
	m_CustomDictPath = pathToDictionary;
	m_CustomDict.RemoveAll();
	
	//init dictionary
	CString str;
	char buf[256];
	try
	{
		std::ifstream load;
		load.open(m_CustomDictPath);
		if (load.is_open())
		{
			do
			{
				load.getline(&buf[0], sizeof(buf), '\n');
				str.SetString(CA2W(buf, CP_UTF8));
				if (!str.IsEmpty())
					m_CustomDict.Add(str);
			} while (!str.IsEmpty());
			load.close();
		}
	}
	catch (...)
	{
		// don't inform about error, do it in Add action
	}
}

/// <summary>Attach new document</summary>
/// <param name="doc">HTML document</param>  
void CSpeller::AttachDocument(MSHTML::IHTMLDocumentPtr doc)
{
	// cancel spell check and destroy dialog
	EndDocumentCheck();

	// clear previous marks
	m_ElementHighlights.clear();

	// clear collected ID's
	m_uniqIDs.clear();

	// clear (assigned to previous document) service arrays
	m_IgnoreWords.RemoveAll();

	// assign new document: all interfaces and variables
	m_doc2 = MSHTML::IHTMLDocument2Ptr(doc);

	// get web browser component
	MSHTML::IHTMLWindow2Ptr pWin(m_doc2->parentWindow);
	IServiceProviderPtr pISP(pWin);
	pISP->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void **)&m_browser);

	m_scrollElement = MSHTML::IHTMLDocument3Ptr(doc)->documentElement;
	m_mkc = MSHTML::IMarkupContainer2Ptr(MSHTML::IHTMLDocument4Ptr(doc));
	m_ims = MSHTML::IMarkupServicesPtr(MSHTML::IHTMLDocument4Ptr(doc));
	m_ihrs = MSHTML::IHighlightRenderingServicesPtr(MSHTML::IHTMLDocument4Ptr(doc));
	m_ids = MSHTML::IDisplayServicesPtr(MSHTML::IHTMLDocument4Ptr(doc));

	// create a render style (wavy red line)
	_bstr_t b;
	m_irs = MSHTML::IHTMLDocument4Ptr(doc)->createRenderStyle(b);
	m_irs->defaultTextSelection = "false";
	m_irs->textBackgroundColor = "transparent";
	m_irs->textColor = "transparent";
	m_irs->textDecoration = "underline";
	m_irs->textDecorationColor = "red";
	m_irs->textUnderlineStyle = "wave";

	// detect document language
	m_Lang = "";
	MSHTML::IHTMLSelectElementPtr elem = MSHTML::IHTMLDocument3Ptr(m_doc2)->getElementById(L"tiLang");
	CString lang = elem->value;
	if (elem) {
		for (int i = 0; i < m_Dictionaries.GetSize(); i++)
		{
			if (m_Dictionaries[i].lang == lang)
			{
				// load dictionary (if needed)
				if (!m_Dictionaries[i].handle)
					m_Dictionaries[i].handle = LoadDictionary(m_Dictionaries[i].filename);
				m_Lang = lang;
				break;
			}
		}
	}
	// initiate background check
	doSpellCheck();
}

/// <summary>Set Enabled state
/// and cspellcheck or clear marks</summary>
/// <param name="Enabled">new Enabled state</param>  
void CSpeller::Enable(bool state)
{
	if (m_Enabled != state)
	{
		m_Enabled = state;
		if (!m_Enabled)
			ClearAllMarks();
		else
			doSpellCheck();
	}
}

/// <summary>Set doc language bases on Title Info and load dictionaries</summary>
/// <param name="lang">Language</param>  
void CSpeller::SetDocumentLanguage(CString lang)
{
	if (lang.IsEmpty())
	{
		m_Lang = "";
		return;
	}
	bool isLangChanged = false;

	for (int i = 0; i < m_Dictionaries.GetSize(); i++)
	{
		if (m_Dictionaries[i].lang == lang)
		{
			// load dictionary (if needed)
			if (!m_Dictionaries[i].handle)
				m_Dictionaries[i].handle = LoadDictionary(m_Dictionaries[i].filename);
			if (m_Lang != lang)
			{
				isLangChanged = true;
				m_Lang = lang;
			}
			break;
		}
	}

	if (isLangChanged)
		doSpellCheck();
}

/// <summary>Undo changes</summary>
void CSpeller::Undo()
{
	if (m_browser)
	{
		IOleCommandTargetPtr ct(m_browser);
		if (ct)
			ct->Exec(&CGID_MSHTML, IDM_UNDO, 0, NULL, NULL);
	}
}

/// <summary>Get Undo coomand status</summary>
/// <returns>true - in Undo state</returns>
bool CSpeller::GetUndoState()
{
	bool stat = false;
	if (m_browser)
	{
		static OLECMD cmd[] = { IDM_UNDO };
		IOleCommandTargetPtr ct(m_browser);
		if (ct)
		{
			ct->QueryStatus(&CGID_MSHTML, 1, cmd, NULL);
			stat = (cmd[0].cmdf != 1);
		}
	}
	return stat;
}

/// <summary>Begin Undo block</summary>
/// <param name="name">Undo block name</param>  
void CSpeller::BeginUndoUnit(const wchar_t *name)
{
	m_undoSrv->BeginUndoUnit((wchar_t *)name);
}

/// <summary>End Undo block</summary>
void CSpeller::EndUndoUnit()
{
	m_undoSrv->EndUndoUnit();
}

/// <summary>Get TextRange of selected word</summary>
/// <returns>TxtRangePtr with word, NULL - if several or none words selected</returns> 
MSHTML::IHTMLTxtRangePtr CSpeller::GetSelWordRange()
{
	MSHTML::IHTMLTxtRangePtr rng = nullptr;
	CString s = L"";

	// fetch exist selection
	MSHTML::IHTMLTxtRangePtr selRange(m_doc2->selection->createRange());
	if (selRange) {
		s.SetString(selRange->text);
		// if selected text block contains 2+ words return, 
		// we don't know which word selected
		for (int i = 0; i<int(s.GetLength()); i++)
		{
			if(iswspace(s.GetAt(i)))
				return nullptr;
		}
		rng = selRange->duplicate();
	}
	if (!rng)
		return nullptr;
	rng->collapse(true); //collapse to start point
	
	// try to detect word
	// IHTMLTextRange expand() has strange behavior
	// before space look word backward, return word with space
	// after space look word forward, return word with space
	// before puctuation (and hyphen) return punctuation
	// before \n return empty string

	int endMoved = rng->moveEnd(L"character", 1);
	s.SetString(rng->text);

	CString charAfter = s.Left(1);
	if (endMoved == 1) {
		endMoved = rng->moveEnd(L"character", -1);
	}

	int startMoved = rng->moveStart(L"character", -1);
	s.SetString(rng->text);

	CString charBefore = s.Right(1);
	if (startMoved == -1) {
		rng->moveStart(L"character", 1);
	}

	// case: between space or punctuation or \n
	if ((iswspace(charBefore.GetAt(0)) || iswpunct(charBefore.GetAt(0)) || charBefore.GetLength() == 0) &&
		(iswspace(charAfter.GetAt(0)) || iswpunct(charAfter.GetAt(0)) || charAfter.GetLength() == 0)) {
		return nullptr;
	}
	// case: between punctuation
	if (iswpunct(charBefore.GetAt(0)) && iswpunct(charAfter.GetAt(0))) {
		return nullptr;
	}

	// case: take previous word
	if ((iswalpha(charBefore.GetAt(0)) || iswdigit(charBefore.GetAt(0))) &&
		(iswspace(charAfter.GetAt(0)) || iswpunct(charAfter.GetAt(0)) || charAfter.GetLength() == 0)) {
			rng->moveStart(L"word", -1);
	}
	// case: take next word or word on carret
	// use trick when "expand()" get forward word or get word on carret
	else {
		rng->expand(L"word");
		// MS expand() can catch space at the end, so remove it
		s.SetString(rng->text);
		if (iswspace(s.GetAt(s.GetLength()-1)))
			rng->moveEnd(L"character", -1);
	}
	return rng;
}

/// <summary>Get word under caret or selected</summary>
/// <returns>Word</returns> 
CString CSpeller::GetSelWord()
{
	CString word(L"");
	MSHTML::IHTMLTxtRangePtr range = GetSelWordRange();
	if (range) {
		word.SetString(range->text);
	}
	return word;
}

/// <summary>Get suggestions for current selected word</summary>
CStrings* CSpeller::GetSuggestions()
{
	CString word = GetSelWord();
	if (word.IsEmpty())
        return nullptr;

	if (!SpellCheck(word))
	{
		return GetSuggestions(word);
	}
    return nullptr;
}

/// <summary>Replace misspelled word with the correct</summary>
/// <param name="word">Correct word</param>  
void CSpeller::Replace(CString word)
{
	MSHTML::IHTMLTxtRangePtr range = GetSelWordRange();
	if (m_numAphChanged)
		word.Replace(L"'", L"’");
	range->put_text(_bstr_t(word));
}

/// <summary>Replace misspelled word with the correct in ALL text</summary>
/// <param name="badWord">misspelled word</param>  
/// <param name="goodWord">Correct word</param>  
void CSpeller::ReplaceAll(CString badWord, CString goodWord)
{
	MSHTML::IHTMLElementPtr fbw_body = MSHTML::IHTMLDocument3Ptr(m_doc2)->getElementById(L"fbw_body");
	MSHTML::IHTMLTxtRangePtr range2 = m_doc2->selection->createRange();
	range2->moveToElementText(fbw_body);
	MSHTML::IHTMLTxtRangePtr range = range2->duplicate();

	while (range->findText(_bstr_t(badWord), 0, 2))
	{
		CString strText = range->text;
		range->put_text(_bstr_t(goodWord));
		range->setEndPoint("EndToEnd", range2); //extend to end of body
	}
}

/// <summary>Add word to ignore list and recheck</summary>
/// <param name="word">Word to ignore</param>  
void CSpeller::IgnoreAll(CString word)
{
	if (word.IsEmpty())
		word = GetSelWord();
	if (!SpellCheck(word))
	{
		m_IgnoreWords.Add(word);
		// recheck page
		ClearAllMarks();
		doSpellCheck();
	}
}

/// <summary>Add word to custom dictionary and recheck page</summary>
/// <param name="word">Word to add, if empty take word on caret</param>  
void CSpeller::AddToDictionary(CString word)
{
	if (word == L"")
		word = GetSelWord();
	if (word == L"")
		return;
	//Hunhandle* currDict = GetDictionary(word);

	// add to Hunspell's runtime dictionary
	//CW2A str(word, CP_UTF8);
	//Hunspell_add(currDict, str);

	// add to custom dictionary & save custom dictionary
	// check in custom dictionary
	if (m_CustomDict.Find(word) != -1)
		return;
	m_CustomDict.Add(word);
	try
	{
		std::ofstream save;
		save.open(m_CustomDictPath, std::ios_base::out | std::ios_base::app);
		if (save.is_open()) {
			word.Replace(L"\u00AD", L"");
			CT2A str(word, CP_UTF8);
			save << str << '\n';
			save.close();
		}
	}
	catch (...) 
	{
		// L"Can't add to dictionary:" + m_CustomDictPath
		AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, IDS_SPELL_CHECK_COMPLETED, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_INFORMATION_ICON);
	}

	// recheck page
	ClearAllMarks();
	doSpellCheck();
}

/// <summary>Get languge based on word (bi-lingual hack)</summary>
/// <param name="word">Word</param>  
/// <returns>Lanhuage</returns> 
CString CSpeller::detectWordLanguage(const CString& word)
{
	CString detectLang = m_Lang;

	// special fix  for english words at Cyrillic document
	if (m_Lang == "ru" || m_Lang == "uk" || m_Lang == "be")
	{
		// try to detect english language by first char: too dirty but simple
		// 0x0 - English or other latin, 0x4 - Russian
		unsigned char* sData = (unsigned char*)word.GetString();
		if (sData[1] != 0x4)
		{
			// Latin detected
			detectLang = "en";
		}
	}
	return detectLang;
}

/// <summary>Get suggestion array for word</summary>
/// <param name="word">Word</param>  
/// <returns>Suggestion array</returns> 
CStrings* CSpeller::GetSuggestions(CString word)
{
	// remove all soft hyphens
	word.Replace(L"\u00AD", L"");

	CStrings* suggestions;
	suggestions = new CStrings();
	CString lang = detectWordLanguage(word);
	CW2A str(word, CP_UTF8);
	for (int i = 0; i < m_Dictionaries.GetSize(); i++)
	{
		// get suggestions list from all dictionarues
		if (m_Dictionaries[i].lang == lang && m_Dictionaries[i].handle)
		{
			char **list;
			int listLength = Hunspell_suggest(m_Dictionaries[i].handle, &list, str);
			suggestions = new CStrings();

			char** p = list;
			for (int j = 0; j < listLength; j++)
			{
				CString s(CA2CT(*p, CP_UTF8));
				suggestions->Add(s);
				p++;
			}
			Hunspell_free_list(m_Dictionaries[i].handle, &list, listLength);
		}
	}
	return suggestions;
}

/// <summary>Spell check the word</summary>
/// <param name="word">Word in UTF encoding</param>  
/// <returns>true-if spellcheck Ok, false in misspell</returns> 
bool CSpeller::SpellCheck(CString word)
{
	if (word.IsEmpty())
		return true;

	// do spell check
	CString checkWord(word);
	//replace last aphostrophe
	if (CString(L"\u0027\u2019\u0301").Find(word[word.GetLength() - 1]) > -1)
		checkWord.Delete(word.GetLength() - 1);

	// replace aphostrophes (dictionaries understand only regular ' aphostrophe
	m_numAphChanged = checkWord.Replace(L"’", L"'");
	// remove all soft hyphens
	checkWord.Replace(L"\u00AD", L"");
	// remove accent
	checkWord.Replace(L"\u0301", L"");

	CString lang = detectWordLanguage(word);

	// special case for Russian letter "ё"
	if (lang == "ru")
		checkWord.Replace(L"ё", L"е");

	bool spellResult = true;
	// encode string to the dictionary encoding 
	CW2A str(checkWord, CP_UTF8);
	int i = 0;
	for (i = 0; i < m_Dictionaries.GetSize(); i++)
	{
		if (m_Dictionaries[i].lang == lang && m_Dictionaries[i].handle)
		{
			spellResult = (bool)Hunspell_spell(m_Dictionaries[i].handle, str);
			if (spellResult)
				return true;
		}
	}

	if (spellResult)
		return true; // no suitable dictionaries at all

	// check ignore_all list first
	if (m_IgnoreWords.Find(word) > -1)
		return true;

	// check in custom dictionary
	if (m_CustomDict.Find(word) > -1)
		return true;

	return spellResult;
}

/// <summary>Highlight word at the pos in the element</summary>
/// <param name="elem">HTML Element</param>  
/// <param name="uniqID">UniqID of the element</param>  
/// <param name="pos">Position of word in element</param>  
void CSpeller::MarkWord(MSHTML::IHTMLElementPtr elem, long uniqID, CString word, int pos)
{
	MSHTML::IMarkupPointerPtr impStart;
	MSHTML::IMarkupPointerPtr impEnd;

	// Create start markup pointer
	m_ims->CreateMarkupPointer(&impStart);
	impStart->MoveAdjacentToElement(elem, MSHTML::ELEM_ADJ_AfterBegin); //after openeing tag
	//shift start markup pointer forward
	for (int i = 0; i < pos; i++)
		impStart->MoveUnit(MSHTML::MOVEUNIT_NEXTCHAR);

	// Create end markup pointer
	m_ims->CreateMarkupPointer(&impEnd);
	impEnd->MoveAdjacentToElement(elem, MSHTML::ELEM_ADJ_BeforeEnd);

	// Locate the misspelled word
	// Update impEnd
	if (impStart->findText(bstr_t(word), FINDTEXT_MATCHCASE, impEnd, NULL) == S_FALSE)
		return;

	// Create a display pointers from the markup pointers
	MSHTML::IDisplayPointerPtr idpStart;
	m_ids->CreateDisplayPointer(&idpStart);
	idpStart->MoveToMarkupPointer(impStart, NULL);

	MSHTML::IDisplayPointerPtr idpEnd;
	m_ids->CreateDisplayPointer(&idpEnd);
	idpEnd->MoveToMarkupPointer(impEnd, NULL);

	// Add or remove the segment
	MSHTML::IHighlightSegmentPtr ihs;
	m_ihrs->AddSegment(idpStart, idpEnd, m_irs, &ihs);

	m_ElementHighlights.insert(HIGHLIGHT(uniqID, ihs));
}

/// <summary>Removes all highlights in the document and clear saved highlights</summary>
void CSpeller::ClearAllMarks()
{
	for each (HIGHLIGHT itr in m_ElementHighlights)
	{
		m_ihrs->RemoveSegment(itr.second);
	}
	m_ElementHighlights.clear();
}

/// <summary>Removes highlights for element</summary>
/// <param name="elemID">Element ID</param>  
void CSpeller::ClearMark(int elemID)
{
	HIGHLIGHTS::iterator itr, lastElement;

	itr = m_ElementHighlights.find(elemID);
	if (itr != m_ElementHighlights.end())
	{
		lastElement = m_ElementHighlights.upper_bound(elemID);
		for (; itr != lastElement; ++itr)
			if (itr->second)
				m_ihrs->RemoveSegment(itr->second);

		m_ElementHighlights.erase(elemID);
	}
}

/// <summary>Spellcheck element</summary>
/// <param name="elem">HTML Element</param>  
/// <param name="HTMLChanged">true if DOM was changed</param>  
void CSpeller::CheckElement(MSHTML::IHTMLElementPtr elem, bool HTMLChanged)
{
	// skip whole document checking
	CString html = elem->innerHTML;
	if (html.Find(L"<DIV") >= 0 || html.Find(L"<div") >= 0)
		return;

	if (HTMLChanged) {
		// recheck whole page
		ClearAllMarks();
		doSpellCheck();
		return;
	}
	long uniqID = MSHTML::IHTMLUniqueNamePtr(elem)->uniqueNumber;

	MarkElement(elem, uniqID);
}
/// <summary>Spellcheck HTML Element</summary>
/// <param name="elem">HTML Element</param>  
/// <param name="uniqID">UniqID of the element</param>  
void CSpeller::MarkElement(MSHTML::IHTMLElementPtr elem, long uniqID)
{
	CString src = elem->innerText;
	if (!src.Trim().IsEmpty())
	{
		// remove underline
		ClearMark(uniqID);

		// split string to words and check words
		int wordStartPos = 0;
		int strlen = src.GetLength();
		for (int i = 0; i < strlen; i++)
		{
			if (!iswdigit(src.GetAt(i)) && !iswalpha(src.GetAt(i)))
			{
				if (i - wordStartPos) {
					CString wrd = src.Mid(wordStartPos, i - wordStartPos);
					if (!SpellCheck(wrd))
						MarkWord(elem, uniqID, wrd, wordStartPos);
				}
				wordStartPos = i + 1;
			}

			// add last word
			else if (i + 1 == strlen) {
				CString wrd = src.Mid(wordStartPos, strlen - wordStartPos);
				if (!SpellCheck(wrd))
					MarkWord(elem, uniqID, wrd, wordStartPos);
			}
		}
	}
}


/// <summary>If scroll changed do spellcheck (run OnIdle)</summary>
void CSpeller::CheckScroll()
{
	if (m_scrollElement)
	{
		long Y = m_scrollElement->scrollTop;
		if (Y != m_prevY)
		{
			doSpellCheck();
			m_prevY = Y;
		}
	}
}

/// <summary>Spellcheck whole visible page</summary>
void CSpeller::doSpellCheck()
{
	if (!m_Enabled)
		return;
	CString tagName;
	std::pair< std::set<long>::iterator, bool > pr;
	int currNum, numChanges = 0, nStartElem, nEndElem;
	MSHTML::IHTMLElementPtr elem, endElem;

	// lookup first element on page
	for (int y = 10; y < m_scrollElement->clientHeight; y += 10)
	{
		elem = m_doc2->elementFromPoint(63, y);
		tagName.SetString(elem->tagName);
		if (tagName.CompareNoCase(L"P") == 0)
			break;
	}
	// lookup last element on page
	for (int y = m_scrollElement->clientHeight - 1; y > 10; y -= 10)
	{
		endElem = m_doc2->elementFromPoint(63, y);
		tagName.SetString(endElem->tagName);
		if (tagName.CompareNoCase(L"P") == 0)
			break;
	}

	// get all document paragraphs
	MSHTML::IHTMLElementPtr fbw_body = MSHTML::IHTMLDocument3Ptr(m_doc2)->getElementById(L"fbw_body");
	MSHTML::IHTMLElementCollectionPtr paras = MSHTML::IHTMLElement2Ptr(fbw_body)->getElementsByTagName(L"P");

	nStartElem = nEndElem = -1;
	for (int i = 0; i < paras->length; i++)
	{
		if ((nStartElem == -1) && paras->item(i) == elem)
			nStartElem = i;
		else if (paras->item(i) == endElem)
		{
			nEndElem = i;
			break;
		}
	}

	if (nStartElem == -1)
		nStartElem = 0;

	if (nEndElem == -1)
		nEndElem = nStartElem + 20; //??
	else if (nEndElem + 1 < paras->length)
		nEndElem++;

	for (int i = nStartElem; i < nEndElem; i++)
	{
		elem = paras->item(i);
		if (elem)
		{
			// get element unique number
			int vv = MSHTML::IMarkupContainer2Ptr(m_doc2)->GetVersionNumber();
			currNum = MSHTML::IHTMLUniqueNamePtr(elem)->uniqueNumber;

			// Getting uniqueNumber from IHTMLUniqueName interface changes
			// the internal HTML document version due to generation of UniqueNumber. We need to correct this
			// issue, because document not really changed
			vv = MSHTML::IMarkupContainer2Ptr(m_doc2)->GetVersionNumber();
			pr = m_uniqIDs.insert(currNum);
			// increase changes only if new unique element inserted 
			if (pr.second) 
				numChanges++;

			MarkElement(elem, currNum);
		}
	}
	// Increase version number
	if (numChanges)
		::PostMessage(m_frame, WM_COMMAND, MAKELONG(ID_VER_ADVANCE, numChanges), 0);
}


void CSpeller::StartDocumentCheck(MSHTML::IMarkupServices2Ptr undoSrv)
{
	// create and show spell dialog
	if (!m_spell_dlg)
	{
		m_spell_dlg = new CSpellDialog(this);
		m_spell_dlg->ShowDialog();
		m_undoSrv = undoSrv;
	}

	// save current selection
	if (!m_prevSelRange)
	{
		m_ims->CreateMarkupPointer(&m_impStart);
		m_ims->CreateMarkupPointer(&m_impEnd);

		m_prevSelRange = m_doc2->selection->createRange();
		m_ims->MovePointersToRange(m_prevSelRange, m_impStart, m_impEnd);
	}

	// fetch selection
	CString selType = m_doc2->selection->type;

	m_selRange = m_doc2->selection->createRange();

	m_selRange->collapse(true);
	m_selRange->expand("word");
	CString sText = m_selRange->text;
	m_selRange->moveStart(L"word", -1);
	m_selRange->moveEnd(L"word", -1);
	sText.SetString(m_selRange->text);

	ContinueDocumentCheck();
}

//
// Spellcheck whole document from beginning
// Returns true if some changes was made, or false if no changes
//
void CSpeller::ContinueDocumentCheck()
{
	bool result;
	CString word;
	CString tag;
	_bstr_t b;

	// find next misspell word from the beginning or current position
	do
	{
		// shift to the next word 
		m_selRange->moveStart(L"word", 1);
		m_selRange->moveEnd(L"word", 1);
		MSHTML::IHTMLElementPtr elem(m_selRange->parentElement());
		tag.SetString(elem->tagName);

		word.SetString(m_selRange->text);
		word.Trim();

		// if word != words delimiter
		if (word.IsEmpty() || iswpunct(word.GetAt(0)))
		{
			result = true;
			continue;
			//_word.Rightword.FindOneOf(Tokens) == -1)
		}
		else
			result = SpellCheck(word);

		// select misspelled word
		if (!result) {
			m_selRange->select();
			m_spell_dlg->m_sBadWord = word;
			if (m_spell_dlg->m_strSuggestions)
				delete m_spell_dlg->m_strSuggestions;
			m_spell_dlg->m_strSuggestions = GetSuggestions(word);
			m_spell_dlg->UpdateData();
		}
	} while (result && tag != L"BODY"); // check for end of selection

	if (tag == L"BODY")
		EndDocumentCheck(false);
}

/// <summary>Finish document spellcheck</summary>
/// <param name="bCancel">User press Cancel button</param>  
void CSpeller::EndDocumentCheck(bool bCancel)
{
	if (m_spell_dlg)
	{
		m_spell_dlg->DestroyWindow();
		m_spell_dlg = nullptr;
	}
	// display message box
	if (!bCancel)
	{
		AtlTaskDialog(::GetActiveWindow(), IDR_MAINFRAME, IDS_SPELL_CHECK_COMPLETED, (LPCTSTR)NULL, TDCBF_OK_BUTTON, TD_INFORMATION_ICON);
	}
	// restore previous selection
	if (m_prevSelRange)
	{
		m_ims->MoveRangeToPointers(m_impStart, m_impEnd, m_prevSelRange);
		m_prevSelRange->select();

		// delete objects
		m_prevSelRange = nullptr;
	}
	// delete spell-check selection range
	if (m_selRange)
	{
		m_selRange = nullptr;
	}
}
