#pragma once

#include <algorithm>
#include <map>
#include <vector>

#include "resource.h"
#include "res1.h"
#include "utils.h"
#include "extras\Serializable.h"

/*const int ILANG_ENGLISH = 0;
const int ILANG_RUSSIAN = 1;*/

class WordsItem : public ISerializable, public IObjectFactory
{
public:
	CString	m_word;

	int		m_count;
	CString	m_sCount;

	int		m_percent;
	int		m_prc_idx;

	WordsItem() { }
	WordsItem(CString word, int count) : m_word(word), m_count(count) { }

	int GetProperties(std::vector<CString>& properties)
	{
		properties.push_back(L"Value");
		properties.push_back(L"Counted");
		return properties.size();
	}

	bool GetPropertyValue(const CString& sProperty, CProperty& property)
	{
		if(sProperty == L"Value")
		{
			property = m_word;
			return true;
		}
		else if(sProperty == L"Counted")
		{
			CString counted;
			counted.Format(L"%d", m_count);
			property = counted;
			return true;
		}
		return false;
	}

	bool SetPropertyValue(const CString& sProperty, CProperty& property)
	{
		if(sProperty == L"Value")
		{
			m_word = property;
			return true;
		}
		else if(sProperty == L"Counted")
		{
			m_count = StrToInt(property.GetStringValue());
			return true;
		}
		return false;
	}

	bool HasMultipleInstances()
	{
		return true;
	}

	CString GetClassName()
	{
		return L"Word";
	}


	CString GetID()
	{
		return L"";
	}

	ISerializable* Create()
	{
		return new WordsItem;
	}

	void Destroy(ISerializable* obj)
	{
		delete obj;
	}

	bool operator==(const WordsItem& wi)
	{
		return m_word.CompareNoCase(wi.m_word) == 0;
	}
};

class CHotkey : public ISerializable, public IObjectFactory
{
public:
	CString m_name;
	CString m_reg_name;
	ACCEL m_accel;
	ACCEL m_def_accel;
	CString m_desc;
	wchar_t m_char_val; // value for symbol hotkey

	CHotkey() {}

	CHotkey(CString reg_name, int IDS_CMD_NAME, BYTE fVirt, WORD cmd, WORD key, CString descr = L"")
	{
		m_reg_name = reg_name;
		m_name.LoadString(IDS_CMD_NAME);

		m_def_accel.fVirt = FVIRTKEY | fVirt;
		m_def_accel.cmd = cmd;
		m_def_accel.key = key;

		m_accel = m_def_accel;

		m_desc = descr;
	}

	CHotkey(CString reg_name, int IDS_CMD_NAME, CString uchar, BYTE fVirt, WORD cmd, WORD key, CString descr = L"")
	{
		m_reg_name = reg_name;
		m_name.LoadString(IDS_CMD_NAME);
		m_name += uchar;

		m_def_accel.fVirt = FVIRTKEY | fVirt;
		m_def_accel.cmd = cmd;
		m_def_accel.key = key;

		m_accel = m_def_accel;

		m_desc = descr;
	}

	CHotkey(CString reg_name, CString name, wchar_t symbol, BYTE fVirt, WORD cmd, WORD key, CString descr = L"")
	{
		m_reg_name = reg_name;
		m_name = name;

		m_char_val = symbol;

		m_def_accel.fVirt = FVIRTKEY | fVirt;
		m_def_accel.cmd = cmd;
		m_def_accel.key = key;

		m_accel = m_def_accel;

		m_desc = descr;
	}

	CHotkey(CString reg_name, CString cmd_name, BYTE fVirt, WORD cmd, WORD key, CString descr = L"")
	{
		m_reg_name = reg_name;
		m_name = cmd_name;

		m_def_accel.fVirt = FVIRTKEY | fVirt;
		m_def_accel.cmd = cmd;
		m_def_accel.key = key;

		m_accel = m_def_accel;

		m_desc = descr;
	}

	bool operator < (const CHotkey& other) const
	{
		return (m_name.CompareNoCase(other.m_name) < 0);
	}

	// ISerializable interface
	int GetProperties(std::vector<CString>& properties)
	{
		properties.push_back(L"Name");
		properties.push_back(L"Accel");

		return properties.size();
	}

	bool GetPropertyValue(const CString& sProperty, CProperty& property)
	{
		if(sProperty == L"Name")
		{
			property = m_reg_name;
			return true;
		}
		if(sProperty == L"Accel")
		{
			CString temp;
			temp.Format(L"%u;%u",
				m_accel.fVirt, m_accel.key);
			property = temp;
			return true;
		}

		return false;
	}

	bool SetPropertyValue(const CString& sProperty, CProperty& sValue)
	{
		if(sProperty == L"Name")
		{
			m_reg_name = sValue.GetStringValue();
			return true;
		}
		else if(sProperty == L"Accel")
		{
			CString str = sValue.GetStringValue();
			int n = 0, curPos = 0;

			while(str.Tokenize(L";", curPos) != L"")
				n++;

			CString* tokens = new CString[n];
			curPos = n =0;

			CString temp;
			while((temp = str.Tokenize(L";", curPos)) != L"")
			{
				tokens[n] = temp;
				n++;
			}

			if(n == 2)
			{
				m_accel.fVirt = BYTE (StrToInt(tokens[0]));
				m_accel.key = WORD(StrToInt(tokens[1]));
			}

			delete[] tokens;

			return true;
		}

		return false;
	}

	bool HasMultipleInstances()
	{
		return false;
	}

	CString GetClassName()
	{
		return L"Hotkey";
	}

	CString GetID()
	{
		return L"";
	}

	ISerializable* Create()
	{
		return new CHotkey;
	}

	void Destroy(ISerializable* obj)
	{
		delete obj;
	}
};

class CHotkeysGroup : public ISerializable, public IObjectFactory
{
public:
	CString m_name;
	CString m_reg_name;
	std::vector<CHotkey> m_hotkeys;

	CHotkey m_hotkey_factory;
	std::vector<void*> m_ptr_hotkeys;

	CHotkeysGroup()
	{
	}

	CHotkeysGroup(CString reg_name, int IDS_GROUP_NAME)
	{
		m_reg_name = reg_name;
		::LoadString(_Module.GetResourceInstance(), IDS_GROUP_NAME, m_name.GetBufferSetLength(MAX_LOAD_STRING + 1), 
			MAX_LOAD_STRING + 1);
	}

	bool operator < (const CHotkeysGroup& other) const
	{
		return (m_name.CompareNoCase(other.m_name) < 0);
	}

	int GetProperties(std::vector<CString>& properties)
	{
		properties.push_back(L"GroupName");
		properties.push_back(L"Hotkey");

		return properties.size();
	}

	bool GetPropertyValue(const CString& sProperty, CProperty& property)
	{
		if(sProperty == L"GroupName")
		{
			property = m_reg_name;
			return true;
		}
		if(sProperty == L"Hotkey")
		{
			for(unsigned long i = 0 ; i < m_hotkeys.size(); ++i)
				m_ptr_hotkeys.push_back(&m_hotkeys[i]);
			
			property = m_ptr_hotkeys;
			property.SetFactory(&m_hotkey_factory);

			return true;
		}

		return false;
	}

	bool SetPropertyValue(const CString& sProperty, CProperty& sValue)
	{
		if(sProperty == L"GroupName")
		{
			m_reg_name = sValue.GetStringValue();
			return true;
		}
		if(sProperty == L"Hotkey")
		{
			std::vector<void*>::iterator iter = m_ptr_hotkeys.begin();

			while(iter != m_ptr_hotkeys.end())
			{
				CHotkey* pHotkey = (CHotkey*)&iter;
				delete pHotkey;
				iter++;
			}

			CProperty::CopyPtrList(m_ptr_hotkeys, sValue.GetObjectList());

			m_hotkeys.clear();
			for(unsigned long i = 0 ; i < m_ptr_hotkeys.size(); ++i)
			{
				CHotkey* temp = (CHotkey*)m_ptr_hotkeys[i];
				m_hotkeys.push_back(*temp);
			}

			return true;
		}

		return false;
	}

	bool HasMultipleInstances()
	{
		return true;
	}

	CString GetClassName()
	{
		return L"HkGroup";
	}

	CString GetID()
	{
		return L"";
	}

	ISerializable* Create()
	{
		return new CHotkeysGroup;
	}

	void Destroy(ISerializable* obj)
	{
		delete obj;
	}
};

class DESCSHOWINFO : public ISerializable, public IObjectFactory
{
public:
	std::map<CString, bool> elements;

	DESCSHOWINFO();

	// Default fields showing in description
	void SetDefaults();

	// ISerializable interface
	int GetProperties(std::vector<CString>& properties);
	bool GetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool SetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool HasMultipleInstances();
	CString GetClassName();
	CString GetID();

	// IObjectFactory
	ISerializable* Create();
	void Destroy(ISerializable*);
};

class TREEITEMSHOWINFO : public ISerializable, public IObjectFactory
{
public:
	std::map<CString, bool> items;

	TREEITEMSHOWINFO();
//	void SetDefaults();

	// ISerializable interface
	int GetProperties(std::vector<CString>& properties);
	bool GetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool SetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool HasMultipleInstances();
	CString GetClassName();
	CString GetID();

	// IObjectFactory
	ISerializable* Create();
	void Destroy(ISerializable*);
};

class CSettings : public ISerializable, public IObjectFactory
{
	enum AppPropertyType {
		tUnknown = 0,
		tBOOL,
		tSTRING,
		tDWORD,
		tWINDOWPOS
	};
	struct Property {
		CString			PropValue;
		AppPropertyType PropType;
		Property(CString s, AppPropertyType ap) 
		{ 
			PropValue = s;
			PropType = ap;
		};
		Property(int s)
		{
			PropValue = L"";
			PropType = tUnknown;
		};
	};
	CSimpleMap<CString, Property> m_AppProperty;

	void AppPropertyInit()
	{
		m_AppProperty.RemoveAll();
		m_AppProperty.Add(L"KeepEncoding", { L"true", tBOOL });
		m_AppProperty.Add(L"DefaultSaveEncoding", { L"utf-8", tSTRING });
		m_AppProperty.Add(L"SearchOptions", { L"0", tDWORD });
		m_AppProperty.Add(L"ColorBG", { L"4278190080", tDWORD });
		m_AppProperty.Add(L"ColorFG", { L"4278190080", tDWORD });
		m_AppProperty.Add(L"XMLSrcWrap", { L"true", tBOOL });
		m_AppProperty.Add(L"XMLSrcSyntaxHL", { L"true", tBOOL });
		m_AppProperty.Add(L"XMLSrcTagHL", { L"true", tBOOL });
		m_AppProperty.Add(L"XMLSrcShowEOL", { L"false", tBOOL });
		m_AppProperty.Add(L"XMLSrcShowSpace", { L"false", tBOOL });
		m_AppProperty.Add(L"XMLSrcShowLineNumbers", { L"false", tBOOL });
		m_AppProperty.Add(L"FastMode", { L"false", tBOOL });

		m_AppProperty.Add(L"FontSize", { L"12", tDWORD });
		m_AppProperty.Add(L"Font", { L"Trebuchet MS", tSTRING });
		m_AppProperty.Add(L"SrcFont", { L"Lucida Console", tSTRING });
		m_AppProperty.Add(L"ViewStatusBar", { L"true", tBOOL });
		m_AppProperty.Add(L"ViewDocumentTree", { L"true", tBOOL });
		m_AppProperty.Add(L"SplitterPos", { L"275", tDWORD });
		m_AppProperty.Add(L"Toolbars", { L"", tSTRING });
		m_AppProperty.Add(L"RestoreFilePosition", { L"true", tBOOL });
		m_AppProperty.Add(L"IntefaceLangID", { L"25", tDWORD });
		m_AppProperty.Add(L"ScriptsFolder", { L"Scripts", tSTRING });
		m_AppProperty.Add(L"UseSpellChecker", { L"true", tBOOL });
		m_AppProperty.Add(L"HighlightMisspells", { L"true", tBOOL });
		m_AppProperty.Add(L"CustomDict", { L"custom.dic", tSTRING });
		m_AppProperty.Add(L"CustomDictCodePage", { L"65001", tDWORD });
		m_AppProperty.Add(L"NBSPChar", { L" ", tSTRING });
		m_AppProperty.Add(L"ChangeKeybLayout", { L"false", tBOOL });
		m_AppProperty.Add(L"KeyboardLayout", { L"1033", tDWORD });
		m_AppProperty.Add(L"PasteImageType", { L"1", tDWORD });
		m_AppProperty.Add(L"JpegQuality", { L"75", tDWORD });
		m_AppProperty.Add(L"InsImageDialog", { L"false", tBOOL });
		m_AppProperty.Add(L"InsClearImage", { L"false", tBOOL });
		m_AppProperty.Add(L"WindowPosition", { L"44;0;1;-1;-1;-1;-1;1433;26;717;1526", tWINDOWPOS });
		m_AppProperty.Add(L"WordsDlgPosition", { L"44;0;1;-1;-1;-1;-1;838;1052;235;1709", tWINDOWPOS });
		m_AppProperty.Add(L"ShowWordsExclusions", { L"true", tBOOL });
	}

	bool getPropValueAsBool(const CString& name)
	{
		Property  pr = m_AppProperty.Lookup(name);
		return pr.PropValue == L"true";
	}

	CString getPropValueAsString(const CString& name)
	{
		Property  pr = m_AppProperty.Lookup(name);
		return pr.PropValue;
	}

	DWORD getPropValueAsDWORD(const CString& name)
	{
		Property pr = m_AppProperty.Lookup(name);
		return _wtol(pr.PropValue);
	}

	void getPropValueAsWINDOWPLACEMENT(const CString& name, WINDOWPLACEMENT &wpl)
	{
		Property pr = m_AppProperty.Lookup(name);
		CString str = pr.PropValue;
		int n = 0, curPos = 0;

		while (str.Tokenize(L";", curPos) != L"")
			n++;

		CString* tokens = new CString[n];
		curPos = n = 0;

		CString temp;
		while ((temp = str.Tokenize(L";", curPos)) != L"")
		{
			tokens[n] = temp;
			n++;
		}

		if (n == 11)
		{
			wpl.length = StrToInt(tokens[0]);
			wpl.flags = StrToInt(tokens[1]);
			wpl.showCmd = StrToInt(tokens[2]);
			wpl.ptMinPosition.x = StrToInt(tokens[3]);
			wpl.ptMinPosition.y = StrToInt(tokens[4]);
			wpl.ptMaxPosition.x = StrToInt(tokens[5]);
			wpl.ptMaxPosition.y = StrToInt(tokens[6]);
			wpl.rcNormalPosition.bottom = StrToInt(tokens[7]);
			wpl.rcNormalPosition.left = StrToInt(tokens[8]);
			wpl.rcNormalPosition.top = StrToInt(tokens[9]);
			wpl.rcNormalPosition.right = StrToInt(tokens[10]);
		}

		delete[] tokens;
	}

	bool setPropValueAsBool(const CString& name, bool Value)
	{
		Property p = { Value? L"true": L"false", tBOOL };
		return m_AppProperty.SetAt(name, p);
	};

	bool setPropValueAsString(const CString& name, CString& Value)
	{
		Property p = { Value, tSTRING };
		return m_AppProperty.SetAt(name, p);
	}

	bool setPropValueAsDWORD(const CString& name, DWORD Value)
	{
		CString str(L"");
		str.Format(L"%u", Value);
		Property p = { str, tDWORD };
		return m_AppProperty.SetAt(name, p);
	}

	bool setPropValueAsWINDOWPLACEMENT(const CString& name, WINDOWPLACEMENT &Value)
	{
		CString temp;
		temp.Format(L"%u;%u;%u;%ld;%ld;%ld;%ld;%ld;%ld;%ld;%ld",
			Value.length,
			Value.flags,
			Value.showCmd,
			Value.ptMinPosition.x,
			Value.ptMinPosition.y,
			Value.ptMaxPosition.x,
			Value.ptMaxPosition.y,
			Value.rcNormalPosition.bottom,
			Value.rcNormalPosition.left,
			Value.rcNormalPosition.top,
			Value.rcNormalPosition.right);
		Property p = { temp, tWINDOWPOS };
		return m_AppProperty.SetAt(name, p);
	}

	///<summary>Load properties from file</summary>
	void CSettings::LoadAppProperty()
	{
		/*CXMLAppSettings xmlStorage1(CIniAppSettings::GetRoamingDataFileLocation(L"FBE", L"FBEditor", L"", L"FBESettings.xml"), true, true);
		for (int i = 0; i<m_AppProperty.GetSize(); i++)
		{
			CString name = m_AppProperty.GetKeyAt(i);
			CString val = xmlStorage1.GetString("FBE", name);
			if (val)
				m_AppProperty.SetAtIndex(i, name, val);
		}*/
	}
	///<summary>Load properties from file</summary>
	void CSettings::SaveAppProperty()
	{
		/*CXMLAppSettings xmlStorage1(CIniAppSettings::GetRoamingDataFileLocation(L"FBE", L"FBEditor", L"", L"FBESettings.xml"), true, true);
		for (int i = 0; i<m_AppProperty.GetSize(); i++)
		{
			CString name = m_AppProperty.GetKeyAt(i);
			Property pr = m_AppProperty.GetValueAt(i);
			xmlStorage1.WriteString("FBE", name, pr.PropValue);
		}*/
	}

	CRegKey		m_key;
	CString		m_key_path;

	CString		m_default_encoding;

	DWORD		m_search_options;

	DWORD		m_collorBG;
	DWORD		m_collorFG;
	DWORD		m_font_size;
	CString		m_font;
	CString		m_srcfont;

	bool		m_view_status_bar;
	bool		m_view_doc_tree;

	// added by SeNS
	DWORD		m_custom_dict_codepage;
	CString		m_nbsp_char;
	CString		m_old_nbsp;
	DWORD		m_keyb_layout;
	DWORD		m_image_type;
	///

	DWORD		m_splitter_pos;
	CString		m_toolbars_settings;

	DWORD		m_interface_lang_id;

	bool		m_need_restart;

	CString		m_scripts_folder;

	bool		m_show_words_excls;

	WINDOWPLACEMENT m_words_dlg_placement;
	WINDOWPLACEMENT m_wnd_placement;

	DESCSHOWINFO m_desc;
	TREEITEMSHOWINFO m_tree_items;

	// added by SeNS: view dimensions for external helper
	int m_viewWidth, m_viewHeight;
	HWND m_hMainWindow;

public:
	std::vector<CHotkeysGroup> m_hotkey_groups;
	int keycodes; // total number of accelerators
	std::vector<WordsItem> m_words;

	bool m_xml_src_wrap;
	bool m_xml_src_syntaxHL;
	bool m_xml_src_showEOL;
	bool m_xml_src_showSpace;
	bool m_show_line_numbers;
	bool m_xml_src_tagHL;
	bool m_usespell_check;
	bool m_highlght_check;
	CString	m_custom_dict;
	bool m_fast_mode;

	bool m_keep_encoding; // save with opened encoding
	bool m_restore_file_position;
	bool m_insimage_ask;
	bool m_ins_clear_image;
	DWORD m_jpeg_quality;
	bool m_change_kbd_layout_check;
public:
	CSettings();
	~CSettings();

	void Init();
	void InitHotkeyGroups();
	void Close();

	// ISerializable interface
	int GetProperties(std::vector<CString>& properties);
	bool GetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool SetPropertyValue(const CString& sProperty, CProperty& sValue);
	bool HasMultipleInstances();
	CString GetClassName();
	CString GetID();

	// IObjectFactory
	ISerializable* Create();
	void Destroy(ISerializable*);
	void SetDefaults();

	void Load();
	void Save();

	void LoadHotkeyGroups();
	void SaveHotkeyGroups();

	CHotkeysGroup* GetGroupByName(const CString& name);
	CHotkey* GetHotkeyByName(const CString& name, CHotkeysGroup& group);

	void LoadWords();
	void SaveWords();

	bool NeedRestart()const;

	bool ViewStatusBar()const;
	bool ViewDocumentTree()const;

	CString GetKeyPath()const;

	CString GetDefaultEncoding()const;
	DWORD	GetSearchOptions()const;
	DWORD	GetColorBG()const;
	DWORD	GetColorFG()const;
	DWORD	GetFontSize()const;
	CString	GetFont()const;
	CString	GetSrcFont()const;
	DWORD	GetSplitterPos()const;	
	CString GetToolbarsSettings()const;
	const CRegKey& GetKey()const;

	// added by SeNS
	DWORD	GetCustomDictCodepage()const;
	CString GetNBSPChar()const;
	CString GetOldNBSPChar()const;
	DWORD	GetKeybLayout()const;
	DWORD	GetImageType()const;

	bool	GetExtElementStyle(const CString& elem)const;
	bool	GetWindowPosition(WINDOWPLACEMENT& wpl)const;
	DWORD	GetInterfaceLanguageID()const;
	CString GetInterfaceLanguageDllName()const;
	CString GetLocalizedGenresFileName()const;
	CString GetInterfaceLanguageName()const;
	CString GetScriptsFolder() const;
	CString GetDefaultScriptsFolder();
	bool	IsDefaultScriptsFolder();
	bool	GetShowWordsExcls()const;
	bool	GetWordsDlgPosition(WINDOWPLACEMENT &wpl)const;

	bool	GetDocTreeItemState(const CString& item, bool default_state);

	void	SetKeepEncoding(bool keep, bool apply = false);
	void	SetDefaultEncoding(const CString& encoding, bool apply = false);	
	void	SetSearchOptions(DWORD opt, bool apply = false);
	void	SetColorBG(DWORD color, bool apply = false);
	void	SetColorFG(DWORD color, bool apply = false);
	void	SetFontSize(DWORD size, bool apply = false);
	void	SetXmlSrcWrap(bool wrap, bool apply = false);
	void	SetXmlSrcSyntaxHL(bool hl, bool apply = false);
	void	SetXmlSrcTagHL(bool hl, bool apply = false);
	void	SetXmlSrcShowEOL(bool eol, bool apply = false);
	void	SetXmlSrcShowSpace(bool eol, bool apply = false);
	void	SetFastMode(bool mode,  bool apply = false);
	void	SetFont(const CString& font, bool apply = false);
	void	SetSrcFont(const CString& font, bool apply = false);
	void	SetViewStatusBar(bool view,  bool apply = false);
	void	SetViewDocumentTree(bool view,  bool apply = false);
	void	SetSplitterPos(DWORD pos,  bool apply = false);
	void	SetToolbarsSettings(CString& settings,  bool apply = false);
	void	SetExtElementStyle(const CString& elem, bool ext, bool apply = false);
	void	SetWindowPosition(const WINDOWPLACEMENT& wpl,  bool apply = false);
	void	SetRestoreFilePosition(bool restore, bool apply = false);	
	void	SetInterfaceLanguage(DWORD Language, bool apply = false);
	void	SetScriptsFolder(const CString fullpath, bool apply = false);
	void	SetInsImageAsking(const bool value, bool apply = false);
	void	SetIsInsClearImage(const bool value, bool apply = false);
	void	SetDocTreeItemState(const CString& item, bool state);
	void	SetShowWordsExcls(const bool value, bool apply = false);
	void	SetWordsDlgPosition(const WINDOWPLACEMENT& wpl,  bool apply = false);

	void	SetNeedRestart();

	CString	m_initial_scripts_folder;

	// added by SeNS
	void	SetUseSpellChecker(const bool value, bool apply = false);
	void	SetHighlightMisspells(const bool value, bool apply = false);
	void	SetCustomDict(const ATL::CString &value, bool apply = false);
	void	SetCustomDictCodepage(const DWORD value, bool apply = false);
	void	SetNBSPChar(const ATL::CString &value, bool apply = false);
	void	SetChangeKeybLayout(const bool value, bool apply = false);
	void	SetKeybLayout(const DWORD value, bool apply = false);
	void	SetXMLSrcShowLineNumbers(const bool value, bool apply = false);
	void	SetImageType(const DWORD value, bool apply = false);
	void	SetJpegQuality(const DWORD value, bool apply = false);

	void	SetViewWidth(int width) { m_viewWidth = width; }
	void	SetViewHeight(int height) { m_viewHeight = height; }
	int		GetViewWidth() { return m_viewWidth; }
	int		GetViewHeight() { return m_viewHeight; }
	void	SetMainWindow(HWND hwnd) { m_hMainWindow = hwnd; }
	HWND	GetMainWindow() { return m_hMainWindow; }

};