#pragma once

typedef bool (*FN_IsMe) (const MSHTML::IHTMLElementPtr elem);
typedef CString (*FN_GetTitle) (const MSHTML::IHTMLElementPtr elem);

const int IMG_BITMAP_TYPE = 0;
const int IMG_ICON_TYPE = 1;

class CElementDescriptor
{
private:
	FN_IsMe m_isMe;         // function to check type of element
	FN_GetTitle m_getTitle; // function to check title of element
	
    CString m_caption;      // display name in TreeView
	int m_imageID;          // Image ID
	bool m_viewInTree;      // view state in TreeView

	// for script descriptor type
	CString m_script_path;  // path to scriptfile
	CString m_class_name;   // classname in script
	bool m_script;          // script flag for descriptor
    HANDLE m_picture;       // image handle
	int m_pictType;         // image type
	bool m_title_from_script; //extract title from script

public:
	CElementDescriptor() :  m_isMe(nullptr), m_getTitle(nullptr), m_caption(L""), m_imageID(0), 
                            m_viewInTree(false), m_script_path(L""), m_class_name(L""), m_script(false),
                            m_picture(0), m_pictType(IMG_BITMAP_TYPE), m_title_from_script(false)
                            {};

	void Init(FN_IsMe fn1, FN_GetTitle fn2, int imageID, bool viewInTree_default, CString caption = L"");
	void CleanUp();

	virtual bool IsMe(const MSHTML::IHTMLElementPtr elem);

	CString GetTitle(const MSHTML::IHTMLElementPtr elem);
	CString GetCaption();
	bool GetPic(HANDLE& handle, int& type);
	int GetDTImageID();
	void SetImageID(int ID);	
	void SetViewInTree(bool view);

	bool ViewInTree();

	// working with script
	bool Load(CString path);
	void ProcessScript();
	
protected:
	// working with script
	// CString GetClassName();
	CString AskClassName();
};

class CElementDescMnr
{
    // like Singleton pattern
private:
	CSimpleArray<CElementDescriptor*> m_stEDs; // array of standard elements
	CSimpleArray<CElementDescriptor*> m_EDs;    // array of scripts elements
public:
	CElementDescMnr();

	void InitStandartEDs();
	void InitScriptEDs();
	void CleanUpAll();
	void CleanTree();
	
	CElementDescriptor* GetStED(int index);
	CElementDescriptor* GetED(int index);
    int GetStEDsCount();
	int GetEDsCount();
	bool GetElementDescriptor(const MSHTML::IHTMLElementPtr elem, CElementDescriptor **desc)const;
};
