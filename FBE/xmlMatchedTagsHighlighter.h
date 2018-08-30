// This file is part of Notepad++ project
// Copyright (C)2003 Don HO <don.h@free.fr>
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either
// version 2 of the License, or (at your option) any later version.
//
// Note that the GPL places important restrictions on "derived works", yet
// it does not provide a detailed definition of that term.  To avoid      
// misunderstandings, we consider an application to constitute a          
// "derivative work" for the purpose of this license if it does any of the
// following:                                                             
// 1. Integrates source code from Notepad++.
// 2. Integrates/includes/aggregates Notepad++ into a proprietary executable
//    installer, such as those produced by InstallShield.
// 3. Links to a library or executes a program that does any of the above.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

#ifndef XMLMATCHEDTAGSHIGHLIGHTER_H
#define XMLMATCHEDTAGSHIGHLIGHTER_H

#pragma once

using namespace std;

// wrapper class
class ScintillaEditView {
public:
	ScintillaEditView(CWindow* source): m_source(source){};

	LRESULT execute (UINT Msg, WPARAM wParam=NULL, LPARAM lParam=NULL) {
	    return m_source->SendMessage(Msg, wParam, lParam); 
	}
	
    int getCurrentDocLen() {
        return int(execute(SCI_GETLENGTH));
    };

	void clearIndicator(int indicatorNumber) {
		int docStart = 0;
		int docEnd = getCurrentDocLen();
		execute(SCI_SETINDICATORCURRENT, indicatorNumber);
		execute(SCI_INDICATORCLEARRANGE, docStart, docEnd-docStart);
	};

	void getText(char *dest, int start, int end) {
		Sci_TextRange tr;
		tr.chrg.cpMin = start;
		tr.chrg.cpMax = end;
		tr.lpstrText = dest;
		execute(SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
	}

	bool isShownIndentGuide()const {
		return false;
	}

private:
	CWindow* m_source;
};

//enum TagCateg {tagOpen, tagClose, inSingleTag, outOfTag, invalidTag, unknownPb};

typedef pair<CString, int> TAG;
typedef vector<TAG>::iterator tagIterator;


class XmlMatchedTagsHighlighter {
public:
	XmlMatchedTagsHighlighter(CWindow* source) {
	  _pEditView = new ScintillaEditView(source);
	};
	bool tagMatch(bool doHilite, bool doHiliteAttr, bool gotoTag);
	
private:
	struct XmlMatchedTagsPos {
		int tagOpenStart;
		int tagNameEnd;
		int tagOpenEnd;

		int tagCloseStart;
		int tagCloseEnd;
	};
	
	struct FindResult {
		int start;
		int end;
		bool success;
	};

	bool getXmlMatchedTagsPos(XmlMatchedTagsPos & tagsPos);

	// Allowed whitespace characters in XML
	bool isWhitespace(int ch) { return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n'; }

	FindResult findText(const char *text, int start, int end, int flags = 0);
	FindResult findOpenTag(const std::string& tagName, int start, int end);
	FindResult findCloseTag(const std::string& tagName, int start, int end);
	int findCloseAngle(int startPosition, int endPosition);

	ScintillaEditView *_pEditView;

	vector< pair<int, int> > getAttributesPos(int start, int end);
};

#endif //XMLMATCHEDTAGSHIGHLIGHTER_H
