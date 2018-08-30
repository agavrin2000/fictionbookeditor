// This file is part of Notepad++ project
// Copyright (C)2012 Don HO <don.h@free.fr>
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


// Tags matching routing rewritten by Dave Brotherstone May 2012
// to remove need for regular expression searches (especially reverse regex searches)
// Reverse regex are slow using the new regex engine, and hence cost too much time.

// Add functions specific for FBEditor by agavrin August 2018
#include "stdafx.h"
#include "xmlMatchedTagsHighlighter.h"

/// <summary>Parse XML-element for attributes</summary>
/// <param name="start">Start position for search</param>  
/// <param name="end">End position for search</param>  
/// <returns>attributes Vector with start&end positions</returns> 
vector< pair<int, int> > XmlMatchedTagsHighlighter::getAttributesPos(int start, int end)
{
	vector< pair<int, int> > attributes;

	// allocate buffer fopr text range
	int bufLen = end - start + 1;
	char *buf = new char[bufLen + 1];
	_pEditView->getText(buf, start, end);

	enum {
		\
		attr_invalid, \
		attr_key, \
		attr_pre_assign, \
		attr_assign, \
		attr_string, \
		attr_value, \
		attr_valid\
	} state = attr_invalid;

	int startPos = -1;
	int oneMoreChar = 1;
	int i = 0;
	for (; i < bufLen; i++)
	{
		switch (buf[i])
		{
		case ' ':
		case '\t':
		case '\n':
		case '\r':
		{
			if (state == attr_key)
				state = attr_pre_assign;
			else if (state == attr_value)
			{
				state = attr_valid;
				oneMoreChar = 0;
			}
		}
		break;

		case '=':
		{
			if (state == attr_key || state == attr_pre_assign)
				state = attr_assign;
			else if (state == attr_assign || state == attr_value)
				state = attr_invalid;
		}
		break;

		case '"':
		{
			if (state == attr_string)
			{
				state = attr_valid;
				oneMoreChar = 1;
			}
			else if (state == attr_key || state == attr_pre_assign || state == attr_value)
				state = attr_invalid;
			else if (state == attr_assign)
				state = attr_string;
		}
		break;

		default:
		{
			if (state == attr_invalid)
			{
				state = attr_key;
				startPos = i;
			}
			else if (state == attr_pre_assign)
				state = attr_invalid;
			else if (state == attr_assign)
				state = attr_value;
		}
		}

		if (state == attr_valid)
		{
			attributes.push_back(pair<int, int>(start + startPos, start + i + oneMoreChar));
			state = attr_invalid;
		}
	}
	if (state == attr_value)
		attributes.push_back(pair<int, int>(start + startPos, start + i - 1));

	delete[] buf;
	return attributes;
}

/// <summary>Define XML-element around current carret position</summary>
/// <param name="&xmlTags">Result of search with open&closing tags posutions</param>  
/// <returns>true, if tag defined, or else - vice versa</returns> 
bool XmlMatchedTagsHighlighter::getXmlMatchedTagsPos(XmlMatchedTagsPos &xmlTags)
{
	bool tagFound = false;
	int caret = static_cast<int32_t>(_pEditView->execute(SCI_GETCURRENTPOS));
	int searchStartPoint = caret;
	LRESULT styleAt;
	FindResult openFound;

	// Search back for the previous open angle bracket.
	// Keep looking whilst the angle bracket found is inside an XML attribute
	do
	{
		openFound = findText("<", searchStartPoint, 0, 0);
		styleAt = _pEditView->execute(SCI_GETSTYLEAT, openFound.start);
		searchStartPoint = openFound.start - 1;
	} while (openFound.success && (styleAt == SCE_H_DOUBLESTRING || styleAt == SCE_H_SINGLESTRING || styleAt == SCE_H_COMMENT) && searchStartPoint > 0);

	if (openFound.success && styleAt != SCE_H_CDATA)
	{
		// Found the "<" before the caret, now check there isn't a > between that position and the caret.
		FindResult closeFound;
		searchStartPoint = openFound.start;
		do
		{
			closeFound = findText(">", searchStartPoint, caret, 0);
			styleAt = _pEditView->execute(SCI_GETSTYLEAT, closeFound.start);
			searchStartPoint = closeFound.end;
		} while (closeFound.success && (styleAt == SCE_H_DOUBLESTRING || styleAt == SCE_H_SINGLESTRING || styleAt == SCE_H_COMMENT) && searchStartPoint <= caret);

		if (!closeFound.success)
		{
			// We're in a tag (either a start tag or an end tag)
			int nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, openFound.start + 1));


			/////////////////////////////////////////////////////////////////////////
			// CURSOR IN CLOSE TAG   
			/////////////////////////////////////////////////////////////////////////
			if ('/' == nextChar)
			{
				xmlTags.tagCloseStart = openFound.start;
				int docLength = static_cast<int32_t>(_pEditView->execute(SCI_GETLENGTH));
				FindResult endCloseTag = findText(">", caret, docLength, 0);
				if (endCloseTag.success)
				{
					xmlTags.tagCloseEnd = endCloseTag.end;
				}
				// Now find the tagName
				int position = openFound.start + 2;

				// UTF-8 or ASCII tag name
				std::string tagName;
				nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, position));
				// Checking for " or ' is actually wrong here, but it means it works better with invalid XML
				while (position < docLength && !isWhitespace(nextChar) && nextChar != '/' && nextChar != '>' && nextChar != '\"' && nextChar != '\'')
				{
					tagName.push_back(static_cast<char>(nextChar));
					++position;
					nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, position));
				}

				// Now we know where the end of the tag is, and we know what the tag is called
				if (tagName.size() != 0)
				{
					/* Now we need to find the open tag.  The logic here is that we search for "<TAGNAME",
					* then check the next character - if it's one of '>', ' ', '\"' then we know we've found
					* a relevant tag.
					* We then need to check if either
					*    a) this tag is a self-closed tag - e.g. <TAGNAME attrib="value" />
					* or b) this tag has another closing tag after it and before our closing tag
					*       e.g.  <TAGNAME attrib="value">some text</TAGNAME></TAGNA|ME>
					*             (cursor represented by |)
					* If it's either of the above, then we continue searching, but only up to the
					* the point of the last find. (So in the (b) example above, we'd only search backwards
					* from the first "<TAGNAME...", as we know there's a close tag for the opened tag.

					* NOTE::  NEED TO CHECK THE ROTTEN CASE: ***********************************************************
					* <TAGNAME attrib="value"><TAGNAME>something</TAGNAME></TAGNAME></TAGNA|ME>
					* Maybe count all closing tags between start point and start of our end tag.???
					*/
					int currentEndPoint = xmlTags.tagCloseStart;
					int openTagsRemaining = 1;
					FindResult nextOpenTag;
					do
					{
						nextOpenTag = findOpenTag(tagName, currentEndPoint, 0);
						if (nextOpenTag.success)
						{
							--openTagsRemaining;
							// Open tag found
							// Now we need to check how many close tags there are between the open tag we just found,
							// and our close tag
							// eg. (Cursor == | )
							// <TAGNAME attrib="value"><TAGNAME>something</TAGNAME></TAGNAME></TAGNA|ME>
							//                         ^^^^^^^^ we've found this guy
							//                                           ^^^^^^^^^^ ^^^^^^^^ Now we need to cound these fellas
							FindResult inbetweenCloseTag;
							int currentStartPosition = nextOpenTag.end;
							int closeTagsFound = 0;
							bool forwardSearch = (currentStartPosition < currentEndPoint);

							do
							{
								inbetweenCloseTag = findCloseTag(tagName, currentStartPosition, currentEndPoint);

								if (inbetweenCloseTag.success)
								{
									++closeTagsFound;
									if (forwardSearch)
									{
										currentStartPosition = inbetweenCloseTag.end;
									}
									else
									{
										currentStartPosition = inbetweenCloseTag.start - 1;
									}
								}

							} while (inbetweenCloseTag.success);

							// If we didn't find any close tags between the open and our close,
							// and there's no open tags remaining to find
							// then the open we found was the right one, and we can return it
							if (0 == closeTagsFound && 0 == openTagsRemaining)
							{
								xmlTags.tagOpenStart = nextOpenTag.start;
								xmlTags.tagOpenEnd = nextOpenTag.end + 1;
								xmlTags.tagNameEnd = nextOpenTag.start + static_cast<int32_t>(tagName.size()) + 1;  /* + 1 to account for '<' */
								tagFound = true;
							}
							else
							{

								// Need to find the same number of opening tags, without closing tags etc.
								openTagsRemaining += closeTagsFound;
								currentEndPoint = nextOpenTag.start;
							}
						}
					} while (!tagFound && openTagsRemaining > 0 && nextOpenTag.success);
				}
			}
			else
			{
				/////////////////////////////////////////////////////////////////////////
				// CURSOR IN OPEN TAG   
				/////////////////////////////////////////////////////////////////////////
				int position = openFound.start + 1;
				int docLength = static_cast<int32_t>(_pEditView->execute(SCI_GETLENGTH));

				xmlTags.tagOpenStart = openFound.start;

				std::string tagName;
				nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, position));
				// Checking for " or ' is actually wrong here, but it means it works better with invalid XML
				while (position < docLength && !isWhitespace(nextChar) && nextChar != '/' && nextChar != '>' && nextChar != '\"' && nextChar != '\'')
				{
					tagName.push_back(static_cast<char>(nextChar));
					++position;
					nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, position));
				}

				// Now we know where the end of the tag is, and we know what the tag is called
				if (tagName.size() != 0)
				{
					// First we need to check if this is a self-closing tag.
					// If it is, then we can just return this tag to highlight itself.
					xmlTags.tagNameEnd = openFound.start + static_cast<int32_t>(tagName.size()) + 1;
					int closeAnglePosition = findCloseAngle(position, docLength);
					if (-1 != closeAnglePosition)
					{
						xmlTags.tagOpenEnd = closeAnglePosition + 1;
						// If it's a self closing tag
						if (_pEditView->execute(SCI_GETCHARAT, closeAnglePosition - 1) == '/')
						{
							// Set it as found, and mark that there's no close tag
							xmlTags.tagCloseEnd = -1;
							xmlTags.tagCloseStart = -1;
							tagFound = true;
						}
						else
						{
							// It's a normal open tag



							/* Now we need to find the close tag.  The logic here is that we search for "</TAGNAME",
							* then check the next character - if it's '>' or whitespace followed by '>' then we've
							* found a relevant tag.
							* We then need to check if
							* our tag has another opening tag after it and before the closing tag we've found
							*       e.g.  <TA|GNAME><TAGNAME attrib="value">some text</TAGNAME></TAGNAME>
							*             (cursor represented by |)
							*/
							int currentStartPosition = xmlTags.tagOpenEnd;
							int closeTagsRemaining = 1;
							FindResult nextCloseTag;
							do
							{
								nextCloseTag = findCloseTag(tagName, currentStartPosition, docLength);
								if (nextCloseTag.success)
								{
									--closeTagsRemaining;
									// Open tag found
									// Now we need to check how many close tags there are between the open tag we just found,
									// and our close tag
									// eg. (Cursor == | )
									// <TAGNAM|E attrib="value"><TAGNAME>something</TAGNAME></TAGNAME></TAGNAME>
									//                                            ^^^^^^^^ we've found this guy
									//                         ^^^^^^^^^ Now we need to find this fella
									FindResult inbetweenOpenTag;
									int currentEndPosition = nextCloseTag.start;
									int openTagsFound = 0;

									do
									{
										inbetweenOpenTag = findOpenTag(tagName, currentStartPosition, currentEndPosition);

										if (inbetweenOpenTag.success)
										{
											++openTagsFound;
											currentStartPosition = inbetweenOpenTag.end;
										}

									} while (inbetweenOpenTag.success);

									// If we didn't find any open tags between our open and the close,
									// and there's no close tags remaining to find
									// then the close we found was the right one, and we can return it
									if (0 == openTagsFound && 0 == closeTagsRemaining)
									{
										xmlTags.tagCloseStart = nextCloseTag.start;
										xmlTags.tagCloseEnd = nextCloseTag.end + 1;
										tagFound = true;
									}
									else
									{

										// Need to find the same number of closing tags, without opening tags etc.
										closeTagsRemaining += openTagsFound;
										currentStartPosition = nextCloseTag.end;
									}
								}
							} while (!tagFound && closeTagsRemaining > 0 && nextCloseTag.success);
						} // end if (selfclosingtag)... else {
					} // end if (-1 != closeAngle)  {

				} // end if tagName.size() != 0
			} // end open tag test
		}
	}
	return tagFound;
}
/// <summary>Search for open tag of XML-element</summary>
/// <param name="tagName">Tag name</param>  
/// <param name="start">Start position for search</param>  
/// <param name="end">End position for search</param>  
/// <returns>result of search as structure with Sci_CharacterRange</returns> 
XmlMatchedTagsHighlighter::FindResult XmlMatchedTagsHighlighter::findOpenTag(const std::string& tagName, int start, int end)
{
	std::string search("<");
	search.append(tagName);
	FindResult openTagFound;
	openTagFound.success = false;
	FindResult result;
	int nextChar = 0;
	int styleAt;
	int searchStart = start;
	int searchEnd = end;
	bool forwardSearch = (start < end);
	do
	{

		result = findText(search.c_str(), searchStart, searchEnd, 0);
		if (result.success)
		{
			nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, result.end));
			styleAt = static_cast<int32_t>(_pEditView->execute(SCI_GETSTYLEAT, result.start));
			if (styleAt != SCE_H_CDATA && styleAt != SCE_H_DOUBLESTRING && styleAt != SCE_H_SINGLESTRING && styleAt != SCE_H_COMMENT)
			{
				// We've got an open tag for this tag name (i.e. nextChar was space or '>')
				// Now we need to find the end of the start tag.

				// Common case, the tag is an empty tag with no whitespace. e.g. <TAGNAME>
				if (nextChar == '>')
				{
					openTagFound.end = result.end;
					openTagFound.success = true;
				}
				else if (isWhitespace(nextChar))
				{
					int closeAnglePosition = findCloseAngle(result.end, forwardSearch ? end : start);
					if (-1 != closeAnglePosition && '/' != _pEditView->execute(SCI_GETCHARAT, closeAnglePosition - 1))
					{
						openTagFound.end = closeAnglePosition;
						openTagFound.success = true;
					}
				}
			}

		}

		if (forwardSearch)
		{
			searchStart = result.end + 1;
		}
		else
		{
			searchStart = result.start - 1;
		}

		// Loop while we've found a <TAGNAME, but it's either in a CDATA section,
		// or it's got more none whitespace characters after it. e.g. <TAGNAME2
	} while (result.success && !openTagFound.success);

	openTagFound.start = result.start;


	return openTagFound;

}
/// <summary>Search for nearest angle bracket '>'.
/// If start < end, it will be reversed in function, so the search is always forward</summary>
/// <param name="startPosition">Start position for search</param>  
/// <param name="endPosition">End position for search</param>  
/// <returns>position of '>'</returns> 
int XmlMatchedTagsHighlighter::findCloseAngle(int startPosition, int endPosition)
{
	// We'll search for the next '>', and check it's not in an attribute using the style
	FindResult closeAngle;

	bool isValidClose;
	int returnPosition = -1;

	// Only search forwards
	if (startPosition > endPosition)
	{
		int temp = endPosition;
		endPosition = startPosition;
		startPosition = temp;
	}

	do
	{
		isValidClose = false;

		closeAngle = findText(">", startPosition, endPosition);
		// if not found - break from loop
		if (closeAngle.success)
		{
			auto style = _pEditView->execute(SCI_GETSTYLEAT, closeAngle.start);
			// As long as we're not in an attribute (  <TAGNAME attrib="val>ue"> is VALID XML. )
			if (style != SCE_H_DOUBLESTRING && style != SCE_H_SINGLESTRING && style != SCE_H_COMMENT)
			{
				returnPosition = closeAngle.start;
				isValidClose = true;
				// break from loop
			}
			else
			{
				// move start to finded '>' and search again
				startPosition = closeAngle.end;
			}
		}

	} while (closeAngle.success && isValidClose == false);

	return returnPosition;
}
/// <summary>Search for close tag</summary>
/// <param name="tagName">Tag name</param>  
/// <param name="start">Start position for search</param>  
/// <param name="end">End position for search</param>  
/// <returns>result of search as structure with Sci_CharacterRange</returns> 
XmlMatchedTagsHighlighter::FindResult XmlMatchedTagsHighlighter::findCloseTag(const std::string& tagName, int start, int end)
{
	std::string search("</");
	search.append(tagName);
	FindResult closeTagFound;
	closeTagFound.success = false;
	FindResult result;
	int nextChar;
	int searchStart = start;
	int searchEnd = end;
	bool forwardSearch = (start < end);
	bool validCloseTag;
	do
	{
		validCloseTag = false;
		result = findText(search.c_str(), searchStart, searchEnd, 0);
		if (result.success)
		{
			nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, result.end));
			auto styleAt = _pEditView->execute(SCI_GETSTYLEAT, result.start);

			// Setup the parameters for the next search, if there is one.
			if (forwardSearch)
			{
				searchStart = result.end + 1;
			}
			else
			{
				searchStart = result.start - 1;
			}

			if (styleAt != SCE_H_CDATA && styleAt != SCE_H_SINGLESTRING && styleAt != SCE_H_DOUBLESTRING && styleAt != SCE_H_COMMENT) // If what we found was in CDATA section, it's not a valid tag.
			{
				// Common case - '>' follows the tag name directly
				if (nextChar == '>')
				{
					validCloseTag = true;
					closeTagFound.start = result.start;
					closeTagFound.end = result.end;
					closeTagFound.success = true;
				}
				else if (isWhitespace(nextChar))  // Otherwise, if it's whitespace, then allow whitespace until a '>' - any other character is invalid.
				{
					int whitespacePoint = result.end;
					do
					{
						++whitespacePoint;
						nextChar = static_cast<int32_t>(_pEditView->execute(SCI_GETCHARAT, whitespacePoint));

					} while (isWhitespace(nextChar));

					if (nextChar == '>')
					{
						validCloseTag = true;
						closeTagFound.start = result.start;
						closeTagFound.end = whitespacePoint;
						closeTagFound.success = true;
					}
				}
			}
		}

	} while (result.success && !validCloseTag);

	return closeTagFound;

}

/// <summary>Find string</summary>
/// <param name="*text">Searched string</param>  
/// <param name="start">Start position for search</param>  
/// <param name="end">End position for search</param>  
/// <param name="flags">Search flags</param>  
/// <returns>result of search as structure with Sci_CharacterRange</returns> 
XmlMatchedTagsHighlighter::FindResult XmlMatchedTagsHighlighter::findText(const char *text, int start, int end, int flags)
{
	FindResult returnValue;

	Sci_TextToFind search;
	search.lpstrText = const_cast<char *>(text); // Grrrrrr
	search.chrg.cpMin = start;
	search.chrg.cpMax = end;
	int result = static_cast<int32_t>(_pEditView->execute(SCI_FINDTEXT, flags, reinterpret_cast<LPARAM>(&search)));
	if (-1 == result)
	{
		returnValue.success = false;
	}
	else
	{
		returnValue.success = true;
		returnValue.start = search.chrgText.cpMin;
		returnValue.end = search.chrgText.cpMax;
	}
	return returnValue;
}
/// <summary>Highlight tags around carret position</summary>
/// <param name="doHilite">Highlight tag, if true</param>  
/// <param name="doHiliteAttr">Highlight attributes, if true</param>  
/// <param name="gotoTag">Move carret to tag, if true</param>  
/// <returns>true-carret inside tag, false - carret outside tag</returns> 
bool XmlMatchedTagsHighlighter::tagMatch(bool doHilite, bool doHiliteAttr, bool gotoTag) 
{
	bool bFound = false;

	// Clean up all marks of previous action
	_pEditView->clearIndicator(SCE_UNIVERSAL_TAGMATCH);
	_pEditView->clearIndicator(SCE_UNIVERSAL_TAGATTR);

	// Get the original targets and search options to restore after tag matching operation
	auto originalStartPos = _pEditView->execute(SCI_GETTARGETSTART);
	auto originalEndPos = _pEditView->execute(SCI_GETTARGETEND);
	auto originalSearchFlags = _pEditView->execute(SCI_GETSEARCHFLAGS);

	XmlMatchedTagsPos xmlTags;

    // Detect if it's a xml/html tag. If yes, Colour it!
	if (getXmlMatchedTagsPos(xmlTags))
	{
		bFound = true;

		if (doHilite)
		{
			_pEditView->execute(SCI_SETINDICATORCURRENT, SCE_UNIVERSAL_TAGMATCH);
			int openTagTailLen = 2;

			// Colourising the close tag firstly
			if ((xmlTags.tagCloseStart != -1) && (xmlTags.tagCloseEnd != -1))
			{
				_pEditView->execute(SCI_INDICATORFILLRANGE,  xmlTags.tagCloseStart, xmlTags.tagCloseEnd - xmlTags.tagCloseStart);
				// tag close is present, so it's not single tag
				openTagTailLen = 1;
			}

			// Colourising the open tag
			_pEditView->execute(SCI_INDICATORFILLRANGE,  xmlTags.tagOpenStart, xmlTags.tagNameEnd - xmlTags.tagOpenStart);
			_pEditView->execute(SCI_INDICATORFILLRANGE,  xmlTags.tagOpenEnd - openTagTailLen, openTagTailLen);

			// Colouising its attributs
			if (doHiliteAttr)
			{
				vector< pair<int, int> > attributes = getAttributesPos(xmlTags.tagNameEnd, xmlTags.tagOpenEnd - openTagTailLen);
				_pEditView->execute(SCI_SETINDICATORCURRENT,  SCE_UNIVERSAL_TAGATTR);
				for (size_t i = 0 ; i < attributes.size() ; i++)
				{
					_pEditView->execute(SCI_INDICATORFILLRANGE,  attributes[i].first, attributes[i].second - attributes[i].first);
				}
			}

			// Colouising indent guide line position
			if (_pEditView->isShownIndentGuide())
			{
				int columnAtCaret  = int(_pEditView->execute(SCI_GETCOLUMN, xmlTags.tagOpenStart));
				int columnOpposite = int(_pEditView->execute(SCI_GETCOLUMN, xmlTags.tagCloseStart));

				int lineAtCaret  = int(_pEditView->execute(SCI_LINEFROMPOSITION, xmlTags.tagOpenStart));
				int lineOpposite = int(_pEditView->execute(SCI_LINEFROMPOSITION, xmlTags.tagCloseStart));

				if (xmlTags.tagCloseStart != -1 && lineAtCaret != lineOpposite)
				{
					_pEditView->execute(SCI_BRACEHIGHLIGHT, xmlTags.tagOpenStart, xmlTags.tagCloseEnd-1);
					_pEditView->execute(SCI_SETHIGHLIGHTGUIDE, (columnAtCaret < columnOpposite)?columnAtCaret:columnOpposite);
				}
			}
		}
	}

	// restore the original targets and search options to avoid the conflit with search/replace function
	_pEditView->execute(SCI_SETTARGETRANGE, originalStartPos, originalEndPos);
	_pEditView->execute(SCI_SETSEARCHFLAGS, originalSearchFlags);


	if (gotoTag && bFound)
	{
		int caretPos = _pEditView->execute(SCI_GETCURRENTPOS);
		// check where we are
		if (xmlTags.tagCloseStart > -1)
		{
			if (caretPos >= xmlTags.tagOpenStart && caretPos <= xmlTags.tagOpenEnd)
				_pEditView->execute(SCI_GOTOPOS, xmlTags.tagCloseStart+1);
			else
				_pEditView->execute(SCI_GOTOPOS, xmlTags.tagOpenStart+1);
		}
	}
	return bFound;
}

/// bool erasedTag(TAG value) { return (value.second == -1); }

///// <summary>Lookup all tags in document</summary>
///// <returns>vector of unpaired tags</returns> 
//vector< pair<CString, int> > XmlMatchedTagsHighlighter::lookupTags()
//{
//	vector<TAG> tags;
//
//	// Get the original targets and search options to restore after tag matching operation
//	auto originalStartPos = _pEditView->execute(SCI_GETTARGETSTART);
//	auto originalEndPos = _pEditView->execute(SCI_GETTARGETEND);
//	auto originalSearchFlags = _pEditView->execute(SCI_GETSEARCHFLAGS);
//	
//	// search options
//    _pEditView->execute(SCI_SETSEARCHFLAGS, SCFIND_REGEXP);
//    int docLen = _pEditView->getCurrentDocLen();
//
//	// regexp pattern to match any XML tag
//	CString pattern(L"(<[^(><)]+>)");
//	int pattLen = pattern.GetLength();
//	CT2A patt (pattern, CP_UTF8);
//	char *tmp=(char *)malloc(1024);
//	if (patt && tmp) 
//	{
//		int pos = 0;
//		while (pos >=0 && pos < docLen)
//		{
//			_pEditView->execute(SCI_SETTARGETSTART,pos);
//			_pEditView->execute(SCI_SETTARGETEND,docLen);
//			pos = _pEditView->execute(SCI_SEARCHINTARGET,pattLen, (LPARAM)(LPSTR) patt);
//
//			if (pos >=0 && pos < docLen)
//			{
//				int tagPos = _pEditView->execute(SCI_GETTARGETSTART);
//				pos = _pEditView->execute(SCI_GETTARGETEND);
//
//				Sci_TextRange tr;
//				tr.chrg.cpMin = tagPos;
//				tr.chrg.cpMax = pos;
//				tr.lpstrText = tmp;
//				_pEditView->execute(SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
//				CString tagName(tmp);
//				// check tag type: open, close or single
//				bool isCloseTag = (tagName.Find (L"</")==0);
//				bool isSingleTag = false;
//				if (!isCloseTag) isSingleTag = (tagName.Find (L"/>")>0);
//				
//				if (!isSingleTag)
//				{
//					// remove tag attributes
//					if (isCloseTag) tagName = tagName.Mid(1, tagName.Find(L">",1)-1);
//					else
//					{
//						int attr = tagName.Find(L" ",1);
//						if (attr > 0) tagName = tagName.Mid(1, attr-1);
//						else tagName = tagName.Mid(1, tagName.GetLength()-2);
//					}
//
//					if (!tags.empty() && tagName.CompareNoCase(L"/"+tags.back().first) == 0) tags.pop_back();
//					else tags.push_back (TAG(tagName,tagPos));
//				}
//			}
//		}
//		free(tmp);
//	}
//
//	// remove matched (correct) tags
//	if (!tags.empty())
//	{
//		tagIterator firstTag = tags.begin();
//		tagIterator lastTag = tags.end()-1;
//
//		while ( &firstTag && &lastTag && firstTag < lastTag)
//		{
//			if ((*lastTag).first.CompareNoCase(L"/"+(*firstTag).first) == 0)
//			{
//				(*firstTag).second = -1;
//				(*lastTag).second = -1;
//				firstTag++;
//				lastTag--;
//			}
//			// locate matched tag pair
//			else 
//			{
//				tagIterator firstTmp = firstTag;
//				while ((firstTmp < lastTag) && ( (*lastTag).first.CompareNoCase(L"/"+(*firstTmp).first) != 0))
//				{
//					firstTmp++;
//				}
//				if (firstTmp < lastTag) firstTag = firstTmp;
//				else lastTag--;
//			}
//		}
//
//		tags.erase (remove_if(tags.begin(), tags.end(), erasedTag), tags.end());
//	}
//
//	// restore the original targets and search options to avoid the conflit with search/replace function
//	_pEditView->execute(SCI_SETTARGETRANGE, originalStartPos, originalEndPos);
//	_pEditView->execute(SCI_SETSEARCHFLAGS, originalSearchFlags);
//	
//	return tags;
//}

///// <summary>Move carret to unpaired tag. Try to move on closest forward wrong tag or on first wrong tag</summary>
///// <returns>true - wrong tag found, carret moved, else - no wring tags</returns> 
//bool XmlMatchedTagsHighlighter::gotoWrongTag()
//{
//	vector<TAG> wrongTags = lookupTags();
//	if (!wrongTags.empty())
//	{
//		int caretPos = _pEditView->execute(SCI_GETCURRENTPOS);
//		int newPos = wrongTags[0].second;
//		for (tagIterator it=wrongTags.begin(); it!=wrongTags.end(); it++)
//			if ((*it).second > caretPos)
//			{
//				newPos = (*it).second;
//				break;
//			}
//		_pEditView->execute(SCI_GOTOPOS, newPos);
//		return true;
//	}
//	return false;
//	/*
//	if (!wrongTags.empty())
//	{
//		_pEditView->execute(SCI_GOTOPOS, wrongTags[0].second);
//		return true;
//	}
//	return false;
//	*/
//}	