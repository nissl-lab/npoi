/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.HSLF.Record;
using NPOI.HSLF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Model
{
	public class HeadersFooters
	{
		private static String _ppt2007tag = "___PPT12";

    private HeadersFootersContainer _container;
    private HSLFSheet _sheet;
    private bool _ppt2007;


    public HeadersFooters(HSLFSlideShow ppt, short headerFooterType)
			:this(ppt.GetSlideMasters().Get(0), headerFooterType)
		{
		}

		public HeadersFooters(HSLFSheet sheet, short headerFooterType)
		{
			_sheet = sheet;

			//@SuppressWarnings("resource")
	
		HSLFSlideShow ppt = _sheet.GetSlideShow();
			Document doc = ppt.GetDocumentRecord();

			// detect if this ppt was saved in Office2007
			String tag = ppt.GetSlideMasters().Get(0).GetProgrammableTag();
			_ppt2007 = _ppt2007tag.Equals(tag);

			SheetContainer sc = _sheet.GetSheetContainer();
			HeadersFootersContainer hdd = (HeadersFootersContainer)sc.FindFirstOfType(RecordTypes.HeadersFooters.typeID);
			// bool ppt2007 = sc.findFirstOfType(RecordTypes.RoundTripContentMasterId.typeID) != null;

			if (hdd == null)
			{
				foreach (Record.Record ch in doc.GetChildRecords())
				{
					if (ch is HeadersFootersContainer

					&& ((HeadersFootersContainer)ch).GetOptions() == headerFooterType) {
					hdd = (HeadersFootersContainer)ch;
					break;
				}
			}
		}

        if (hdd == null) {
            hdd = new HeadersFootersContainer(headerFooterType);
		Record lst = doc.findFirstOfType(RecordTypes.List.typeID);
		doc.addChildAfter(hdd, lst);
        }
	_container = hdd;
    }

/**
 * Headers's text
 *
 * @return Headers's text
 */
public String getHeaderText()
{
	CString cs = _container == null ? null : _container.getHeaderAtom();
	return getPlaceholderText(Placeholder.HEADER, cs);
}

/**
 * Sets headers's text
 *
 * @param text headers's text
 */
public void setHeaderText(String text)
{
	setHeaderVisible(true);
	CString cs = _container.getHeaderAtom();
	if (cs == null)
	{
		cs = _container.addHeaderAtom();
	}

	cs.setText(text);
}

/**
 * Footer's text
 *
 * @return Footer's text
 */
public String getFooterText()
{
	CString cs = _container == null ? null : _container.getFooterAtom();
	return getPlaceholderText(Placeholder.FOOTER, cs);
}

/**
 * Sets footers's text
 *
 * @param text footers's text
 */
public void setFootersText(String text)
{
	setFooterVisible(true);
	CString cs = _container.getFooterAtom();
	if (cs == null)
	{
		cs = _container.addFooterAtom();
	}

	cs.setText(text);
}

/**
 * This is the date that the user wants in the footers, instead of today's date.
 *
 * @return custom user date
 */
public String getDateTimeText()
{
	CString cs = _container == null ? null : _container.getUserDateAtom();
	return getPlaceholderText(Placeholder.DATETIME, cs);
}

/**
 * Sets custom user date to be displayed instead of today's date.
 *
 * @param text custom user date
 */
public void setDateTimeText(String text)
{
	setUserDateVisible(true);
	setDateTimeVisible(true);
	CString cs = _container.getUserDateAtom();
	if (cs == null)
	{
		cs = _container.addUserDateAtom();
	}

	cs.setText(text);
}

/**
 * whether the footer text is displayed.
 */
public bool isFooterVisible()
{
	return isVisible(HeadersFootersAtom.fHasFooter, Placeholder.FOOTER);
}

/**
 * whether the footer text is displayed.
 */
public void setFooterVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasFooter, flag);
}

/**
 * whether the header text is displayed.
 */
public bool isHeaderVisible()
{
	return isVisible(HeadersFootersAtom.fHasHeader, Placeholder.HEADER);
}

/**
 * whether the header text is displayed.
 */
public void setHeaderVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasHeader, flag);
}

/**
 * whether the date is displayed in the footer.
 */
public bool isDateTimeVisible()
{
	return isVisible(HeadersFootersAtom.fHasDate, Placeholder.DATETIME);
}

/**
 * whether the date is displayed in the footer.
 */
public void setDateTimeVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasDate, flag);
}

/**
 * whether the custom user date is used instead of today's date.
 */
public bool isUserDateVisible()
{
	return isVisible(HeadersFootersAtom.fHasUserDate, Placeholder.DATETIME);
}

public CString getHeaderAtom()
{
	return _container.getHeaderAtom();
}

public CString getFooterAtom()
{
	return _container.getFooterAtom();
}

public CString getUserDateAtom()
{
	return _container.getUserDateAtom();
}

/**
 * whether the date is displayed in the footer.
 */
public void setUserDateVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasUserDate, flag);
}

/**
 * whether today's date is used.
 */
public bool isTodayDateVisible()
{
	return isVisible(HeadersFootersAtom.fHasTodayDate, Placeholder.DATETIME);
}

/**
 * whether the todays date is displayed in the footer.
 */
public void setTodayDateVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasTodayDate, flag);
}

/**
 * whether the slide number is displayed in the footer.
 */
public bool isSlideNumberVisible()
{
	return isVisible(HeadersFootersAtom.fHasSlideNumber, Placeholder.SLIDE_NUMBER);
}

/**
 * whether the slide number is displayed in the footer.
 */
public void setSlideNumberVisible(bool flag)
{
	setFlag(HeadersFootersAtom.fHasSlideNumber, flag);
}

/**
 *  An integer that specifies the format ID to be used to style the datetime.
 *
 * @return an integer that specifies the format ID to be used to style the datetime.
 */
public int getDateTimeFormat()
{
	return _container.getHeadersFootersAtom().getFormatId();
}

/**
 *  An integer that specifies the format ID to be used to style the datetime.
 *
 * @param formatId an integer that specifies the format ID to be used to style the datetime.
 */
public void setDateTimeFormat(int formatId)
{
	_container.getHeadersFootersAtom().setFormatId(formatId);
}

private bool isVisible(int flag, Placeholder placeholderId)
{
	bool visible;
	if (_ppt2007)
	{
		HSLFSimpleShape ss = _sheet.getPlaceholder(placeholderId);
		visible = ss instanceof HSLFTextShape && ((HSLFTextShape)ss).getText() != null;
	}
	else
	{
		visible = _container.getHeadersFootersAtom().getFlag(flag);
	}
	return visible;
}

private String getPlaceholderText(Placeholder ph, CString cs)
{
	String text;
	if (_ppt2007)
	{
		HSLFSimpleShape ss = _sheet.getPlaceholder(ph);
		text = (ss instanceof HSLFTextShape) ? ((HSLFTextShape)ss).getText() : null;

// default text in master placeholders is not visible
if ("*".equals(text))
{
	text = null;
}
        } else
{
	text = (cs == null) ? null : cs.getText();
}
return text;
    }

    private void setFlag(int type, bool flag)
{
	_container.getHeadersFootersAtom().setFlag(type, flag);
}

/**
 * @return true, if this is a ppt 2007 document and header/footer are stored as placeholder shapes
 */
public bool isPpt2007()
{
	return _ppt2007;
}

public HeadersFootersContainer getContainer()
{
	return _container;
}
	}
}
