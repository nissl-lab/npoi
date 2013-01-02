/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSLF.Model;

/**
 * This class represents a slide's notes in a PowerPoint Document. It
 *  allows access to the text within, and the layout. For now, it only
 *  does the text side of things though
 *
 * @author Nick Burch
 */

public class Notes : Sheet
{
  private TextRun[] _Runs;

  /**
   * Constructs a Notes Sheet from the given Notes record.
   * Initialises TextRuns, to provide easier access to the text
   *
   * @param notes the Notes record to read from
   */
  public Notes (NPOI.HSLF.record.Notes notes) {
      base(notes, notes.GetNotesAtom().GetSlideID());

	// Now, build up TextRuns from pairs of TextHeaderAtom and
	//  one of TextBytesAtom or TextCharsAtom, found inside
	//  EscherTextboxWrapper's in the PPDrawing
	_Runs = FindTextRuns(getPPDrawing());

	// Set the sheet on each TextRun
	for (int i = 0; i < _Runs.Length; i++)
		_Runs[i].SetSheet(this);
  }


  // Accesser methods follow

  /**
   * Returns an array of all the TextRuns found
   */
  public TextRun[] GetTextRuns() { return _Runs; }

    /**
     * Return <code>null</code> - Notes Masters are not yet supported
     */
    public MasterSheet GetMasterSheet() {
        return null;
    }

}





