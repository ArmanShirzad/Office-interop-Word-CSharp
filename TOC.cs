using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace interop_usermanual
{
    public class TOC
    {

        public TOC(int UpperHeadingLevel, int LowerHeadingLevel, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oTemplate, ref object oEndOfDoc)
        {
            object oTrueValue = true;
            Object what = WdGoToItem.wdGoToPage;
            object count = 2;
            oWord.Selection.GoTo(what, ref oMissing, ref count, ref oMissing);

            oWord.Selection.Font.Bold = 1;
            oWord.Selection.TypeText("TableOfContents\n");
            oDoc.TablesOfContents.Add(oWord.Selection.Range, ref oTrueValue, UpperHeadingLevel, LowerHeadingLevel,
                    UseFields: true, TableID: "TableOfContents", RightAlignPageNumbers: true, IncludePageNumbers: true,
                    AddedStyles: true, UseHyperlinks: true, HidePageNumbersInWeb: true, UseOutlineLevels: true);

            oWord.Selection.TypeText("\f");//page break
        }
    }
}
