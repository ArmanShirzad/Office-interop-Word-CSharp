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
    class header
    {
        public string Header { get; set; }
        public string Font { get; set; }
        public int Fontsize { get; set; }
        public string Align { get; set; }
        public int Bold { get; set; }
        public string Color { get; set; }
        public header(string header, string font, int fontsize, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oEndOfDoc, int bold = 0, string color = "black", string align = "left")
        {
            foreach (Word.Section section in oDoc.Sections)
            {
                //Get the header range and add the header details.
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, ref oMissing);
                headerRange.Text = header;
                headerRange.Font.Name = font;
                headerRange.Font.Size = fontsize;
                headerRange.Font.Bold = bold;
                if (bold == 1) { headerRange.Font.Bold = 1; }
                else
                {
                    headerRange.Font.Bold = 0;
                }
                switch (color)      //color items can be modified according to AcidPro needs
                {
                    case "blue":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                        break;

                    case "green":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                        break;
                    case "teal":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdTeal;
                        break;
                    case "gray":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                        break;
                    case "violet":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdViolet;
                        break;
                    case "red":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdRed;
                        break;
                    case "yellow":
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkYellow;
                        break;
                    default:
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                        break;
                }
                if (align == "center") { headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; }
                else if (align == "right") { headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; }
                else { headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; }
            }

        }
    }
}
