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
    class footer
    {
        public string Footer { get; set; }
        public string Font { get; set; }
        public int Fontsize { get; set; }
        public string Align { get; set; }
        public int Bold { get; set; }
        public string Color { get; set; }
        public footer(string footer, string font, int fontsize, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oEndOfDoc, int bold = 0, string color = "black", string align = "left")
        {
            foreach (Word.Section section in oDoc.Sections)
            {
                //Get the footer range and add the footer details.
                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Fields.Add(footerRange, ref oMissing);
                footerRange.Text = footer;
                footerRange.Font.Name = font;
                footerRange.Font.Size = fontsize;
                footerRange.Font.Bold = bold;


                if (bold == 1) { footerRange.Font.Bold = 1; }
                else
                {
                    footerRange.Font.Bold = 0;
                }
                switch (color)      //color items can be modified according to AcidPro needs
                {
                    case "blue":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                        break;

                    case "green":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                        break;
                    case "teal":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdTeal;
                        break;
                    case "gray":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                        break;
                    case "violet":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdViolet;
                        break;
                    case "red":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdRed;
                        break;
                    case "yellow":
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkYellow;
                        break;
                    default:
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                        break;
                }
                if (align == "center") { footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; }
                else if (align == "right") { footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; }
                else { footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; }
            }
        }
    }
}
