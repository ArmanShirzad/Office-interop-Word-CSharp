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
    class text
    {
        public string Context { get; set; }
        public string Font { get; set; }
        public int Fontsize { get; set; }
        public string Align { get; set; }
        public int Bold { get; set; }
        public text(string context, string font, int fontsize, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oEndOfDoc, string captionText="",string color = "black", string align = "left", int bold = 0,string style="plain")
        {

            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Word.Paragraph oPara = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = context;
            oPara.Range.Font.Name = font;
            oPara.Range.Font.Size = fontsize;
            oPara.Range.Font.Bold = bold;   //1 is bold 0 is normal text
            switch (style)
            {
                case "h1":
                    oPara.Range.ParagraphFormat.set_Style(WdBuiltinStyle.wdStyleHeading1);
                    break;
                case "h2":
                    oPara.Range.ParagraphFormat.set_Style(WdBuiltinStyle.wdStyleHeading2);
                    break;

                case "h3":
                    oPara.Range.ParagraphFormat.set_Style(WdBuiltinStyle.wdStyleHeading3);
                    break;

                case "toch":
                    oPara.Range.ParagraphFormat.set_Style(WdBuiltinStyle.wdStyleTocHeading);
                    break;

                case "caption":
                    context = "";
                    Range wrdRng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng1.ParagraphFormat.SpaceAfter = 0;
                    oWord.CaptionLabels.Add(captionText);
                    object caption = oWord.CaptionLabels[captionText];
                    wrdRng1.InsertCaption(ref caption, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    // insert the table after the caption…
                    
                    wrdRng1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    wrdRng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng1.InsertParagraphAfter();
                    break;
                default:
                    oPara.Range.ParagraphFormat.set_Style(WdBuiltinStyle.wdStyleNormal);
                    break;



            }

            switch (color)      //color items can be modified according to AcidPro needs
            {
                case "blue":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                    break;

                case "green":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                    break;
                case "teal":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdTeal;
                    break;
                case "gray":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                    break;
                case "violet":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdViolet;
                    break;
                case "red":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdRed;
                    break;
                case "yellow":
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdDarkYellow;
                    break;
                default:
                    oPara.Range.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    break;
            }
            if (bold == 1) { oPara.Range.Font.Bold = 1; }
            else
            {
                oPara.Range.Font.Bold = 0;
            }


            if (align == "left") { oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; }
            //else if (align == "right") { oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; }
            //else { oPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; }

            oPara.Range.ParagraphFormat.SpaceAfter = 0.0f;
           oPara.Range.InsertParagraphAfter();
        
        }

    }
}
