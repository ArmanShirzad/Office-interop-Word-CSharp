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
    class Bulletlist
    {
        public List<string> Parentbullet { get; set; }
        public List<string> Childbullet { get; set; }

        public Bulletlist(List<string> parentbullet, List<string> childbullet,List<int> index, ref Word._Document oDoc,
              ref object oMissing, ref Word._Application oWord, ref object oTemplate, ref object oEndOfDoc)
        {
           
            Range range1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object indentStyle = "Normal Indent";
            range1.set_Style(ref indentStyle);
            ListGallery listGallery = oWord.ListGalleries[WdListGalleryType.wdBulletGallery];
            range1.Select();
            oWord.Selection.Range.ListFormat.ApplyListTemplateWithLevel(
            listGallery.ListTemplates[1],
            ContinuePreviousList: false,
            ApplyTo: WdListApplyTo.wdListApplyToSelection,
            DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
            foreach (string root in parentbullet)
            {
                oWord.Selection.Range.ListFormat.ListOutdent();
                oWord.Selection.TypeText(root);
                // Set text to key in
                oWord.Selection.TypeParagraph();  // Simulate typing in MS Word
                int i = parentbullet.IndexOf(root);
                if (index.Any(x=>x==i))
                {
                    oWord.Selection.Range.ListFormat.ListIndent();

                    foreach (string child in childbullet)
                    {
                        oWord.Selection.TypeText(child);
                        // Set text to key in
                        oWord.Selection.TypeParagraph();  // Simulate typing in MS Word


                    }
                    //oWord.Selection.Range.ListFormat.ListOutdent();

                }

            }
            //foreach (string root in parentbullet)
            //{
            //    oWord.Selection.TypeText(root);
            //    oWord.Selection.TypeParagraph();  
            //    oWord.Selection.Range.ListFormat.ListOutdent();
            //}
            //oWord.Selection.Range.ListFormat.ListIndent();
            //foreach (string child in childbullet)
            //{

            //    oWord.Selection.TypeText(child);
            //    oWord.Selection.TypeParagraph();  


            //}
         
            oWord.Selection.TypeBackspace();
            oWord.Selection.ClearParagraphAllFormatting();

        }
    }
}
