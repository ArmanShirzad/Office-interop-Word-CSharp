using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Word = Microsoft.Office.Interop.Word;

namespace interop_usermanual
{
  internal  class generator
    {

        public generator()
        {
//DEFINE OBJECTS OF THE CODE
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            object oStartofDoc = "\\startofdoc";

//THE 2  MAIN INTERFACE OF OFFICE.INTEROP.WORD, MOST OF OUR WORK WILL BE DONE USING ONE OF THE TWO AND THEIR CHILDREN
            Word._Application oWord;
            Word._Document oDoc;

//CREATING THE WORD APPLICATION
            oWord = new Word.Application();
            oWord.Visible = true;

//DEFINING A DYNAMIC ADDRESS WHERE THE TEMPLATE IS HERE YOU GO ;)
            var temp = Directory.GetCurrentDirectory();
            DirectoryInfo temp1 = Directory.GetParent(temp);
            Object oTemplate = Path.Combine(Directory.GetParent(temp1.FullName).FullName, "Template\\Template.docx");

//AN INSTANCE OF WORD ALREADY INITIALIZED BUT THERES NOTHING IN IT SO LETS CREATE A BLANK DOCUMENT
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            ref oMissing, ref oMissing);

//NEEDED OBJECTS DEFINED FOR PAGE BREAK AND SECTION BREAK ,WDSTORY REFERS TO THE CURRENT DOCUMENT FROM THE START TO THE END!
            Object objBreak = Word.WdBreakType.wdPageBreak;
            Object objUnit = Word.WdUnits.wdStory;
            Object objbreak1 = Word.WdBreakType.wdSectionBreakNextPage;
            Object objUnit1 = Word.WdUnits.wdStory;

//ON THE FIRST PAGE

            //set first page to have a different header &footer
            oDoc.PageSetup.DifferentFirstPageHeaderFooter = -1;

            //// Setting Different First page Header & Footer null so theres no header /footer in THE first page
            oDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
            oDoc.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
//ANOTHER APPROACH FOR THE TASK above
            //oWord.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = -1;
            //oDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            //oDoc.ActiveWindow.Selection.TypeText("HEader Text");
            
//LINE BELOW IS A TRICK NOT TO HAVE AN EXTRA TITLE FOR TOC, COMMENT IT AND CHECK OUT WHAT HAPPENS!
            text h2 = new text("", "Segoe UI Light (Headings)", 26, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "black", "left", 0, "toch");
          
           //since toc places in the  first page all the next elements will be imported from the beginning of second page
           //set style of toc to heading 1 so all h1 s will be in our table of contents
            string style = "h1";

            text h3 = new text("Welcome", "Segoe UI (Body)", 11, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "black", "left", 0, "h1");

        
            //ADD A HORIZONTAL LINE
            Word.Range rng2 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oDoc.InlineShapes.AddHorizontalLineStandard(rng2);
            
            text h6 = new text("Lets Write Some text here with CAlibri font and size font 12 thats all u need to Write in the constructor really", "Calibri", 12, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc);

//I WANNA START THE NEXT PAGE LETS INSERT A PAGE BREAK USING OBJE
            oWord.Selection.EndKey(ref objUnit, ref oMissing);
            oWord.Selection.InsertBreak(ref objBreak);
            //PAGE 3
            text h4 = new text("Whenever we want some item we just create a new object from the required item class ", "Segoe UI (Body)", 11, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "black", "left", 0, style);

            Word.Range rng3 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oDoc.InlineShapes.AddHorizontalLineStandard(rng3);

            //insert linebreak  BY PUTTING  "\v" AT THE END OF THE LAST PRINTED SENTENCE
            text h8 = new text("some random text again look at the font \t", "Segoe UI (Body)", 12, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc);
            oWord.Selection.EndKey(ref objUnit, ref oMissing);
            oWord.Selection.InsertBreak(ref objBreak);
            //PAGE 4
            text h9 = new text("New Features: By Default the color is black and alignment is left , theres no caption text and the style is plain normal text but you can go ahead and change em according to your needs cheers! ", "Segoe UI (Body)", 12, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "black", "left", 1);

            text h10 = new text("Bullet list here it comes:", "Segoe UI (Body)", 12, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "black", "left", 1);
 //ADD A BULLETLIST REJOICE!

            //add a range to the end of doc this is how we define were the incoming element should print ... how? its easy end of doc always refers to the position right after the previous (last) element was printed

            Range somerange = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            var parentbullet = new List<string> {
                "Item 1",
                "Item 2",
                "Item 3",
                "Item 4",
                "Item 5",
                "Item 5",
                ""

                            };

           

            var childbullet = new List<string>{
                "sub Item 1",
                "sub Item 2",
                "sub Item 3",
                "sub Item 4",
                "sub Item 5",
                ""

            };
            // so here in the following list each item shows for which index of the main indent list we want a sub item ,here we say 1,3 means for the second and fourth item of the main list we want subitems (index starts from zero)
            List<int> index = new List<int>() {

                1,3

            };
            Bulletlist mybulletlist1 = new Bulletlist(parentbullet, childbullet, index, ref oDoc,
                ref oMissing, ref oWord, ref oTemplate, ref oEndOfDoc);

            //insert page break
            oWord.Selection.EndKey(ref objUnit, ref oMissing);
            oWord.Selection.InsertBreak(ref objBreak);
            //add table

            List<List<string>> list1 = new List<List<string>>();
            list1.Add(new List<string>());
            // rownum = list index +1
            list1[0].Add("ID");
            list1[0].Add("Name");
            list1[0].Add("Last name");
            list1[0].Add("Age");
            list1[0].Add("Gender");
            list1[0].Add("Email");

            list1.Add(new List<string>());
            list1[1].Add("one");
            list1[1].Add("Arman");
            list1[1].Add("Shirzad");
            list1[1].Add("24");
            list1[1].Add("Male");
            list1[1].Add("a_shirzad76@yahoo.com");

            list1.Add(new List<string>());
            list1[2].Add("one");
            list1[2].Add("Arman");
            list1[2].Add("Shirzad");
            list1[2].Add("24");
            list1[2].Add("Male");
            list1[2].Add("a_shirzad76@yahoo.com");

            list1.Add(new List<string>());
            list1[3].Add("one");
            list1[3].Add("Arman");
            list1[3].Add("Shirzad");
            list1[3].Add("24");
            list1[3].Add("Male");
            list1[3].Add("a_shirzad76@yahoo.com");

            // with the below dictionary we say for what row we want which color to set on the table 

            Dictionary<int, string> indices = new Dictionary<int, string>();
            indices.Add(1, "blue");
            indices.Add(2, "yellow");

            table tbl = new table(4, 6, 1, 1, ref oMissing, ref oMissing, ref oDoc,
                                  ref oMissing, ref oWord, ref oTemplate, ref oEndOfDoc, list1, 1, indices);
            //Range newrng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //newrng.Select();
            //oWord.Selection.Range.ParagraphFormat.SpaceBefore = 50.0f;
            //string img = "C:\\Users\\a_shi\\Desktop\\Mapsa\\ACIDPRO\\Generate word document report\\1.jpg";
            //picture p1 = new picture(img, 500, 250, ref oDoc,
            //ref oMissing, ref oWord, ref oEndOfDoc);
            //object newrng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //string img1 = "C:\\Users\\a_shi\\Desktop\\Mapsa\\ACIDPRO\\svgsample.svg";
            //Word.Paragraph oPara = oDoc.Content.Paragraphs.Add(ref newrng1);
            //oPara.Range.ParagraphFormat.LineSpacing = 20.0f;
            //picture svg = new picture(img1, 500, 250, ref oDoc,
            //ref oMissing, ref oWord, ref oEndOfDoc);
            oDoc.Paragraphs.Add();

            #region secondtbl 

            // uncomment the body if you would like to see the second table

            //List<List<string>> list2 = new List<List<string>>();
            //list2.Add(new List<string>());
            //// r2wnum = list index +1
            //list1[0].Add("one");
            //list1[0].Add("Arman");
            //list1[0].Add("Shirzad");
            //list1[0].Add("24");
            //list1[0].Add("Male");
            //list1[0].Add("a_shirzad76@yahoo.com");

            //list1[3].Add("one");
            //list1[3].Add("Arman");
            //list1[3].Add("Shirzad");
            //list1[3].Add("24");
            //list1[3].Add("Male");
            //list1[3].Add("a_shirzad76@yahoo.com");

            //list1[3].Add("one");
            //list1[3].Add("Arman");
            //list1[3].Add("Shirzad");
            //list1[3].Add("24");
            //list1[3].Add("Male");
            //list1[3].Add("a_shirzad76@yahoo.com");


            //list1[3].Add("one");
            //list1[3].Add("Arman");
            //list1[3].Add("Shirzad");
            //list1[3].Add("24");
            //list1[3].Add("Male");
            //list1[3].Add("a_shirzad76@yahoo.com");




            //table tbl2 = new table(4, 6, 1, 1, ref oMissing, ref oMissing, ref oDoc,
            //                      ref oMissing, ref oWord, ref oTemplate, ref oEndOfDoc, list2, 2, indices);
            #endregion secondtbl

            text toctest = new text("heading 1 and its going in the table of contents why because we set style of toc to h1 and pay attention to how the last parameter is the text style which now is h1  \v", "Segoe UI (Body)", 20, ref oDoc, ref oMissing, ref oWord, ref oEndOfDoc, "", "gray", "left", 0, "h1");


            #region Toc
            TOC some = new TOC(1, 3, ref oDoc,
            ref oMissing, ref oWord, ref oTemplate, ref oEndOfDoc);

            #endregion Toc
            Word.Range ranger = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oDoc.Paragraphs.Add(ranger);
            ranger.ParagraphFormat.LineSpacing= 200f;


            // lets add a picture

            picture p1 = new picture("C:\\Users\\a_shi\\source\\repos\\interop-usermanual\\images\\thanks.jpg", 400, 200,ref oDoc,ref oMissing,ref oWord,ref oEndOfDoc);


            // in office interop is very important to release the com objects in the end

             //if (oDoc != null)
             //{
             //    oDoc.Close();
             //    Marshal.ReleaseComObject(oDoc);
             //}

             //if (oWord != null)
             //{
             //    oWord.Quit();
             //    Marshal.ReleaseComObject(oWord);
             //}
            }

        }


         //in case you would like to draw a chart in your word files here are the codes 


        //    Word.InlineShape oShape;
        //    object oFilename = "C:\\Users\\a_shi\\Desktop\\Mapsa\\ACIDPRO\\svgsample.svg";
        //    object oClassType = "MSGraph.Chart.8";
        //    object Displayasicon = "False";
        //    object oLinktofile = "True";
        //    object oIconFileName = "C:\\Users\\a_shi\\Desktop\\Mapsa\\ACIDPRO\\svgsample.svg";

        //    Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oFilename,
        //    ref Displayasicon, ref Displayasicon, ref oMissing,
        //    ref oMissing, ref oMissing,  wrdRng);
        // thank you @ArmanShirzad
    }







