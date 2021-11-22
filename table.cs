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
    class table
    {
        public int NumRows { get; set; }
        public int NumColumns { get; set; }
        public List<List<string>> Data { get; set; }
        public int TableIndex { get; set; }
        public table(int numRows, int numColumns,int border, int boldedrowindex ,ref object defaultTableBehavior, ref object autoFitBehavior, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oTemplate, ref object oEndOfDoc, List<List<string>> data, int tableIndex, Dictionary <int,string> index,string align = "center")
        {
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, numRows, numColumns, ref oMissing, ref oMissing);
            //new features

            oTable.Rows[boldedrowindex].Range.Font.Bold = 1;
            oTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);


            foreach (int i in index.Keys)
            {


            
            switch (index[i])      //color items can be modified according to AcidPro needs
            {
                case "blue":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue;
                    break;

                case "green":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGreen;
                    break;
                case "teal":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorTeal;
                    break;
               
                case "violet":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorViolet;
                    break;
                case "red":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorRed;
                    break;
                case "yellow":
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow;
                    break;
                default:
                    oTable.Rows[i].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                    break;
            }
            }
         
            //oTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //oTable.Columns[6].Cells.PreferredWidth = 125;
            //oTable.Rows[2].Height = 125;
            if (border == 1) { 
           oTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
           oTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
           oTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
           oTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
           oTable.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
           oTable.Borders[Word.WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }
            if (align == "left") { oTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; }
            else if (align == "right") { oTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; }
            else
            {
                oTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            //
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[0].Count; j++)
                {
                    oDoc.Tables[tableIndex+1].Cell(i + 1, j + 1).Range.Text = data[i][j];
                }
            }
            oTable.Range.ParagraphFormat.SpaceAfter = 10;
         

        }

    }
}
