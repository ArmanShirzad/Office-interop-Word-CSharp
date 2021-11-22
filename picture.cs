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
    class picture
    {
        public string Imgpath { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public int Halign { get; set; }
        public int Valign { get; set; }
        public picture(string imgpath, int width, int height, ref Word._Document oDoc,
            ref object oMissing, ref Word._Application oWord, ref object oEndOfDoc, string halign = "center", string valign = "top")
        {
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Word.InlineShape pictureShape = wrdRng.InlineShapes.AddPicture(imgpath);
            pictureShape.Width = width;
            pictureShape.Height = height;
            Word.Shape shape1 = pictureShape.ConvertToShape();
           
            shape1.WrapFormat.Type = Word.WdWrapType.wdWrapInline;
            switch (halign)
            {
                
                case "left":
                    shape1.Left = (float)Word.WdShapePosition.wdShapeLeft;
                    break;

                case "right":
                    shape1.Left = (float)Word.WdShapePosition.wdShapeRight;
                    break;
                default:
                    shape1.Left = (float)Word.WdShapePosition.wdShapeCenter;
                    break;
            }
            if (valign == "bottom")
            {
                shape1.Top = (float)Word.WdShapePosition.wdShapeBottom;

            }
            else
            {
                shape1.Top = (float)Word.WdShapePosition.wdShapeTop;

            }


        }

    }
}
