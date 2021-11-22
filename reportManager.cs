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
    class reportManager
    {
        static void Main(string[] args)
        {
            generator obj = new generator();
            System.Console.ReadKey();
        }
    }
}
