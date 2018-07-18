using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;

namespace BiGPay
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            Application.Run(new Form1());
            
            //Chemins des fichiers
            //string path = System.IO.Path.GetDirectoryName(SelectPdf.FileName);
            //string pdfFile = SelectPdf.FileName;
            //string nameFile = System.IO.Path.GetFileNameWithoutExtension(SelectPdf.FileName);
            //string wordFile = path + @"\" + nameFile + @".docx";
            //string excelFile = path + @"\" + nameFile + @".xlsx";




        }
    }
}
