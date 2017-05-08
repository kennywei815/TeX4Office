using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace TeX4Office_WindowsFormsApplication
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            EditorForm editorForm = new EditorForm();

            // Parse command line arguments
            Debug.WriteLine("args.Length = " + args.Length);
            if (args.Length >= 1)
            {
                //Debug.WriteLine("TeX4Office.exe [tex file] [png file]");
                Debug.WriteLine("TeX4Office.exe [tex file]");

                string texFilePath = args[0];

                //[TODO]: check if suffix of texFilePath is ".tex" or not
                if ( !texFilePath.EndsWith(".tex"))
                {

                    //const string message = "[ERROR]  The source TeX file should be given a .tex suffix.";
                    const string message = "[發生錯誤]  輸入的 TeX 檔必須以.tex結尾！";
                    const string caption = "TeX4Office";
                    var result = MessageBox.Show(message, caption);

                    helpMsg();
                    //return 1;
                }

                editorForm.setTexFileName(texFilePath);
                // TODO: 若檔案不存在，則開新檔，所以下面就不载入
                if (File.Exists(texFilePath))
                    editorForm.loadTeXFile(texFilePath);
            }

            // Start application
            Application.Run(editorForm);

            return 0;
        }

        static void helpMsg()
        {
            const string help_text = "Usage:  TeX4Office.exe [tex_file]";
            const string caption = "TeX4Office";
            var result = MessageBox.Show(help_text, caption);
        }
    }
}
