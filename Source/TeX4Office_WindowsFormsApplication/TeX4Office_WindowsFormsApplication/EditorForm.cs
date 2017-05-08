using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Runtime;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace TeX4Office_WindowsFormsApplication
{
    public partial class EditorForm : Form
    {
        [DataContract]
        public class Config
        {
            [DataMember]
            public string engine;

            [DataMember]
            public string dpi;

            public Config(string engine = "PDFLaTeX", string dpi = "600")
            {
                this.engine = engine;
                this.dpi = dpi;
            }

            public static Config load(string path)
            {
                Config newConfig = new Config();

                try
                {
                    using (FileStream Config_file = new FileStream(path, FileMode.Open))
                    {
                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Config));
                        newConfig = (Config)ser.ReadObject(Config_file);
                    }
                }
                catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "載入設定檔時發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }

                return newConfig;
            }

            public void save(string path)
            {
                try
                {
                    using (FileStream Config_file = new FileStream(path, FileMode.Create))
                    {
                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Config));
                        ser.WriteObject(Config_file, this);
                    }
                }
                catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "儲存設定檔時發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        string startUpPath;
        string texFileBaseName;
        string texFileDir;
        string texCode;
        string configFilePath;

        Config config;

        public EditorForm()
        {
            InitializeComponent();

            // Set initial work dir & file name
            startUpPath = System.Windows.Forms.Application.StartupPath;

            texFileDir = "C:\\Temp";  // [TODO]: need to be portable to Mac OS X
            texFileBaseName = "tex_file";
            this.ribbonTextBox1.TextBoxText = texFileBaseName;

            if ( ! Directory.Exists(texFileDir) )
            {
                Directory.CreateDirectory(texFileDir);
            }

            configFilePath = Path.Combine(startUpPath, "config.json");
            if ( File.Exists(configFilePath) )
            {
                config = Config.load(configFilePath);
            }
            else
            {
                config = new Config();
            }

            engineComboBox1.TextBoxText = config.engine;
            dpiComboBox1.TextBoxText = config.dpi;
        }

        public void setTexFileName(string path)
        {
            // Get base name of the tex file
            texFileDir = Path.GetDirectoryName(path);  // [TODO]: need to be portable to Mac OS X
            texFileBaseName = Path.GetFileNameWithoutExtension(path);
            this.ribbonTextBox1.TextBoxText = texFileBaseName;
        }

        public void setTexFileContent(string content)
        {
            this.richTextBox1.Text = content;
        }

        public void loadTeXFile(string path)
        {
            try
            {
                using (StreamReader tex_file = new StreamReader(path))
                {
                    this.setTexFileContent(tex_file.ReadToEnd());
                }
            }
            catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "載入發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        public void saveTeXFile(string path)
        {
            try
            {
                using (StreamWriter tex_file = new StreamWriter(path))
                {
                    tex_file.Write(this.richTextBox1.Text);
                }
            }
            catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "儲存發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void output_botton_click(object sender, EventArgs e)
        {
            //[DONE] TeX4Office Editor要可以載入、存檔
            //[DONE] TeX4Office Editor要可以選引擎、DPI
            //[DONE] 修lualatex、xelatex的template到可以用中文、萬國碼
            //[DONE] 都換成MessageBox.Show(Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error)

            //TODO: 加上足夠Try...Catch...Finally錯誤處理，加強錯誤處理

            //TODO: 開始output時應該要顯示一個彈跳方塊，顯示目前狀態；如果Time out或有錯誤，也要告訴使用者

            //TODO: 輸出加上純輸出、輸出並關閉的選項

            //TODO: 實作Help => 先用方塊or直接寫說明檔???

            //TODO: 更多template

            //TODO: 國際化功能

            //TODO: 可以用任意外部程式執行TeX檔 => midi、電路模擬

            //-----------------------------------------------------------------------------------------------------------------------------
            // Step 0a: Check if the file name and the code are valid, and then set up the names of corresponding output files.
            //-----------------------------------------------------------------------------------------------------------------------------
            config.engine = engineComboBox1.TextBoxText;
            config.dpi = dpiComboBox1.TextBoxText;
            config.save(configFilePath);

            string engine = engineComboBox1.TextBoxText.ToLower();
            int dpi = Convert.ToInt32(dpiComboBox1.TextBoxText.ToLower());

            texFileBaseName = this.ribbonTextBox1.TextBoxText;
            texCode = this.richTextBox1.Text;

            Debug.WriteLine("texFileDir = " + texFileDir);
            Debug.WriteLine("texFileBaseName = " + texFileBaseName);
            Debug.WriteLine("texCode = " + texCode);

            // [TODO] 在輸入格中提示 "請命名"
            // [TODO] 英文版訊息： "[ERROR]  The source TeX file should be given a name."
            {
                bool error = false;

                if (String.IsNullOrWhiteSpace(texFileBaseName)) { MessageBox.Show("請先命名，再按輸出！", "輸出發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); error = true; }
                if (String.IsNullOrWhiteSpace(texCode)) { MessageBox.Show("輸入區未包含程式碼！", "輸出發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); error = true; }
                if (error) return;
            }

            string texFilePath = Path.Combine(texFileDir, texFileBaseName + ".tex");
            string dviFilePath = Path.Combine(texFileDir, texFileBaseName + ".dvi");
            string pdfFilePath = Path.Combine(texFileDir, texFileBaseName + ".pdf");
            string pngFilePath = Path.Combine(texFileDir, texFileBaseName + ".png");

            Debug.WriteLine("texFilePath = " + texFilePath);

            this.saveTeXFile(texFilePath);

            //-----------------------------------------------------------------------------------------------------------------------------
            // Step 0b: Prepare to run external programs in child process and redirect the output stream of the child process.
            //-----------------------------------------------------------------------------------------------------------------------------
            const int timeout_latex = 150 * 1000;
            const int timeout_toPNG = 300 * 1000;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;

            //-----------------------------------------------------------------------------------------------------------------------------
            // Step1: Run LaTeX engine in the child process.
            //-----------------------------------------------------------------------------------------------------------------------------
            if (engine == "pdflatex")
            {
                startInfo.FileName = "pdflatex";
                startInfo.Arguments = "-halt-on-error -shell-escape -output-format=dvi -output-directory=" + texFileDir + " " + texFilePath;
            }
            else if (engine == "lualatex")
            {
                startInfo.FileName = "lualatex";
                startInfo.Arguments = "-halt-on-error -shell-escape -output-format=pdf -output-directory=" + texFileDir + " " + texFilePath; //dvi字型不會載入
            }
            else if (engine == "xelatex")
            {
                startInfo.FileName = "xelatex";
                startInfo.Arguments = "-halt-on-error -shell-escape -output-directory=" + texFileDir + " " + texFilePath;  //沒有-output-format=dvi ，只有--no-pdf用來輸出.XDV(extended dvi)檔
            }

            try
            {
                Process p = new Process();
                p.StartInfo = startInfo;
                p.EnableRaisingEvents = true;
                p.Start();

                // Read the output stream first and then wait.
                // [NOTE] Do not wait for the child process to exit before reading to the end of its redirected stream.

                string output = p.StandardOutput.ReadToEnd();
                string error = p.StandardError.ReadToEnd();

                bool result = p.WaitForExit(timeout_latex);

                if (!result)              { MessageBox.Show("LaTeX 編譯已逾時\n\n錯誤訊息：\n" + output + "\n" + error,   "LaTeX 編譯發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                else if (p.ExitCode != 0) { MessageBox.Show("LaTeX 編譯發生錯誤\n\n錯誤訊息：\n" + output + "\n" + error, "LaTeX 編譯發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                //debug
                Debug.WriteLine("p.startInfo.FileName = " + startInfo.FileName);
                Debug.WriteLine("p.StartInfo.Arguments = " + p.StartInfo.Arguments);
                Debug.WriteLine("p.StandardOutput.ReadToEnd() = " + output);
                Debug.WriteLine("p.StandardError.ReadToEnd() = " + error);
            }
            catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "LaTeX 編譯發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }


            //-----------------------------------------------------------------------------------------------------------------------------
            // Step2: Convert DVI or PDF to PNG in the child process.
            //-----------------------------------------------------------------------------------------------------------------------------
            if (engine == "pdflatex")
            {
                startInfo.FileName = "dvipng.exe";
                startInfo.Arguments = "-q -D " + dpi + " -bg Transparent -T tight  " + dviFilePath + "  -o  " + pngFilePath;
            }
            else
            {
                startInfo.FileName = Path.Combine(startUpPath, "convert.exe"); //[NOTE] 需要自帶convert，不然會和DOS內建指令衝突  TODO: 在別的機器上試試
                startInfo.Arguments = " -units PixelsPerInch  -density " + dpi + " -trim " + dviFilePath + "  " + pngFilePath; //TODO: dpi*4 才和dvipng的解析度相同
            }

            try
            {
                Process p = new Process();
                p.StartInfo = startInfo;
                p.EnableRaisingEvents = true;
                p.Start();

                // Read the output stream first and then wait.
                // [NOTE] Do not wait for the child process to exit before reading to the end of its redirected stream.
                string output = p.StandardOutput.ReadToEnd();
                string error = p.StandardError.ReadToEnd();

                bool result = p.WaitForExit(timeout_toPNG);

                if (!result)              { MessageBox.Show("PNG 轉換已逾時\n\n錯誤訊息：\n" + output + "\n" + error,   "PNG 轉換發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                else if (p.ExitCode != 0) { MessageBox.Show("PNG 轉換發生錯誤\n\n錯誤訊息：\n" + output + "\n" + error, "PNG 轉換發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                Debug.WriteLine("p.startInfo.FileName = " + startInfo.FileName);
                Debug.WriteLine("p.StartInfo.Arguments = " + p.StartInfo.Arguments);
                Debug.WriteLine("p.StandardOutput.ReadToEnd() = " + output);
                Debug.WriteLine("p.StandardError.ReadToEnd() = " + error);
            }
            catch (Exception exception) { MessageBox.Show(exception.Message + "\n\n" + exception.ToString(), "PNG 轉換發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }


            //-----------------------------------------------------------------------------------------------------------------------------
            // Step3: Close everything down.
            //-----------------------------------------------------------------------------------------------------------------------------
            Application.Exit();
        }

        private void save_botton_click(object sender4, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = Path.Combine(startUpPath, "templates\\");
            saveFileDialog1.Filter = "TeX files (*.tex)|*.tex|All files (*.*)|*.*";
            saveFileDialog1.Title = "儲存至範本";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                this.saveTeXFile(saveFileDialog1.FileName);
            }
        }

        private void load_botton_click(object sender4, EventArgs e)
        {
            //[DONE] 應放在 "%安裝資料夾%\templates" 下
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Path.Combine(startUpPath, "templates\\");
            openFileDialog1.Filter = "TeX files (*.tex)|*.tex|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "從範本載入";

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.loadTeXFile(openFileDialog1.FileName);
            }
        }

        private void help_botton_click(object sender4, EventArgs e)
        {
            //TODO
        }

        private void ribbon_main_botton_click(object sender, EventArgs e)
        {
            // We'll never enter this function.
        }
    }
}
