using DocumentFormat.OpenXml.Drawing.Charts;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Autograding
{
    public partial class Form1 : Form
    {
        enum Anser : int
        {
            ColumnName,
            Name,
            SrcName,
            Input,
            Correct0,
            Correct1,
            SrcCode,
            Incorrect,
        }

        List<List<string>> anserList = new List<List<string>>();
        List<string> userList = new List<string>();

        public Form1()
        {
            InitializeComponent();

            const string fileName = "answer.xlsx";
            try
            {
                using var ans = MyExcelBook.Open(fileName);
                ans.SelectSheet("Sheet1");

                foreach (var row in ans.GetAllCellValues())
                {
                    anserList.Add(new List<string>());

                    if (row.Length != 0)
                    {
                        anserList[anserList.Count - 1].Add(IntExtensions.ToColumnName(anserList.Count + 1));

                        foreach (var cell in row)
                        {
                            anserList[anserList.Count - 1].Add(cell.Value);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void exec(string[] files)
        {

            MyExcelBook book = null;
            var is_fopen = files.Length > 1;
            //var is_fopen = true;

            if (is_fopen)
            {
                userList.Clear();

                book = MyExcelBook.CreateBook("result.xlsx");
                book.CreateSheet("Sheet1");

                var index = 1;
                foreach (var t in anserList)
                {
                    book.SetValue(t[(int)Anser.Name] + "_" + t[(int)Anser.SrcName], t[(int)Anser.ColumnName], 1);
                }
            }

            for (int i = 0; i < files.Length; i++)
            {
                var fileName = files[i];
                textBox2.Text += fileName + System.Environment.NewLine;

                FileInfo fi = new FileInfo(fileName);

                string user = fi.Directory.Name;

                string exe = System.IO.Path.GetFileNameWithoutExtension(fi.Name) + ".exe";

                List<string> result = anserList.Find(n => fi.FullName.Contains(n[(int)Anser.Name]) && fi.FullName.Contains(n[(int)Anser.SrcName]));

                if (result == null)
                {
                    continue;
                }

                if (result.Count > (int)Anser.SrcCode && result[(int)Anser.SrcCode] != null && result[(int)Anser.SrcCode].Length != 0)
                {
                    StreamReader sr = fi.OpenText();
                    string src_code = sr.ReadToEnd();
                    sr.Close();

                    src_code = Regex.Replace(src_code, "\r\n", "\n");
                    src_code = Regex.Replace(src_code, @"\n+$", "");

                    if (src_code == result[(int)Anser.SrcCode])
                    {
                        continue;
                    }
                }

                string strout = "";
                string strerr = "";
                string match = "×";

                int exitCode;

                using (Process myProcess = new Process())
                {
                    myProcess.StartInfo.FileName = "g++";
                    //myProcess.StartInfo.Arguments = "-fdiagnostics-color=always -finput-charset=UTF-8 -g -Wall -Wextra -Wshadow -Wconversion -o " + exe + " " + fi.Name;
                    myProcess.StartInfo.Arguments = "-finput-charset=UTF-8 -g -Wall -Wextra -Wshadow -Wconversion -o " + exe + " " + fi.Name;
                    myProcess.StartInfo.WorkingDirectory = fi.DirectoryName;
                    myProcess.StartInfo.CreateNoWindow = true; // コンソール・ウィンドウを開かない
                    myProcess.StartInfo.UseShellExecute = false; // シェル機能を使用しない
                    myProcess.StartInfo.RedirectStandardOutput = true; // 標準出力をリダイレクト
                    myProcess.StartInfo.RedirectStandardError = true;
                    myProcess.StartInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                    myProcess.StartInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                    myProcess.Start();

                    myProcess.WaitForExit(10000);

                    {
                        string s = myProcess.StandardError.ReadToEnd();
                        if (s != null && s.Length != 0)
                        {
                            strerr += s;
                        }
                    }

                    try
                    {
                        exitCode = myProcess.ExitCode;
                    }
                    catch
                    {
                        exitCode = -1;
                        myProcess.Kill();
                    }
                }

                if (exitCode == 0)
                {
                    using (Process myProcess = new Process())
                    {
                        myProcess.StartInfo.FileName = fi.DirectoryName + "\\" + exe;
                        myProcess.StartInfo.WorkingDirectory = fi.DirectoryName;
                        myProcess.StartInfo.CreateNoWindow = true; // コンソール・ウィンドウを開かない
                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.StartInfo.RedirectStandardInput = true;
                        myProcess.StartInfo.RedirectStandardOutput = true; // 標準出力をリダイレクト
                        myProcess.StartInfo.RedirectStandardError = true;
                        myProcess.StartInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                        myProcess.StartInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                        myProcess.Start();

                        using (StreamWriter myStreamWriter = myProcess.StandardInput)
                        {
                            if (result[(int)Anser.Input] != null)
                            {
                                myStreamWriter.WriteLine(result[(int)Anser.Input]);
                            }
                            //myStreamWriter.WriteLine("");
                            //myStreamWriter.WriteLine("");
                            //myStreamWriter.WriteLine("");
                            myStreamWriter.Close();
                        }

                        myProcess.WaitForExit(10000);

                        try
                        {
                            exitCode = myProcess.ExitCode;

                            match = "〇";
                            {
                                string s = myProcess.StandardOutput.ReadToEnd();
                                if (s != null && s.Length != 0)
                                {
                                    string tmp = Regex.Replace(s, "\r\n", "\n");
                                    tmp = Regex.Replace(tmp, @"\n+$", "");
                                    if (tmp.Length == 0)
                                    {
                                        match = "×";
                                    }
                                    //else if (result.Count > (int)Anser.Incorrect && result[(int)Anser.Incorrect] != null && result[(int)Anser.Incorrect].Length != 0 && tmp == result[(int)Anser.Incorrect])
                                    //{
                                    //    match = "×";
                                    //}
                                    else if (
                                        (result.Count > (int)Anser.Correct0 && result[(int)Anser.Correct0] != null && result[(int)Anser.Correct0].Length != 0 && tmp == result[(int)Anser.Correct0]) ||
                                        (result.Count > (int)Anser.Correct1 && result[(int)Anser.Correct1] != null && result[(int)Anser.Correct1].Length != 0 && tmp == result[(int)Anser.Correct1]))
                                    {
                                        //match = "〇";
                                    }
                                    else
                                    {
                                        match = "△";
                                    }
                                    strout += s;
                                }
                                else
                                {
                                    match = "×";
                                }
                            }
                            {
                                string s = myProcess.StandardError.ReadToEnd();
                                if (s != null && s.Length != 0)
                                {
                                    strerr += s;
                                    match = "×";
                                }
                            }
                        }
                        catch
                        {
                            exitCode = -1;
                            myProcess.Kill();
                        }

                    }
                }

                textBox1.Text += user + " , " + result[(int)Anser.Name] + " , " + match + System.Environment.NewLine;

                textBox2.Text += strout + System.Environment.NewLine;
                textBox2.Text += strerr + System.Environment.NewLine;

                if (is_fopen)
                {
                    int idx = userList.IndexOf(user);

                    if (idx < 0)
                    {
                        idx = userList.Count;
                        userList.Add(user);

                        book.SetValue(user, "A", (uint)idx + 2);
                    }

                    book.SetValue(match, result[(int)Anser.ColumnName], (uint)idx + 2);
                }
            }

            if (is_fopen)
            {
                book.Save();
            }

        }


        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            exec(files);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "フォルダを指定してください。";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.SelectedPath = @"C:\Windows";
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
//                Console.WriteLine(fbd.SelectedPath);
                DirectoryInfo di = new DirectoryInfo(fbd.SelectedPath);
                if (di.Exists)
                {
                   FileInfo[] fi_array = di.GetFiles("*.c", SearchOption.AllDirectories);
                    List<string> files = new List<string>();

                    foreach (FileInfo fi in fi_array)
                    {
                        files.Add(fi.FullName);
                    }
                    exec(files.ToArray());
                }
            }
        }
    }
}