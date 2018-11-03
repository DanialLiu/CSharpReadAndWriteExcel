using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Reflection;
using Path = System.IO.Path;
using System.IO;

namespace ExcelSplit
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private static Random random = new Random();
        public delegate void onProcessEvent(string item);
        public delegate void onFinishEvent();
        public event onProcessEvent onProcess;
        public event onFinishEvent onFinish;
        public Thread thread;
        public MainWindow()
        {
            InitializeComponent();
            //onProcess += updateListItem;
            //onFinish += onSplitFinish;
        }        
        private string genPassword(int pass_len)
        {
            const string chars = "abcdefghijksmnprstuvwxyz123456789";
            return new string(Enumerable.Repeat(chars, pass_len)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
        private void button_Click(object sender, RoutedEventArgs e)
        {
            buttonSplit.IsEnabled = false;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件|*.xls*|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == true)
            {
                labelStatus.Content = "loading...";
                int pass_len = int.Parse(textBoxPassLen.Text);
                if (pass_len <= 0 || pass_len > 100)
                    pass_len = 6;
                thread = new Thread(() => run(openFileDialog.FileName, pass_len));
                thread.Start();
            }

        }
        //private void onProcessItem(string item)
        //{
        //    //if (this.textBox1.InvokeRequired)
        //    //{
        //    //    StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(SetText);
        //    //    this.Invoke(d, new object[] { text });
        //    //}
        //    //else
        //    //{
        //    //    this.textBox1.Text = text;
        //    //}
        //    private delegate void SetTextCallback(System.Windows.Controls.TextBox control, string text);
        //labelStatus.Dispatcher.Invoke(d, new object[] { control, text });
        //    //this.Invoke(new Action(() => updateListItem(item)));
        //    //listBoxEmployee.Items.Insert(0, item);
        //}
        //private void updateListItem(string item)
        //{
        //    listBoxEmployee.Items.Insert(0, item);
        //}

        //private void onSplitFinish()
        //{
        //    labelStatus.Content = "done!";
        //    buttonSplit.IsEnabled = true;
        //}
        private void run(string excel_path, int pass_len)
        {
            string work_path = Path.GetDirectoryName(excel_path);
            string split_folder = Path.Combine(work_path, "split_encrypt");
            if (!Directory.Exists(split_folder))
            {
                Directory.CreateDirectory(split_folder);
            }            
            Excel.Application xApp = new Excel.Application();

            xApp.Visible = false;
            xApp.DisplayAlerts = false;
            xApp.AlertBeforeOverwriting = false;
            Excel.Workbook xBook = xApp.Workbooks._Open(excel_path);
            Excel.Worksheet xSheet = (Excel.Worksheet)xBook.Sheets[1];

            Excel.Workbook passwordBook;
            Excel.Worksheet passwordSheet;
            var passwordPath = Path.Combine(work_path, "passwords.xlsx");
            try
            {
                passwordBook = xApp.Workbooks._Open(passwordPath);
                passwordSheet = (Excel.Worksheet)passwordBook.Sheets[1];
            }
            catch
            {
                passwordBook = xApp.Workbooks.Add();
                passwordSheet = (Excel.Worksheet)passwordBook.Sheets[1];
                passwordSheet.Cells[1, 1] = "employee";
                passwordSheet.Cells[1, 2] = "password";
            }
            var passwords = new Dictionary<String, String>();
            var new_passwords = new Dictionary<String, String>();
            for (int i = 2; i <= passwordSheet.UsedRange.Rows.Count; i++)
            {
                passwords.Add(passwordSheet.Rows[i].Columns[1].Text, passwordSheet.Rows[i].Columns[2].Text);
            }
            Excel.Range head = xSheet.Rows[1];
            for (int i = 2; i <= xSheet.UsedRange.Rows.Count; i++)
            {
                var name = xSheet.Cells[i, 2].Text;
                var id = xSheet.Cells[i, 1].Text;
                var filename = id + '-' + name;
                Application.Current.Dispatcher.Invoke(
                () =>
                {
                    labelStatus.Content = "spliting:" + filename;
                });
                //onProcess(filename);
                string password;
                if (passwords.ContainsKey(filename))
                {
                    password = passwords[filename];
                }
                else
                {
                    password = genPassword(pass_len);
                    new_passwords[filename] = password;
                }

                

                var destname = filename + ".xlsx";
                Excel.Workbook destworkBook = xApp.Workbooks.Add();
                Excel.Worksheet destworkSheet = destworkBook.Worksheets.Add();

                Excel.Range to = destworkSheet.Rows[1];
                head.Copy(to);
                xSheet.Rows[i].Copy(destworkSheet.Rows[2]);
                destworkSheet.Columns.AutoFit();
                destworkBook.SaveAs(Path.Combine(split_folder, destname), Missing.Value, password);
                destworkBook.Close();
            }
            if (new_passwords.Count > 0)
            {
                int insertIndex = passwordSheet.UsedRange.Rows.Count + 1;
                foreach (KeyValuePair<string, string> kvp in new_passwords)
                {
                    passwordSheet.Cells[insertIndex, 1] = kvp.Key;
                    passwordSheet.Cells[insertIndex, 2] = kvp.Value;
                    insertIndex++;
                }
                passwordSheet.Columns.AutoFit();
                passwordBook.SaveAs(passwordPath, Missing.Value);
                passwordBook.Close();
            }
            xBook.Close();
            Application.Current.Dispatcher.Invoke(
            () =>
            {
                labelStatus.Content = "done! saved in folder: split_encrypt";
                buttonSplit.IsEnabled = true;
            });
        }
    }
}
