using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ExcelDataReader;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using ExportToExcel;


namespace Запрос_в_суды
{
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        string nowpath = "";
       // DirectoryInfo di;
        string fullName = Path.Combine(Environment.ExpandEnvironmentVariables("%temp%"), "Template.docx");
        private BackgroundWorker bw = new BackgroundWorker();
        int cou = 0;
        int cat = 0;
        int chk;
        public DataTable dt = new DataTable();
        public DataTable dt_copy = new DataTable();
        public DataTable to_excel = new DataTable();
        public DataSet ds = new DataSet();
        public DataTable finddata = new DataTable();
        public DataTable today = new DataTable();
        public DataTable yesterday = new DataTable();
        StringBuilder sb = new StringBuilder();


        string sourcefile;
        public string ExcelFilePath { get; set; } = string.Empty;
        public RadForm1()
        {
            InitializeComponent();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
            File.WriteAllText(fullName, Properties.Resources.Template, Encoding.Default);
        }

        public void addDataToExcel(string path, DataTable data)
        {

        }

        public void UniqueEx() // найти уникальные значения При изменении править тут
        {
            try
            {
                dt_copy = dt.Copy();
                to_excel = dt.Copy();
                dt_copy = dt_copy.DefaultView.ToTable(true, dt_copy.Columns[11].ColumnName.ToLower()); //distinct values from column 0
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось выделить уникальные значения " + ex);
                sb.Append(DateTime.Now + ": Не удалось выделить уникальные значения\r\n" + ex);
            }
        }

        public void FindEx(DataTable data, int y) // обработка эксель
        {
            try
            {
                //to_excel.Clear();
                finddata.Clear();
                    for (int i = 0; i < dt.Rows.Count; i++) //сбор данных по одному объекту
                    {
                        if (Convert.ToString(dt.Rows[i][11]).ToLower() == Convert.ToString(dt_copy.Rows[y][0]).ToLower())  //При изменении править тут
                        {
                            finddata.ImportRow(dt.Rows[i]);
                            dt.Rows.RemoveAt(i); //гениально!!!!
                            i--;
                        }

                }
                cou += finddata.Rows.Count;

                InsertToDocX(finddata, Convert.ToString(dt_copy.Rows[y][0]));







            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось записать файлы. Проблема с " + Convert.ToString(dt_copy.Rows[y][0]));
                sb.Append(DateTime.Now + ": Не удалось записать файлы.Наименование суда слишком длинное " + Convert.ToString(dt_copy.Rows[y][0]) + ex);
            }
        }

        public string InsertStrings(string text, string insertString, params int[] rangeLengths)
        {
            var sb1 = new StringBuilder(text);
            try
            {
                
                var indexes = new int[rangeLengths.Length];
                for (int i = 0; i < indexes.Length; i++)
                    indexes[i] = rangeLengths[i] + indexes.ElementAtOrDefault(i - 1) + insertString.Length;

                    for (int i = 0; i < indexes.Length; i++)
                    {
                        if (indexes[i] < sb1.Length)
                            sb1.Insert(indexes[i], insertString);
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось записать файлы. Проблема с форматом снилса");
                sb.Append(DateTime.Now + ": Не удалось записать файлы. Проблема с форматом снилса" + ex);
            }
            return sb1.ToString();
        }

        void InsertToDocX(DataTable finddata, string meds)
        {
            meds = meds.Replace("\n", "");
            if (Directory.Exists(@"C:\Sort-SUD\" + meds + "\\"))
            { }
            else
            {
                DirectoryInfo di = Directory.CreateDirectory(@"C:\Sort-SUD\" + meds + "\\");  
            }
            sb.Append("\r\n");
            sb.Append(DateTime.Now + ": Обработка файла\r\n");

            sb.Append(DateTime.Now + ": Убираем лишнее...\r\n");
            sb.Append(DateTime.Now + ": " + meds + " -> ");
            string name = meds.Replace("\"", "");
            sb.Append(name + "\r\n");
            // name = meds.TrimEnd('"');
            if (name.Length == 0)
            {
                name = "Не найдены";
            }

            

            string path = "";
            string snils = "";
            string fio = "";
            string ishod = "";
           // string numsud = "";

            for (int y = 0; y < finddata.Rows.Count; y++)
            {
                string sud = Convert.ToString(finddata.Rows[y].ItemArray[11]);
                if(Convert.ToString(finddata.Rows[y].ItemArray[5]).Length == 0)
                {
                    sb.Append("\r\n");
                    sb.Append(DateTime.Now + ": Регномер пуст!\r\n");
                    ishod = "";
                }
                else
                {
                    ishod = Convert.ToString(finddata.Rows[y].ItemArray[5]).Substring(17, 7);
                }
                string today = "От " + Convert.ToString(DateTime.Today).Substring(0, 10) + "   ";
                string zdate = Convert.ToString(finddata.Rows[y].ItemArray[6]);
                string suddate = Convert.ToString(finddata.Rows[y].ItemArray[10]);
                string deti = Convert.ToString(finddata.Rows[y].ItemArray[7]);
                string detirod = Convert.ToString(finddata.Rows[y].ItemArray[8]);
                string doljnik = Convert.ToString(finddata.Rows[y].ItemArray[9]);
                snils = Convert.ToString(finddata.Rows[y].ItemArray[0]);
                fio = Convert.ToString(finddata.Rows[y].ItemArray[1]) + " " + Convert.ToString(finddata.Rows[y].ItemArray[2]) + " " + Convert.ToString(finddata.Rows[y].ItemArray[3]);
             
                //  numsud = Convert.ToString(finddata.Rows[y].ItemArray[8]).Replace('/', '-') ;

                if (snils.Length == 10)
                {

                    snils = "0" + snils;
                    snils = InsertStrings(snils, "-", 2, 3);
                    snils = InsertStrings(snils, " ", 10);
                }
                else if (snils.Length == 11)
                {
                    snils = InsertStrings(snils, "-", 2, 3);
                    snils = InsertStrings(snils, " ", 10);
                }
                else
                {  }

                using (var originalDoc = WordprocessingDocument.Open(fullName, true))
                { 
                    if (meds.Length == 0)
                    {
                        if (Directory.Exists(@"C:\Sort-SUD\Нет суда" + "\\"))
                        {
                            if (File.Exists(@"C:\Sort-SUD\Нет суда" + "\\" + fio + " " +  /*numsud +*/ ".docx"))
                            {
                                path = @"C:\Sort-SUD\Нет суда" + "\\" + fio + " " + /*numsud +*/ "_1.docx";
                            }
                            else
                            {
                                path = @"C:\Sort-SUD\Нет суда" + "\\" + fio + " " + /*numsud +*/ ".docx";
                            }
                        }
                        else
                        {
                            DirectoryInfo di = Directory.CreateDirectory(@"C:\Sort-SUD\Нет суда" + "\\");
                            path = @"C:\Sort-SUD\Нет суда" + "\\" + fio + " " + /*numsud +*/ ".docx";
                        }
                    }
                    else
                    {
                        if (Directory.Exists(@"C:\Sort-SUD\" + meds + "\\"))
                        {
                            if (File.Exists(@"C:\Sort-SUD\" + meds + "\\" + fio + " " + /*numsud +*/ ".docx"))
                            {
                                path = @"C:\Sort-SUD\" + meds + "\\" + fio + " " + /*numsud +*/ "_" + y + ".docx";
                            }
                            else
                            {
                                path = @"C:\Sort-SUD\" + meds + "\\" + fio + " " + /*numsud +*/ ".docx";
                            }
                        }
                        else
                        {
                            DirectoryInfo di = Directory.CreateDirectory(@"C:\Sort-SUD\" + meds + "\\");
                            path = @"C:\Sort-SUD\" + meds + "\\" + fio + " " + /*numsud +*/ ".docx";
                        }
                        
                    }

                    var newDoc = (WordprocessingDocument)originalDoc.Clone(path, true);
                    originalDoc.Close();
                    MainDocumentPart mainPart = newDoc.MainDocumentPart;
                    var document = mainPart.Document;
                    var bookmarks = document.Body.Descendants<BookmarkStart>();

                    // Наименование судебного органа:
                    var nsud2 = bookmarks.First(bms => bms.Name == "nsud2");
                    var runnsud2 = new Run(new Text(sud));
                    nsud2.Parent.InsertAfter(runnsud2, nsud2);

                    // Наименование судебного органа:
                    var ntoday = bookmarks.First(bms => bms.Name == "today");
                    var runtoday = new Run(new Text(today));
                    ntoday.Parent.InsertAfter(runtoday, ntoday);

                    // исходящий номер:
                    var nishod = bookmarks.First(bms => bms.Name == "ishod");
                    var runnishod = new Run(new Text(ishod));
                    nishod.Parent.InsertAfter(runnishod, nishod);

                    //ФИО заявителя
                    var zayav = bookmarks.First(bms => bms.Name == "zayav");
                    var runzayav = new Run(new Text( fio + " " + snils));
                    zayav.Parent.InsertAfter(runzayav, zayav);

                    //Дата рождения заявителя
                    var zayadate = bookmarks.First(bms => bms.Name == "zayadate");
                    if (zdate != "" && zdate.Length == 10)
                    {
                        var runzayadate = new Run(new Text(zdate.Substring(0,10)));
                        zayadate.Parent.InsertAfter(runzayadate, zayadate);
                    }
                    else
                    {
                        var runzayadate = new Run(new Text(Convert.ToString("")));
                        zayadate.Parent.InsertAfter(runzayadate, zayadate);
                    }
                    
                    // Наименование судебного органа:
                    var nsud = bookmarks.First(bms => bms.Name == "nsud");
                    var runnsud = new Run(new Text(sud));
                    nsud.Parent.InsertAfter(runnsud, nsud);

                    //  Дата вынесения судебного решения, номер дела(при наличии): 
                    var datesud = bookmarks.First(bms => bms.Name == "datesud");
                    var rundatesud = new Run(new Text(suddate));
                    datesud.Parent.InsertAfter(rundatesud, datesud);

                    //ФИО, снилс несовершеннолетнего:    
                    var childfio = bookmarks.First(bms => bms.Name == "childfio");
                    var runchildfio = new Run(new Text(deti));
                    childfio.Parent.InsertAfter(runchildfio, childfio);

                    //Дата рождения несовершеннолетнего:
                    var childdate = bookmarks.First(bms => bms.Name == "childdate");
                    if (detirod != "")
                    {
                        if(detirod.Contains(","))
                        {
                           // string[] arr = detirod.Split(',');
                            var runchilddate = new Run(new Text(detirod));
                            childdate.Parent.InsertAfter(runchilddate, childdate);
                        }
                        
                        else
                        {
                            var runchilddate = new Run(new Text(detirod.Substring(0, 10)));
                            childdate.Parent.InsertAfter(runchilddate, childdate);
                        }
                    }
                    else
                    {
                        var runchilddate = new Run(new Text(Convert.ToString("")));
                        childdate.Parent.InsertAfter(runchilddate, childdate);
                    }

                    //ФИО должника на момент вынесения решения: 
                    var fiodolg = bookmarks.First(bms => bms.Name == "fiodolg");
                    var runfiodolg = new Run(new Text(doljnik));
                    fiodolg.Parent.InsertAfter(runfiodolg, fiodolg);

                    string ishod2;
                    if (Convert.ToString(to_excel.Rows[chk].ItemArray[5]).Length == 0)
                    {
                        ishod2 = "";
                    }
                    else
                    {
                        ishod2 = Convert.ToString(to_excel.Rows[chk].ItemArray[5]).Substring(17, 7);
                    }

                    to_excel.Rows[chk][12] = ishod2;
                    to_excel.Rows[chk][13] = Convert.ToString(DateTime.Today).Substring(0, 10);
                    chk++;

                 /*    DataRow workrow = to_excel.NewRow();
                     workrow[12] = ishod;
                     workrow[13] = today;
                     to_excel.Rows.Add(workrow);*/




                    sb.Append(DateTime.Now + ": Создаем файл " + fio + "\r\n");
                    cat++;
                    //newDoc.Save();
                    newDoc.Close();


                   // sb.Append(DateTime.Now + ": Скопировано строк :" + finddata.Rows.Count + "\r\n");
                }
                sb.Append(DateTime.Now + ": Скопировано строк :" + cat + "\r\n");
            }
        }

        public void CheckDir()
        {
            string aaa = Convert.ToString(DateTime.Today).Substring(0, 10);
            if (Directory.Exists(@"C:\Sort-SUD\"))
            { }
            else
            {
                DirectoryInfo di = Directory.CreateDirectory(@"C:\Sort-SUD\");
            }
        }

        public DataTable GetTableDataFromXl(string path, bool hasHeader = true)
        {
            dt.Clear();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // var result = reader.AsDataSet();
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        // Gets or sets a value indicating whether to set the DataColumn.DataType 
                        // property in a second pass.
                        UseColumnDataType = true,

                        // Gets or sets a callback to determine whether to include the current sheet
                        // in the DataSet. Called once per sheet before ConfigureDataTable.
                        FilterSheet = (tableReader, sheetIndex) => true,

                        // Gets or sets a callback to obtain configuration options for a DataTable. 
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            // Gets or sets a value indicating the prefix of generated column names.
                            // EmptyColumnNamePrefix = "Column",

                            // Gets or sets a value indicating whether to use a row from the 
                            // data as column names.
                            UseHeaderRow = true,

                            // Gets or sets a callback to determine which row is the header row. 
                            // Only called when UseHeaderRow = true.
                            /* ReadHeaderRow = (rowReader) => {
                                 // F.ex skip the first row and use the 2nd row as column headers:
                                 rowReader.Read();
                             },*/

                            // Gets or sets a callback to determine whether to include the 
                            // current row in the DataTable.
                            FilterRow = (rowReader) => {
                                return true;
                            },

                            // Gets or sets a callback to determine whether to include the specific
                            // column in the DataTable. Called once per column after reading the 
                            // headers.
                            FilterColumn = (rowReader, columnIndex) => {
                                return true;
                            }
                        }
                    });

                    // The result of each spreadsheet is in result.Tables
                    dt = result.Tables[0];

                }
            }
            return dt;
        }


        #region Кнопки
        private void radButton5_Click(object sender, EventArgs e)
        {
            if (bw.IsBusy != true)
            {
                bw.RunWorkerAsync();
            }
        }
        private void radButton6_Click(object sender, EventArgs e)
        {
            if (bw.WorkerSupportsCancellation == true)
            {
                bw.CancelAsync();
            }
        }
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            CheckDir();
            BackgroundWorker worker = sender as BackgroundWorker;

            chk = 0;
            cou = 0;
            cat = 0;
            finddata = dt.Clone();
            for (int y = 0; y < dt_copy.Rows.Count; y++)
            {
                if ((worker.CancellationPending == true))
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    int percentage = (y + 1) * 100 / dt_copy.Rows.Count;
                    FindEx(finddata, y);
                    worker.ReportProgress(percentage);
                }
                File.AppendAllText(@"C:\Sort-SUD\log.txt", sb.ToString());
                sb.Clear();
            }
            ds.Tables.Add(to_excel);
            CreateExcelFile.CreateExcelDocument(ds, ExcelFilePath);


            finddata.Dispose();
        }
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((e.Cancelled == true))
            {
                progressBar1.Text = "Отменено!";
                radRichTextEditor1.Text += "Отменено!\n";
                sb.Append("\r\n");
                sb.Append(DateTime.Now + ": Отменено!\r\n");
            }

            else if (!(e.Error == null))
            {
                progressBar1.Text = ("Ошибка: " + e.Error.Message);
                radRichTextEditor1.Text += "Ошибка: " + e.Error.Message + "\n";
                sb.Append("\r\n");
                sb.Append(DateTime.Now + ": Ошибка: " + e.Error.Message + "\r\n");
            }

            else
            {
                progressBar1.Text = "Выполнено!";
                radRichTextEditor1.Text += "Выполнено!\n";
                sb.Append("\r\n");
                sb.Append(DateTime.Now + ": Выполнено!\r\n");
            }
            radRichTextEditor1.Text += "Обработано записей в файле: " + cou + "\n";
            sb.Append(DateTime.Now + ": Обработано записей в файле: " + cou + "\r\n");
            radRichTextEditor1.Text += "Создано файлов: " + cat + "\n";
            sb.Append(DateTime.Now + ": Создано файлов :" + cat + "\r\n");

            File.AppendAllText(@"C:\Sort-SUD\log.txt", sb.ToString());
            sb.Clear();
        }
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value1 = e.ProgressPercentage;
            progressBar1.Text = (e.ProgressPercentage.ToString() + "%");
        }
        private void radButton2_Click(object sender, EventArgs e)
        {
                sb.Append("\r\n");
                sb.Append("\r\n");
                sb.Append("------------------------ " + DateTime.Now + " ------------------------\r\n");
               // nowpath = @"C:\Sort-SUD\Беременные_" + Convert.ToString(DateTime.Now).Substring(0, 16).Replace(":", "-") + "\\";
                sb.Append("Будет создана папка: " + nowpath + "\r\n");
                CheckDir(); //проверяем папки

                string strlen = "";
                OpenFileDialog fbd = new OpenFileDialog();
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    sourcefile = Path.GetFileNameWithoutExtension(fbd.SafeFileName);
                    ExcelFilePath = fbd.FileName;
                    radRichTextEditor1.Text += "Выбран файл: " + fbd.FileName + "\n";
                    sb.Append(DateTime.Now + ": Выбран файл: " + fbd.FileName + "\r\n");
                    string Ext1 = Path.GetExtension(ExcelFilePath);
                    if (Ext1 == ".xls" || Ext1 == ".xlsx")
                    {
                        radRichTextEditor1.Text += "Файл успешно открыт\n";
                        sb.Append(DateTime.Now + ": Файл успешно открыт\r\n");
                        radRichTextEditor1.Text += "Обработка файла, подождите...\n";

                        GetTableDataFromXl(fbd.FileName);
                        cou = dt.Rows.Count;
                        strlen = dt.Rows[1].ItemArray[0].ToString();

                        radRichTextEditor1.Text += "Обнаружено записей в файле: " + dt.Rows.Count + "\n";
                        sb.Append(DateTime.Now + ": Обнаружено записей в файле: " + dt.Rows.Count + "\r\n");

                        UniqueEx();
                        radRichTextEditor1.Text += "Обнаружено учреждений в файле: " + dt_copy.Rows.Count + "\n";
                        sb.Append(DateTime.Now + ": ООбнаружено учреждений в файле: " + dt_copy.Rows.Count + "\r\n");
                        radRichTextEditor1.Text += "Нажмите кнопку Начать\n";
                    }
                    else
                    {
                        radRichTextEditor1.Text += "Не удалось открыть файл. Это не файл MS Excel!" + "\n";
                        sb.Append(DateTime.Now + ": Не удалось открыть файл.Это не файл MS Excel!\r\n");
                    }
                }
                File.AppendAllText(@"C:\Sort-SUD\\log.txt", sb.ToString());
                sb.Clear();
            }
        private void radButton4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void radButton1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", "C:\\Sort-SUD");
        }
        #endregion


        public void CompareXLS(DataTable _new, DataTable old)
        {
            try
            {
                finddata.Clear();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (Convert.ToString(dt.Rows[i][10]) == Convert.ToString(dt_copy.Rows[0][0])) // откуда - куда
                    {
                        finddata.ImportRow(dt.Rows[i]);
                        dt.Rows.RemoveAt(i); //гениально!!!!
                        i--;
                    }
                }
                cou += finddata.Rows.Count;
                // InsertToDocX(Convert.ToString(dt_copy.Rows[y][0]), y);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось записать файлы " + ex);
                sb.Append(DateTime.Now + ": ООбнаружено учреждений в файле: " + dt_copy.Rows.Count + "\r\n");
            }
        }
    }
}

