using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Doc_pdf_generator
{
    public class MainViewModel : ViewModel
    {
        private string _rayon = "";
        private string _adres = "";
        private string _FIO = "";
        private string _proectirovshic = "";
        private string _pasport = "";

        public string Rayon { get { return _rayon; } set { _rayon = value; OnPropertyChanged(); } }
        public string Adres { get { return _adres; } set { _adres = value; OnPropertyChanged(); } }
        public string Fio { get { return _FIO; } set { _FIO = value; OnPropertyChanged(); } }
        public string Proectirovshic { get {  return _proectirovshic; }
            set
            {
                _proectirovshic = value;
                OnPropertyChanged();
            } }
        public string Pasport { get { return _pasport; }
            set
            {
                _pasport = value;
                OnPropertyChanged();
            } }

        private string _selectVid = "Кондиционер";

        public string SelectVid { get { return _selectVid; } set { _selectVid = value; OnPropertyChanged(); } }

        private List<string> _list = new List<string>()
        {
            "Кондиционер",
             "Декоративный экран",
              "Вентиляционная решетка",
               "Вентиляционный трубопровод",
                "Роллеты",
        };

        public List<string> List { get { return _list; } }


        public MainViewModel()
        {
          
        }
        List<System.Drawing.Image> printList = new List<System.Drawing.Image>();
        public ICommand SaveCommand
        {
            get
            {
                return new RelayCommand((obj) =>
                {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Сохранить документы?", "Сохранить", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        // Создаём объект документа
                        Word.Document doc = null;
                        Word.Document doc2 = null;
                        try
                        {
                            doc = SaveOne();
                            doc2 = SaveTwo();


                            string SelectedPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\";
                            SelectedPath += !Adres.Equals("") ? Adres : "Новый адрес";
                            DirectoryInfo dirInfo = new DirectoryInfo(SelectedPath);
                            if (!dirInfo.Exists)
                            {
                                dirInfo.Create();
                            }
                            

                            foreach (Word.Window window in doc2.Windows)
                            {
                                foreach (Word.Pane pane in window.Panes)
                                {

                                    int j = 0;
                                    foreach (Word.Page p in pane.Pages)
                                    {
                                        var bits = p.EnhMetaFileBits;
                                        try
                                        {
                                            using (var ms = new MemoryStream((byte[])(bits)))
                                            {
                                                printList.Add(System.Drawing.Image.FromStream(ms));
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            System.Console.WriteLine(ex);
                                        }
                                        j++;
                                    }

                                }
                            }
                          /*  MessageBox.Show("1.Изображение записано в буфер");
                            PrintDocument print = new PrintDocument();

                            int g = 0;
                            MessageBox.Show("2.Класс принтера создан");
                            print.PrintPage += (o, e) =>
                            {
                                e.Graphics.DrawImage(printList[g], new System.Drawing.Point(0, 0));
                                g++;
                                if (g == printList.Count)
                                    e.HasMorePages = false;
                                else
                                    e.HasMorePages = true;
                            };
                            MessageBox.Show("3.Изображение записано");
                            PrintDialog prn = new PrintDialog();
                            prn.PrintToFile = true;
                            MessageBox.Show("4.Печатаем в файл");
                            prn.PrinterSettings.PrintFileName = SelectedPath + "\\d.pdf";
                            print.PrinterSettings.PrintFileName = SelectedPath + "\\d.pdf";
                            MessageBox.Show("5.Указали путь");
                            print.PrinterSettings.PrintToFile = true;
                            prn.Document = print;
                            MessageBox.Show("6.Связали файлы");
                            print.Print();
                            MessageBox.Show("7.Напечатали");*/

                            // printList[0].Save("1.jpg");
                            // SaveImageAsPdf("1.jpg", SelectedPath + "\\d.pdf", 600, true);
                            /*      PrintDocument print = new PrintDocument();

                                    int g = 0;
                                    print.PrintPage += (o, e) =>
                                    {
                                        e.Graphics.DrawImage(printList[g], new System.Drawing.Point(0, 0));
                                        g++;
                                        if (g == printList.Count)
                                            e.HasMorePages = false;
                                        else
                                            e.HasMorePages = true;
                                    };

                                 var printrers = PrinterSettings.InstalledPrinters;
                                  print.PrinterSettings.PrinterName = printrers[1];
                                foreach (var s in printrers)
                                {
                                    MessageBox.Show(s.ToString());
                                    if (s.ToString().Contains("PDF"))
                                        print.PrinterSettings.PrinterName = (string)s;
                                }
                                  print.PrinterSettings.PrintFileName = SelectedPath + "\\d.pdf";
                                  print.PrinterSettings.PrintToFile = true;
                                  print.Print();*/



                            /* printList[0].Save("1.jpg");
                             iTextSharp.text.Rectangle pageSize = null;

                             var srcImage = new Bitmap("1.jpg");

                                 pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);

                             using (var ms = new MemoryStream())
                             {
                                 var document = new iTextSharp.text.Document(pageSize, 0, 0, 0, 0);
                                 iTextSharp.text.pdf.PdfWriter.GetInstance(document, ms).SetFullCompression();
                                 document.Open();
                                 var image = iTextSharp.text.Image.GetInstance("1.jpg");
                                 document.Add(image);
                                 document.Close();

                                 File.WriteAllBytes("1.pdf", ms.ToArray());
                             }
                            */

                            doc2.Close(false);
                            doc2 = null;
                            app2.Quit(false);


                            object oMissing = System.Reflection.Missing.Value;
                            Word._Application oWord;
                            Word._Document oDoc;
                            oWord = new Word.Application();
                            oWord.Visible = false;
                            oDoc = oWord.Documents.Open(System.Windows.Forms.Application.StartupPath.ToString() + "\\Data\\3.docx");
                            object f = false;
                            object t = true;
                            object left = 0;
                            object top = 0;
                            object width = 210*3-30;
                            object height = 297*3-30;
                            printList[0].Save("1.jpg");

                            oDoc.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath.ToString() + "\\1.jpg", f,t, left, top, width, height);
                            oDoc.ExportAsFixedFormat(SelectedPath + "\\d.pdf",
                                   WdExportFormat.wdExportFormatPDF);

                            oDoc.Close(false);
                            oDoc = null;
                            oWord.Quit(false);
                            File.Delete("1.jpg");

                            doc.SaveAs2(SelectedPath+"\\лист согласования "+Adres+".doc");


                                    // Закрываем документ
                                    doc.Close(false);
                                    doc = null;
                                    app.Quit(false);


                                    System.Windows.Forms.MessageBox.Show("Документы успешно сгенерированы");
                           if (Directory.Exists(SelectedPath) )
                                System.Diagnostics.Process.Start("explorer.exe", SelectedPath);
                           else
                                System.Diagnostics.Process.Start("explorer.exe", System.Windows.Forms.Application.StartupPath.ToString());


                        }
                        catch (Exception ex)
                        {
                            // Если произошла ошибка, то
                            // закрываем документ и выводим информацию
                            doc.Close(false);
                            doc = null;
                            doc2.Close(false);
                            doc2 = null;
                            System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                            System.Windows.Forms.MessageBox.Show(ex.StackTrace.ToString());
                        }
                    }
                });
            }
        }



        Word.Application app;
        Word.Application app2;
        public Word.Document SaveOne()
        {
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                app = new Word.Application();
            // Открываем
            doc = app.Documents.Open(System.Windows.Forms.Application.StartupPath.ToString() + "\\Data\\1.doc");
            doc.Activate();
            // Добавляем информацию
            // wBookmarks содержит все закладки
            Word.Bookmarks wBookmarks = doc.Bookmarks;

            foreach (Word.Bookmark wRange in wBookmarks)
            {
                var s = wRange.ToString();
                Word.Range range;
                switch (wRange.Name)
                {

                    case "адрес":
                        range = wRange.Range;
                        range.Text = Adres;
                        break;
                    case "район":
                        range = wRange.Range;
                        range.Text = Rayon;
                        break;
                    case "проектировщик":
                        range = wRange.Range;
                        range.Text = Proectirovshic;
                        break;
                    case "заявитель":
                        range = wRange.Range;
                        range.Text = Fio;
                        break;
                    case "вид_объекта":
                        range = wRange.Range;
                        range.Text = SelectVid;
                        break;
                }
            }
            }
            catch (Exception ex)
            {
                // Если произошла ошибка, то
                // закрываем документ и выводим информацию
                doc.Close(false);

                doc = null;
                System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                System.Windows.Forms.MessageBox.Show(ex.StackTrace.ToString());
            }
            return doc;
        }

        public Word.Document SaveTwo()
        {
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                app2 = new Word.Application();
                // Открываем
                doc = app2.Documents.Open(System.Windows.Forms.Application.StartupPath.ToString() + "\\Data\\2.docx");
                doc.Activate();
                // Добавляем информацию
                // wBookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;

                var firstChar = Fio.Trim().ToLower()[0];

                foreach (Word.Bookmark wRange in wBookmarks)
                {
                    var s = wRange.ToString();
                    Word.Range range;
                    switch (wRange.Name)
                    {

                        case "адрес":
                            range = wRange.Range;
                            range.Text = Adres;
                            break;
                        case "фио":
                            range = wRange.Range;
                            range.Text = Fio;
                            break;
                        case "паспорт":
                            range = wRange.Range;
                            range.Text = Pasport;
                            break;
                        case "дата":
                            range = wRange.Range;
                            range.Text = DateTime.Now.ToShortDateString();
                            break;
                        case "подпись":
                            if (File.Exists(System.Windows.Forms.Application.StartupPath.ToString() + "\\Data\\подписи\\" + firstChar + ".png"))
                            {
                                range = wRange.Range;
                                range.InlineShapes.AddPicture(System.Windows.Forms.Application.StartupPath.ToString() + "\\Data\\подписи\\" + firstChar + ".png");
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("Файл с подписью не найден");
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Если произошла ошибка, то
                // закрываем документ и выводим информацию
                doc.Close(false);
                doc = null;
                System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                System.Windows.Forms.MessageBox.Show(ex.StackTrace.ToString());
            }
            return doc;
        }
    }
}
