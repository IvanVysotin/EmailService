using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
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
using System.Xml.Linq;
using static OfficeOpenXml.ExcelErrorValue;
using Word = Microsoft.Office.Interop.Word;
//using LingvoNET;

namespace EmailService
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool _dbChosen = false;
        private bool _txtChosen = false;
        private bool _presentationChosen = false;
        private bool _wordChosen = false;
        private string _xlsxPath;
        private string _txtPath;
        private string _presentationPath;
        private string _wordPath;
        private int _totalSend = int.Parse(ConfigurationManager.AppSettings["totalSend"]);

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Метод, которые читает данные с первого листа .xlsx таблицы
        /// </summary>
        private void LoadDataFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //Строка, необходимая для работы EPPlus в некоммерческом режиме

            var clients = new List<Client>();
            using (var package = new ExcelPackage(new FileInfo(_xlsxPath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Данные с первого листа таблицы "Sheet1"

                if (worksheet != null)
                {
                    clientDataGrid.ItemsSource = null;
                    clientDataGrid.Items.Clear();
                    clients.Clear();
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Начало со второй строки, чтобы не захватывать заголовки столбцов
                    {
                        clients.Add(new Client
                        {
                            Company = worksheet.Cells[row, 1].Value?.ToString(),
                            FullName = worksheet.Cells[row, 2].Value?.ToString(),
                            Position = worksheet.Cells[row, 3].Value?.ToString(),
                            Email = worksheet.Cells[row, 4].Value?.ToString(),
                            Phone = worksheet.Cells[row, 5].Value?.ToString()
                        });
                    }
                }
                else MessageBox.Show("Вы выбрали пустую таблицу Excel");
            }
            clientDataGrid.ItemsSource = clients;
        }


        public string WordEdit()
        {
            string tempFilePath = System.IO.Path.GetTempFileName(); // Путь к временному файлу
            File.Copy(_wordPath, tempFilePath, true); // Копирование драфта письма во временный файл
            Word.Application wordApp = new Word.Application();
            try
            {
                foreach (Client client in clientDataGrid.ItemsSource)
                {
                    if (client.FullName != null && client.Email != null)
                    {
                        Word.Document doc = wordApp.Documents.Open(tempFilePath);
                        doc.Content.Find.Execute("...", ReplaceWith: client.FullName.ToString(), Replace: Word.WdReplace.wdReplaceAll);
                        doc.Content.Find.Execute("дд.мм.гггг", ReplaceWith: DateTime.Now.ToString("dd.MM.yyyy"), Replace: Word.WdReplace.wdReplaceAll);
                        doc.Content.Find.Execute(" *", ReplaceWith: _totalSend, Replace: Word.WdReplace.wdReplaceAll);

                        Word.Table table = doc.Tables[2]; // Вторая таблица в документе
                        Word.Cell cell = table.Cell(1, 2); //Вторая ячейка в первой строке
                        cell.Range.Text = client.Position + " " + client.Company + " ";
                        string[] nameParts = client.FullName.Split(' ');
                        if (nameParts.Length >= 2)
                        {
                            // Формирование Фамилии и инициалов
                            string lastName = nameParts[0];
                            string initials = "";
                            for (int i = 1; i < nameParts.Length; i++)
                            {
                                initials += nameParts[i][0] + ".";
                            }

                            // Запись Фамилии и инициалов в ячейку
                            cell.Range.Text += lastName + " " + initials;
                        }
                        else cell.Range.Text += client.FullName;

                        doc.Save();
                        doc.Close();
                    }
                }
            }
            finally
            {
                wordApp.Quit();
            }
            return tempFilePath;
        }

        /// <summary>
        /// Метод для отправки письма
        /// </summary>
        /// <returns></returns>
        private async Task EmailSending() 
        {
            int sendCount = 0;
            string invalidMessage = null;
            string emailSubject = null;
            string emailBody = null;
            Attachment mailAttachment1 = new(_presentationPath);
            var appSettings = ConfigurationManager.AppSettings;

            await Task.Run(async () =>
            {
                emailSubject = File.ReadLines(_txtPath).FirstOrDefault(); // Запись в переменную для темы письма
                foreach (Client client in clientDataGrid.ItemsSource)
                {
                    /*string tempFilePath = System.IO.Path.GetTempFileName(); // Путь к временному файлу
                    File.Copy(_wordPath, tempFilePath, true); // Копирование драфта письма во временный файл
                    Word.Application wordApp = new Word.Application();*/
                    if (client.FullName != null && client.Email != null) 
                    {
                        
                        // Запись в переменные для тела письма
                        emailBody = string.Join(Environment.NewLine, 
                            File.ReadLines(_txtPath).Skip(2)).Replace("...", client.FullName);

                        if (emailSubject == null || emailBody == null) throw new IOException("Ошибка с файлом содержимого письма");
                        try
                        {
                            // Объект SmtpClient для отправки почты
                            using (SmtpClient smtpClient = new SmtpClient("smtp.mail.ru", 2525))
                            {
                                /*Word.Document doc = wordApp.Documents.Open(tempFilePath);
                                doc.Content.Find.Execute("...", ReplaceWith: client.FullName.ToString(), Replace: Word.WdReplace.wdReplaceAll);
                                doc.Content.Find.Execute("дд.мм.гггг", ReplaceWith: DateTime.Now.ToString("dd.MM.yyyy"), Replace: Word.WdReplace.wdReplaceAll);
                                doc.Content.Find.Execute(" *", ReplaceWith: _totalSend, Replace: Word.WdReplace.wdReplaceAll);

                                doc.Save();
                                doc.Close();
                                wordApp.Quit();*/
                                Attachment mailAttachment2 = new(WordEdit());
                                mailAttachment2.Name = "Официальное письмо.docx";

                                smtpClient.EnableSsl = true; // Включение SSL-протокола
                                smtpClient.Credentials = new NetworkCredential("mr.zombik123@mail.ru", "hwix8KXYbewdm2Txvmhw"); // Данные для почты отправителя

                                // Адреса От, Кому и Ответить
                                MailAddress from = new("mr.zombik123@mail.ru");
                                MailAddress to = new(client.Email, client.FullName);
                                MailAddress replyTo = new("mr.zombik123@mail.ru");

                                MailMessage mailMessage = new(from, to); // Письмо (От, Кому)
                                mailMessage.ReplyToList.Add(replyTo); // Ответить

                                mailMessage.Attachments.Add(mailAttachment1);
                                mailMessage.Attachments.Add(mailAttachment2);

                                if (emailSubject != null || emailBody != null)
                                {
                                    mailMessage.Subject = emailSubject; // Тема письма
                                    mailMessage.SubjectEncoding = Encoding.UTF8;

                                    mailMessage.Body = emailBody; // Содержимое письма
                                    mailMessage.BodyEncoding = Encoding.UTF8;
                                    mailMessage.IsBodyHtml = false;
                                }
                                else return;
                                await smtpClient.SendMailAsync(mailMessage);
                            };
                            _totalSend++;
                            sendCount++;
                        }
                        catch (IOException ex)
                        {
                            invalidMessage += ex.Message + "\n";
                        }
                        catch (SmtpException ex)
                        {
                            invalidMessage += ex.Message + "\n";
                        }
                    }
                }
                
                mailAttachment1.Dispose();
            });
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["totalSend"].Value = _totalSend.ToString();
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            if (invalidMessage != null) MessageBox.Show(invalidMessage);
            MessageBox.Show($"Было отправленно писем – {sendCount}\n");
        }

        /// <summary>
        /// Метод обработки нажатия на кнопку.
        /// Вызывает метод отправки письма
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="ApplicationException"></exception>
        private async void SendButtonClick(object sender, RoutedEventArgs e)
        {
            await EmailSending();
        }

        private void SelectDBClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("ВНИМАНИЕ. Учтите, что файл таблицы должен иметь следующую струтуру:\n" +
                "Первая строка - заголовок столбцов;\n" +
                "Содержание столбцов:\n\n" +
                "Наименование организации | ФИО | Должность | e-mail | Телефон\n");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Таблица Excel|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                _xlsxPath = openFileDialog.FileName;
                _dbChosen = true;
                LoadDataFromExcel();
            }

            IsAllSelected();
        }

        private void SelectTxtClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("ВНИМАНИЕ. Учтите, что текстовый файл должен иметь следующую струтуру:\n" +
                "Первая строка - заголовок письма;\n" +
                "Вторая строка - пропуск;\n" +
                "Третья строка и последующие - содержимое письма.\n\n" +
                "В ином случае целостность письма будет нарушена");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Текстовый файл|*.txt";
            if (openFileDialog.ShowDialog() == true)
            {
                _txtPath = openFileDialog.FileName;
                _txtChosen = true;
            }
            
            IsAllSelected();
        }

        private void SelectPresentaionClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Файл PDF|*.pdf";
            if (openFileDialog.ShowDialog() == true)
            {
                _presentationPath = openFileDialog.FileName;
                _presentationChosen = true;
            }

            IsAllSelected();
        }

        private void SelectLetterClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("ВНИМАНИЕ. В вашем письме должны присутствовать определенные символы для их замены\n\n" +
                "Символ '*' для номера исходящего письма.\n" +
                "Символ '...' для обращения к адресату.");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Документ Word|*.doc;*.docx;*.dot;*.dotx";
            if (openFileDialog.ShowDialog() == true)
            {
                _wordPath = openFileDialog.FileName;
                _wordChosen = true;
            }

            IsAllSelected();
        }

        private void IsAllSelected() 
        {
            if (_dbChosen && _txtChosen && _presentationChosen && _wordChosen) SendButton.IsEnabled = true;
            else SendButton.IsEnabled = false;
        }
    }
}