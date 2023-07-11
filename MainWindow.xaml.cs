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
using MailKit.Net.Smtp;
using MailKit.Security;
using System.Security;

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
        private bool _credentialCheckResult = false;
        private string _xlsxPath;
        private string _txtPath;
        private string _presentationPath;
        private string _wordPath;
        private int _totalSend = int.Parse(ConfigurationManager.AppSettings["totalSend"]);
        private string _smtpServer = null;
        private int _smtpPort = 465;

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

        /// <summary>
        /// Проверка данных для авторизации
        /// </summary>
        /// <returns></returns>
        private bool CheckCredentials()
        {
            if (_txtEmailAddress.Text.Length <= 7 || !_txtEmailAddress.Text.Contains('@') || _txtEmailPassword.SecurePassword.ToString() == null)
            {
                return false;
            }
            else 
            {
                switch (_txtEmailAddress.Text.Split('@')[1])
                {
                    case "mail.ru":
                        _smtpServer = "smtp.mail.ru";
                        _smtpPort = 2525;
                        break;
                    case "outlook.com":
                        _smtpServer = "smtp.office365.com";
                        _smtpPort = 587;
                        break;
                    case "gmail.com":
                        _smtpServer = "smtp.gmail.com";
                        _smtpPort = 465;
                        break;
                    case "yandex.ru":
                        _smtpServer = "smtp.yandex.ru";
                        _smtpPort = 465;
                        break;
                    case "yahoo.com":
                        _smtpServer = "smtp.mail.yahoo.com";
                        _smtpPort = 587;
                        break;
                    case "aol.com":
                        _smtpServer = "smtp.aol.com";
                        _smtpPort = 587;
                        break;
                }
            }

            if (_smtpServer != null)
            {
                using (var client = new MailKit.Net.Smtp.SmtpClient())
                {
                    client.Connect(_smtpServer, _smtpPort, SecureSocketOptions.StartTls);
                    client.Authenticate(_txtEmailAddress.Text, new System.Net.NetworkCredential(string.Empty, _txtEmailPassword.SecurePassword).Password);

                    // Если аутентификация прошла успешно, значит логин и пароль корректны
                    bool credentialsValid = client.IsAuthenticated;

                    client.Disconnect(true);
                    return credentialsValid;
                }
            }
            else return false;
        }

        /// <summary>
        /// Метод, который редактирует данные в документе Word согласно данным из excel
        /// </summary>
        /// <returns></returns>
        public async Task<string> WordEditAsync(Client client)
        {
            string tempFilePath = System.IO.Path.GetTempFileName(); // Путь к временному файлу
            File.Copy(_wordPath, tempFilePath, true); // Копирование драфта письма во временный файл
            Word.Application wordApp = new Word.Application();
            try
            {
                Word.Document doc = wordApp.Documents.Open(tempFilePath);
                // Замены символов
                await Task.Run(() => 
                {
                    doc.Content.Find.Execute("...", ReplaceWith: client.FullName.ToString(), Replace: Word.WdReplace.wdReplaceAll);
                    doc.Content.Find.Execute("дд.мм.гггг", ReplaceWith: DateTime.Now.ToString("dd.MM.yyyy"), Replace: Word.WdReplace.wdReplaceAll);
                    doc.Content.Find.Execute(" *", ReplaceWith: _totalSend, Replace: Word.WdReplace.wdReplaceAll);
                });

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

                await Task.Run(() => doc.Save());
                doc.Close();
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
        private async Task EmailSendingAsync() 
        {
            int sendCount = 0;
            string invalidMessage = null;
            string emailSubject = null;
            string emailBody = null;
            Attachment mailAttachment1 = new(_presentationPath);
            var appSettings = ConfigurationManager.AppSettings;

            emailSubject = File.ReadLines(_txtPath).FirstOrDefault(); // Запись в переменную для темы письма
            foreach (Client client in clientDataGrid.ItemsSource)
            {
                if (client.FullName != null && client.Email != null)
                {
                    // Запись в переменную для тела письма
                    emailBody = string.Join(Environment.NewLine,
                        File.ReadLines(_txtPath).Skip(2)).Replace("...", client.FullName);

                    if (emailSubject == null || emailBody == null) throw new IOException("Ошибка с файлом содержимого письма");
                    try
                    {
                        // Объект SmtpClient для отправки почты
                        using (System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(_smtpServer, _smtpPort))
                        {
                            string attachmentPath = await WordEditAsync(client);
                            Attachment mailAttachment2 = new(attachmentPath);
                            mailAttachment2.Name = "Официальное письмо.docx";

                            smtpClient.EnableSsl = true; // Включение SSL-протокола
                            MailAddress from = null;
                            MailAddress to = null;
                            MailAddress replyTo = null;
                            Dispatcher.Invoke(() =>
                            {
                                smtpClient.Credentials = new NetworkCredential(_txtEmailAddress.Text, _txtEmailPassword.SecurePassword); // Данные для почты отправителя
                                // Адреса От, Кому и Ответить
                                from = new(_txtEmailAddress.Text);
                                to = new(client.Email, client.FullName);
                                replyTo = new(_txtEmailAddress.Text);
                            }); 
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
            
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["totalSend"].Value = _totalSend.ToString();
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            if (invalidMessage != null) MessageBox.Show(invalidMessage);
            MessageBox.Show($"Было отправленно писем – {sendCount}\n");
        }

        /// <summary>
        /// Метод для отображения кнопки отправить, если все файлы были выбраны
        /// </summary>
        private void IsAllSelected()
        {
            if (_dbChosen && _txtChosen && _presentationChosen && _wordChosen) SendButton.IsEnabled = true;
            else SendButton.IsEnabled = false;
        }

        /// <summary>
        /// Метод обработки нажатия на кнопку.
        /// Вызывает метод отправки письма и редактирования Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="ApplicationException"></exception>
        private async void SendButtonClickAsync(object sender, RoutedEventArgs e)
        {
            //bool _credentialCheckResult = CheckCredentials().Result;
            if (CheckCredentials()) await Task.Run(() => EmailSendingAsync());
            else MessageBox.Show("Ошибка с данными электронной почты");
        }

        /// <summary>
        /// Метод обработки нажатия кнопки выбора БД
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Метод обработки нажатия кнопки выбора тела письма
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Метод обработки нажатия кнопки выбора презентации
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Метод обработки нажатия кнопки выбора официального файла письма
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
    }
}