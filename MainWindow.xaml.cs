using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection.Metadata;
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
using System.Xml.Linq;

namespace EmailService
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadDataFromExcel();
        }

        /// <summary>
        /// Метод, которые читает данные с первого листа .xlsx таблицы
        /// </summary>
        private void LoadDataFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //Строка, необходимая для работы EPPlus в некоммерческом режиме
            var clients = new List<Client>();
            using (var package = new ExcelPackage(new FileInfo("Assets\\Client_base.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"]; // Данные с первого листа таблицы "Sheet1"

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Начало со второй строки, чтобы не захватывать заголовки столбцов
                {
                    clients.Add(new Client
                    {
                        FullName = worksheet.Cells[row, 1].Value?.ToString(),
                        Position = worksheet.Cells[row, 2].Value?.ToString(),
                        Email = worksheet.Cells[row, 3].Value?.ToString(),
                        Phone = worksheet.Cells[row, 4].Value?.ToString()
                    });
                }
            }
            clientDataGrid.ItemsSource = clients;
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
            Attachment mailAttachment = new("Assets\\EYECONT.pdf");
            await Task.Run(async () =>
            {
                emailSubject = File.ReadLines("Assets\\Text.txt").FirstOrDefault(); // Запись в переменную для темы письма
                foreach (Client client in clientDataGrid.ItemsSource)
                {
                    // Запись в перенную для тела письма
                    emailBody = string.Join(
                        Environment.NewLine, 
                        File.ReadLines("Assets\\Text.txt").Skip(2)).Replace("...", client.FullName);

                    if (emailSubject == null || emailBody == null) throw new IOException("Ошибка с файлом для отправки");
                    try
                    {
                        // Объект SmtpClient для отправки почты
                        using (SmtpClient smtpClient = new SmtpClient("smtp.mail.ru", 2525))
                        {
                            smtpClient.EnableSsl = true; // Включение SSL-протокола
                            smtpClient.Credentials = new NetworkCredential("your_email_here", "your_password_here"); // Данные для почты отправителя

                            // Адреса От, Кому и Ответить
                            MailAddress from = new("your_email_here");
                            MailAddress to = new(client.Email, client.FullName);
                            MailAddress replyTo = new("your_email_here");

                            MailMessage mailMessage = new(from, to); // Письмо (От, Кому)
                            mailMessage.ReplyToList.Add(replyTo); // Ответить

                            mailMessage.Attachments.Add(mailAttachment);

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
            });
            if (invalidMessage != null) MessageBox.Show(invalidMessage);
            MessageBox.Show($"Количество отправленных писем – {sendCount}");
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
    }
}