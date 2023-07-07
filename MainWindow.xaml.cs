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
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadDataFromExcel();
        }

        private void LoadDataFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var clients = new List<Client>();
            //string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Client_base.xlsx");
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\mrzom\\source\\repos\\EmailService\\Assets\\Client_base.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"]; // Assuming data is on the first sheet

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Start from row 2 to skip header
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

        private async void SendButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // Объект SmtpClient для отправки почты
                SmtpClient smtpClient = new SmtpClient("smtp.mail.ru", 2525);
                smtpClient.EnableSsl = true;

                smtpClient.Credentials = new NetworkCredential("mr.zombik123@mail.ru", "2wayTqhJQChZ76aKtfng");

                // Адреса От, Кому и Ответить
                MailAddress from = new("mr.zombik123@mail.ru");
                MailAddress to = new("mr.zombik123@mail.ru");
                MailAddress replyTo = new("mr.zombik123@mail.ru");

                MailMessage mailMessage = new(from, to); // От, Кому
                mailMessage.ReplyToList.Add(replyTo); // Ответить

                // Тема письма и содержимое
                mailMessage.Subject = "Test Subject";
                mailMessage.SubjectEncoding = Encoding.UTF8;

                mailMessage.Body = "Test Body";
                mailMessage.BodyEncoding = Encoding.UTF8;

                mailMessage.IsBodyHtml = false;

                await smtpClient.SendMailAsync(mailMessage);
            }
            catch (SmtpException ex)
            {
                throw new ApplicationException("SmtpError occured" + ex.Message);
            }
            // Переменные, необходимые для отправки письма
            /*string senderEmail = "ivntz.apptest.main@mail.ru"; // Почта отправителя
            string recipientEmail = "ivntz.apptest.ivan@mail.ru"; // Почта получателя
            string subject = "test subject"; // Тема письма
            string content = "test content"; // Текст письма*/
            /*string recipientEmail = recipientTextBox.Text; // Почта получателя
            string subject = subjectTextBox.Text; // Тема письма
            string content = contentTextBox.Text; // Текст письма*/



            // Объект MailMessage с информацией о письме
            //MailMessage mailMessage = new MailMessage(senderEmail, recipientEmail, subject, content);
            /*try
            {
                // Отправка письма
                smtpClient.Send(mailMessage);
                Console.WriteLine("Письмо отправлено успешно!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при отправке письма: " + ex.Message);
            }*/
        }
    }
}
