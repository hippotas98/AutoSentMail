using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace Sent.Models
{
    public class EmailService : IEmailService
    {
        public List<Email> SendEmail(List<string> mails, User user)
        {
            
            using (SmtpClient gmail = new SmtpClient("smtp.gmail.com", 587))
            {
                gmail.UseDefaultCredentials = false;
                gmail.EnableSsl = true;
                gmail.Credentials = new NetworkCredential(user.Username, user.Password);
                
                foreach (var email in mails)
                {
                    try
                    {
                        
                        user.emails.Add(new Email(email));
                        MailMessage message = new MailMessage();
                        message.From = new MailAddress(user.Username);
                        message.To.Add(email);
                        message.Subject = user.Header;
                        message.Body = user.Content;
                        gmail.Send(message);
                        user.emails.Where(s => s.Mail == email).FirstOrDefault().Status = 1;
                        message = null;
                    }
                    catch (Exception e)
                    {
                        user.emails.Where(s => s.Mail == email).FirstOrDefault().Status = -1;
                        Console.WriteLine(email + "cannot be sent" + e);
                    }
                }

                Save(user);
            }
            
            return user.emails;
        }

        public List<Email> GetAllEmails()
        {
            using (StreamReader reader = new StreamReader("./Data.txt"))
            {
                var result = JsonConvert.DeserializeObject<List<Email>>(reader.ReadToEnd());
                return result;
            }
        }
        private bool Save(User user)
        {
            using (StreamWriter file = new StreamWriter("./Data.txt"))
            {
                JsonSerializer serializer= new JsonSerializer();
                try
                {
                    serializer.Serialize(file, user.emails);
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    return false;
                }
                
            }
        }

        public List<Email> SendEmailByExcel(User user, string filePath, string sheetName, int NumOfCol)
        {
            DataTable dt = ReadFileExcel(filePath, sheetName, NumOfCol);
            List<Email> UnsentEmails = GetListMailFromExcel(dt);
            //return UnsentEmails;
            using (SmtpClient gmail = new SmtpClient("smtp.gmail.com", 587))
            {
                gmail.UseDefaultCredentials = false;
                gmail.EnableSsl = true;
                gmail.Credentials = new NetworkCredential(user.Username, user.Password);
                foreach (var email in UnsentEmails)
                {
                    try
                    {
                        user.emails.Add(email);
                        MailMessage message = new MailMessage();
                        message.From = new MailAddress(user.Username);
                        message.To.Add(email.Mail);
                        message.Subject = user.Header;
                        message.Body = "Xin chao " + email.Name + "\n" + user.Content;
                        gmail.Send(message);
                        user.emails.Where(s => s.Mail == email.Mail).FirstOrDefault().Status = 1;
                        message = null;
                    }
                    catch (Exception e)
                    {
                        user.emails.Where(s => s.Mail == email.Mail).FirstOrDefault().Status = -1;
                        Console.WriteLine(email + "cannot be sent" + e);
                    }
                }

                Save(user);
                WriteFileExcel("/Users/apple/Projects/Sent/Result.xlsx", sheetName, NumOfCol, user);
                return user.emails;
            }
        }

        private List<Email> GetListMailFromExcel(DataTable dataTable)
        {
            List<Email> emails = new List<Email>();
            for(int numRow = 1;numRow<dataTable.Rows.Count;numRow++)
            {
                DataRow row = dataTable.Rows[numRow];
                Email mail= new Email();
                mail.Mail = new MailAddress(row[1].ToString()).Address;
                mail.Name = row[0].ToString();
                if(mail.Mail==row[1].ToString())
                    emails.Add(mail);
            }

            return emails;
        }
        private DataTable ReadFileExcel(string filePath, string sheetName, int NumOfCol)
        {
            
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                DataTable dt = new DataTable();
                if (package.Workbook.Worksheets.Count != 0)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Where(s => s.Name == sheetName).FirstOrDefault();
                    if (worksheet == null) return null;
                    for(var index = 1;index <= NumOfCol;++index)
                    {
                        dt.Columns.Add();
                    }
                    for (var rowNumber = 1; rowNumber <= worksheet.Dimension.End.Row; rowNumber++ )
                    {
                        
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, NumOfCol];
                        var newRow = dt.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column-1] = cell.Text;
                        }

                        dt.Rows.Add(newRow);
                    }
                    return dt;
                    
                }

                return null;
            }
        }

        private bool WriteFileExcel(string filePath, string sheetName, int NumOfCol, User user)
        {
           
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count != 0)
                {
                    ExcelWorksheet worksheet =
                        package.Workbook.Worksheets.Where(s => s.Name == sheetName).FirstOrDefault();
                    //IEnumerable<int> status = user.emails.Select(s => s.Status);
                    var temp = user.emails.Select(s => new
                    {
                        Name = s.Name,
                        Mail = s.Mail,
                        Status = s.Status == 1 ? "Thanh cong" : (s.Status==0 ? "Chua gui" : "That Bai")
                    });
                    worksheet.Cells[2,1].Clear();
                    worksheet.Cells[2, 1].LoadFromCollection(temp,false);
                    
                    package.Save();
                    return true;

                }
                
            }

            return false;
        }
    }
}