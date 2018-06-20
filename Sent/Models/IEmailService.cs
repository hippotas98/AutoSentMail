using System;
using System.Collections.Generic;
using Sent.Models;
namespace Sent.Models
{
    public interface IEmailService
    {
        List<Email> SendEmail(List<string> emails, User user);
       // bool Save(User user);
        List<Email> GetAllEmails();
        List<Email> SendEmailByExcel(User user, string filePath, string sheetName, int NumOfCol);
    }
}