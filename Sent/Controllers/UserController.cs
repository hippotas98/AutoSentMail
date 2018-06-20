using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Sent.Models;

namespace Sent.Controllers
{
    [Route("api/Users")]
    public class UserController : Controller
    {
        private IEmailService iemailService;

        public UserController(IEmailService iemailService)
        {
            this.iemailService = iemailService;
        }
        [HttpGet("Email")]
        public List<Email> getAllEmails()
        {
            return this.iemailService.GetAllEmails();
        }

        [HttpPost("Email")]
        public List<Email> SendMails([FromBody]User user)
        {
            List<Email> emails = this.iemailService.SendEmail(user.UnsentMail, user);
            
            return emails;
        }

        [HttpPost("EmailExcel")]
        public List<Email> SendEmailFromExcels([FromBody] User user)
        {
            List<Email> emails =
                this.iemailService.SendEmailByExcel(user, "/Users/apple/Projects/Sent/SampleDatabase.xlsx", "Sheet1",2);
            return emails;
        }
    }
}