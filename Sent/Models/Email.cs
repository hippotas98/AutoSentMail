using System;
using System.Reflection.Metadata.Ecma335;

namespace Sent.Models
{
    public class Email
    {
        public string Name { get; set; }
        public string Mail { get; set; }
        public int Status { get; set; }
        public Email(string Mail)
        {
            this.Mail = Mail;
            this.Status = 0;
        }

        public Email()
        {
            
        }
    }
}
