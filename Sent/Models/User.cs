using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;

namespace Sent.Models
{
	public class User
	{
        public List<Email> emails {get; set;}
		public List<string> UnsentMail { get; set; }
        public string Username {get;set;}
        public string Password {get;set;}
		public string Header { get; set; }
		public string Content { get; set; }
		public User(string UserName, string Password, string Header, string Content, List<string> unsentMail)
		{
			this.Username = UserName;
			this.Password = Password;
			this.Header = Header;
			this.Content = Content;
			this.UnsentMail = unsentMail;
			this.emails = new List<Email>();
		}

		public User()
		{
			this.emails = new List<Email>();
		}
        
	}
}
