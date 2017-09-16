using System;
using Microsoft.Exchange.WebServices.Data;

namespace LyncLogger
{
	class Conversation365
	{
		public EmailMessage Message{ get; set; }
		public DateTime LastSaved { get; set; }
		public int UnsavedMessageCount = 0;
	}
}
