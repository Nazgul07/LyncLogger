using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
