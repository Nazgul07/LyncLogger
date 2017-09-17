using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Conversation.AudioVideo;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using Notifications.Wpf;
using Conversation = Microsoft.Lync.Model.Conversation.Conversation;

namespace LyncLogger
{
	internal class LyncLogger
	{
		public static bool TextLoggingEnabled {
			get => SettingsManager.ReadSetting("TextLogging").ToUpper() == "ON";
			set => SettingsManager.AddUpdateAppSettings("TextLogging", value ? "On" : "Off");
		}
		private static LyncClient _client = null;
		private static readonly Dictionary<Conversation, Conversation365> OutlookConversations = new Dictionary<Conversation, Conversation365>();
		private const string LogHeader = "// Convestation started with {0} on {1}"; //header of the file
		private const string LogMiddleHeader = "---- conversation resumed ----"; //middle header of the file
		private const string LogMessage = "{0} ({1}): {2}"; //msg formating
		private const string LogMessageHtml = "<span style=\"font-size:11px;font-variant:normal;text-transform:none;\"><b>{0}&nbsp;{1}</b></span>:<br/>&nbsp&nbsp{2}"; //msg formating
		private const string LogAudio = "Audio conversation {0} at {1}"; //msg audio started/ended formating

		private const int DelayRetryAuthentication = 20000; // delay before authentication retry (in ms)
		private const string ExceptionLyncNoclient = "The host process is not running";

		private static DirectoryInfo _folderLog;
		private static string _fileLog;
		private static string _nameShortener;

		private static readonly System.Timers.Timer Timer365Save = new System.Timers.Timer(60000);

		public static void Run(string folderLog)
		{
			_folderLog = new DirectoryInfo(folderLog);
			_fileLog = Path.Combine(folderLog, "{0}_{1}.log");
			_nameShortener = SettingsManager.ReadSetting("shortenName");
			Timer365Save.Elapsed += (sender, args) =>
			{
				foreach (var keyPair in OutlookConversations)
				{
					if (keyPair.Value.LastSaved.AddMinutes(5) <= DateTime.Now)
					{
						Save365Conversation(keyPair.Key);
					}
				}
			};
			Timer365Save.Start();
			Run();
		}

		/// <summary>
		/// Constructor, Listen on new opened conversations
		/// </summary>
		private static void Run()
		{
			try
			{
				//Start the conversation
				if (_client == null)
				{
					_client = LyncClient.GetClient();
				}
				if (_client.State == ClientState.SignedIn)
				{
					//handles the states of the logger displayed in the systray
					_client.StateChanged += ClientOnStateChanged;

					//listen on conversation in order to log messages
					ConversationManager conversations = _client.ConversationManager;

					//check our listener is not already registered
					var handler = typeof(ConversationManager).GetField("ConversationAdded", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(conversations) as Delegate;

					if (handler == null)
					{
						conversations.ConversationAdded += Conversations_ConversationAdded;
						NotifyIconSystray.ChangeStatus(true);
					}
					else
					{
						NotifyIconSystray.ChangeStatus(true);
					}
				}
				else
				{
					Thread.Sleep(DelayRetryAuthentication / 10);
					Run();
				}

				SaveAll365Conversations();
				OutlookConversations.Clear();
			}
			catch (ClientNotFoundException)
			{
				Thread.Sleep(DelayRetryAuthentication);
				Run();
			}
			catch (SystemException systemException)
			{
				if (!IsLyncException(systemException))
				{
					// Rethrow the SystemException which did not come from the Lync Model API.
					throw;
				}
			}
		}

		private static void ClientOnStateChanged(object s, ClientStateChangedEventArgs e)
		{
			if (e.NewState == ClientState.SignedOut)
			{
				NotifyIconSystray.ChangeStatus(false);
				Client client = ((Client) s);
				client.StateChanged -= ClientOnStateChanged;
				ConversationManager conversations = client.ConversationManager;
				conversations.ConversationAdded -= Conversations_ConversationAdded;
				Run();
			}
		}

		/// <summary>
		/// Create conversation log file and listen on what participants say
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private static void Conversations_ConversationAdded(object sender, ConversationManagerEventArgs e)
		{
			String firstContactName = e.Conversation.Participants.Count > 1
			? e.Conversation.Participants[1].Contact.GetContactInformation(ContactInformationType.DisplayName).ToString()
			: "meet now";
			DateTime currentTime = DateTime.Now;

			String fileLog = String.Format(_fileLog, firstContactName.Replace(", ", "_"), currentTime.ToString("yyyyMMdd"));

			String logHeader;
			FileInfo[] files = _folderLog.GetFiles("*.log");
			if (files.Count(f => f.Name == fileLog.Substring(fileLog.LastIndexOf('\\') + 1)) == 0)
			{
				logHeader = String.Format(LogHeader, firstContactName, currentTime.ToString("yyyy/MM/dd"));
			}
			else
			{
				logHeader = String.Format(LogMiddleHeader);
			}

			if (TextLoggingEnabled)
			{
				using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
				{
					using (StreamWriter writer = new StreamWriter(stream))
					{
						writer.WriteLine(logHeader);
					}
				}
			}

			Conversation conv = e.Conversation;
			Start365Conversation(conv);

			//detect all participant (including user)
			conv.ParticipantAdded += (_sender, _e) =>
			{
				var participant = _e.Participant;
				InstantMessageModality remoteImModality = (InstantMessageModality)participant.Modalities[ModalityTypes.InstantMessage];

				//detect all messages (including user's)
				remoteImModality.InstantMessageReceived += (__sender, __e) =>
				{
					RemoteImModality_InstantMessageReceived(__sender, __e, fileLog);
				};
			};

			//get audio conversation informations about user (not the other participants)
			AVModality callImModality = (AVModality)conv.Participants[0].Modalities[ModalityTypes.AudioVideo];
			//notify call 
			callImModality.ModalityStateChanged += (_sender, _e) =>
			{
				CallImModality_ModalityStateChanged(_e, fileLog + ".mp3");
			};
		}
		

		/// <summary>
		/// log to fileLog the date of the start and end of a call
		/// (ModalityStateChanged callback)
		/// </summary>
		/// <param name="e"></param>
		/// <param name="fileLog"></param>
		private static void CallImModality_ModalityStateChanged(ModalityStateChangedEventArgs e, String fileLog)
		{
			if (TextLoggingEnabled)
			{
				//write log only on connection or disconnection
				if (e.NewState == ModalityState.Connected || e.NewState == ModalityState.Disconnected)
				{
					//write start/end info to log
					using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
					{
						using (StreamWriter writer = new StreamWriter(stream))
						{
							writer.WriteLine(LogAudio, (e.NewState == ModalityState.Connected) ? "started" : "ended",
								DateTime.Now.ToString("HH:mm:ss"));
						}
					}
				}
			}
			//record conversation
			if (e.NewState == ModalityState.Connected)
			{
				AudioLogger.Instance.Start(fileLog);
			}

			//end recording
			if (e.NewState == ModalityState.Disconnected)
			{
				AudioLogger.Instance.Stop();
			}
		}

		/// <summary>
		/// log to fileLog all messages of a conversation
		/// (InstantMessageReceived callback)
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <param name="fileLog"></param>
		private static void RemoteImModality_InstantMessageReceived(object sender, MessageSentEventArgs e, String fileLog)
		{
			InstantMessageModality modality = (InstantMessageModality)sender;
			//gets the participant name
			string name = (string)modality.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName);


			//reads the message in its plain text format (automatically converted)
			string message = e.Text.Trim();

			if (TextLoggingEnabled)
			{
				//write message to log
				using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
				{
					using (StreamWriter writer = new StreamWriter(stream))
					{
						if (name.Contains(_nameShortener))
						{
							name = name.Substring(name.IndexOf(_nameShortener) + _nameShortener.Length);
						}
						writer.WriteLine(LogMessage, name, $"{DateTime.Now:hh:mm:ss tt}", message);
					}
				}
			}

			AppendTo365Conversation(modality.Conversation, string.Format(LogMessageHtml, name, $"{DateTime.Now:h:mm tt}", e.Contents[InstantMessageContentType.Html]));
		}

		private static void Start365Conversation(Conversation converation)
		{
			//Connect to exchange
			var ewsProxy = new ExchangeService() { Url = new Uri("https://outlook.office365.com/ews/exchange.asmx") };
			
			//Create the conversation message
			var message = new EmailMessage(ewsProxy);

			string username = SettingsManager.ReadSetting("office365username");
			string password = SettingsManager.ReadSetting("office365password");
			var cred = new NetworkCredential(username, SecureCredentials.DecryptString(password));

			ewsProxy.Credentials = cred;
			
			if (Program.Validate365Credentials(cred))
			{
				message.Body = string.Empty;
				
				message.Sender = new EmailAddress(((converation.Properties[ConversationProperty.Inviter] as Microsoft.Lync.Model.Contact).GetContactInformation(ContactInformationType.EmailAddresses) as List<object>).First() as string);
			
				ExtendedPropertyDefinition extendedPropertyDefinition =
					new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
				message.SetExtendedProperty(extendedPropertyDefinition, 1); // sets the message as a non-draft

				OutlookConversations[converation] = new Conversation365
				{
					Message = message,
					LastSaved = DateTime.Now,
					UnsavedMessageCount = 0
				};
			}
			else if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
			{
				Notifications.Send("Error connecting to Office 365", NotificationType.Error);
			}
		}

		private static void AppendTo365Conversation(Conversation converation, string text)
		{
			if (OutlookConversations.ContainsKey(converation))
			{
				Conversation365 conversation365 = OutlookConversations[converation];
				EmailMessage message = conversation365.Message;
				message.Body = new MessageBody(message.Body.Text + text + "<br/>");
				StringBuilder subject = new StringBuilder();
				subject.Append("Conversation with ");

				foreach (Participant participant in converation.Participants)
				{
					if (!participant.IsSelf)
					{
						subject.Append(participant.Contact.GetContactInformation(ContactInformationType.DisplayName));
						if (converation.Participants.IndexOf(participant) < converation.Participants.Count - 1)
						{
							subject.Append(", ");
						}
					}
					string email = (participant.Contact.GetContactInformation(ContactInformationType.EmailAddresses) as List<object>).First() as string;
					if (message.ToRecipients.All(x => x.Address != email))
					{
						message.ToRecipients.Add(email);
					}
				}
				message.Subject = subject.ToString();

				if (conversation365.UnsavedMessageCount++ == 25)
				{
					Save365Conversation(converation);
				}
			}
		}

		private static void Save365Conversation(Conversation converation)
		{
			if (OutlookConversations.ContainsKey(converation))
			{
				Conversation365 conversation365 = OutlookConversations[converation];
				EmailMessage message = conversation365.Message;
				if (message.Id == null)
				{
					message.Save(WellKnownFolderName.ConversationHistory);
				}
				else
				{
					message.Update(ConflictResolutionMode.AutoResolve);
				}
				conversation365.UnsavedMessageCount = 0;
				conversation365.LastSaved = DateTime.Now;
				Notifications.Send("Conversation synced to Office 365", NotificationType.Information);
			}
		}

		private static void SaveAll365Conversations()
		{
			foreach (var keyPair in OutlookConversations)
			{
				Save365Conversation(keyPair.Key);
			}
		}

		/// <summary>
		/// Identify if a particular SystemException is one of the exceptions which may be thrown
		/// by the Lync Model API.
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		private static bool IsLyncException(SystemException ex)
		{
			return
				ex is NotImplementedException ||
				ex is ArgumentException ||
				ex is NullReferenceException ||
				ex is NotSupportedException ||
				ex is IndexOutOfRangeException ||
				ex is InvalidOperationException ||
				ex is TypeLoadException ||
				ex is TypeInitializationException ||
				ex is InvalidComObjectException ||
				ex is InvalidCastException;
		}
	}
}
