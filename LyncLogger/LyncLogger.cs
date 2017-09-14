using System;
using System.Linq;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Conversation.AudioVideo;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using log4net;

namespace LyncLogger
{
	internal class LyncLogger
	{
		private const string LogHeader = "// Convestation started with {0} on {1}"; //header of the file
		private const string LogMiddleHeader = "---- conversation resumed ----"; //middle header of the file
		private const string LogMessage = "{0} ({1}): {2}"; //msg formating
		private const string LogAudio = "Audio conversation {0} at {1}"; //msg audio started/ended formating

		private const int DelayRetryAuthentication = 20000; // delay before authentication retry (in ms)
		private const string ExceptionLyncNoclient = "The host process is not running";

		private static DirectoryInfo _folderLog;
		private static string _fileLog;
		private static string _nameShortener;

		private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		public LyncLogger(string folderLog)
		{
			_folderLog = new DirectoryInfo(folderLog);
			_fileLog = Path.Combine(folderLog, "conversation_{0}_{1}.log");
			_nameShortener = SettingsManager.ReadSetting("shortenName");

			Run();
		}

		/// <summary>
		/// Constructor, Listen on new openned conversations
		/// </summary>
		public void Run()
		{
			try
			{
				//Start the conversation
				LyncClient client = LyncClient.GetClient();

				//handles the states of the logger displayed in the systray
				client.StateChanged += (s, e) =>
				{
					if (e.NewState == ClientState.SignedOut)
					{
						Log.Info("User signed out. Watch for signed in event");
						NotifyIconSystray.ChangeStatus(false);
						Run();
					}
				};

				if (client.State == ClientState.SignedIn)
				{
					//listen on conversation in order to log messages
					ConversationManager conversations = client.ConversationManager;

					//check our listener is not already registered
					var handler = typeof(ConversationManager).GetField("ConversationAdded", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(conversations) as Delegate;

					if (handler == null)
					{
						Log.Info("watch conversation");
						conversations.ConversationAdded += conversations_ConversationAdded;
						NotifyIconSystray.ChangeStatus(true);
					}
					else
					{
						Log.Info("Conversation already in watching state");
						Log.Info(handler);
					}

				}
				else
				{
					Log.Info(
						$"Not signed in. Watch for signed in event. Retry in {DelayRetryAuthentication / 10} ms");
					Thread.Sleep(DelayRetryAuthentication / 10);
					Run();
				}

			}
			catch (LyncClientException lyncClientException)
			{
				if (lyncClientException.Message.Equals(ExceptionLyncNoclient))
				{
					Log.Info($"Lync Known Exception: no client. Retry in {DelayRetryAuthentication} ms");
					Thread.Sleep(DelayRetryAuthentication);
					Run();
				}
				else
				{
					Log.Warn("Lync Exception", lyncClientException);
				}
			}
			catch (SystemException systemException)
			{
				if (IsLyncException(systemException))
				{
					// Log the exception thrown by the Lync Model API.
					Log.Warn("Lync Exception", systemException);
				}
				else
				{
					Log.Warn("Exception: ", systemException);
					// Rethrow the SystemException which did not come from the Lync Model API.
					throw;
				}
			}
		}

		/// <summary>
		/// Create conversation log file and listen on what participants say
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private static void conversations_ConversationAdded(object sender, ConversationManagerEventArgs e)
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

			using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
			{
				using (StreamWriter writer = new StreamWriter(stream))
				{
					writer.WriteLine(logHeader);
				}
			}

			Conversation conv = e.Conversation;

			//detect all participant (including user)
			conv.ParticipantAdded += (_sender, _e) =>
			{
				var participant = _e.Participant;
				InstantMessageModality remoteImModality = (InstantMessageModality)participant.Modalities[ModalityTypes.InstantMessage];

				//detect all messages (including user's)
				remoteImModality.InstantMessageReceived += (__sender, __e) =>
				{
					Log.Info("message event: " + __e.Text);
					remoteImModality_InstantMessageReceived(__sender, __e, fileLog);
				};
			};

			//get audio conversation informations about user (not the other participants)
			AVModality callImModality = (AVModality)conv.Participants[0].Modalities[ModalityTypes.AudioVideo];
			//notify call 
			callImModality.ModalityStateChanged += (_sender, _e) =>
			{
				Log.Info("call event: " + _e.NewState);
				callImModality_ModalityStateChanged(_e, fileLog + ".mp3");
			};
		}

		/// <summary>
		/// log to fileLog the date of the start and end of a call
		/// (ModalityStateChanged callback)
		/// </summary>
		/// <param name="e"></param>
		/// <param name="fileLog"></param>
		private static void callImModality_ModalityStateChanged(ModalityStateChangedEventArgs e, String fileLog)
		{
			//write log only on connection or disconnection
			if (e.NewState == ModalityState.Connected || e.NewState == ModalityState.Disconnected)
			{
				//write start/end info to log
				using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
				{
					using (StreamWriter writer = new StreamWriter(stream))
					{
						writer.WriteLine(String.Format(LogAudio,
							(e.NewState == ModalityState.Connected) ? "started" : "ended",
							DateTime.Now.ToString("HH:mm:ss")
						));
					}
				}
			}

			//record conversation
			if (e.NewState == ModalityState.Connected)
			{
				Log.Info("Start recording to " + fileLog);
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
		private static void remoteImModality_InstantMessageReceived(object sender, MessageSentEventArgs e, String fileLog)
		{
			InstantMessageModality modality = (InstantMessageModality)sender;

			//gets the participant name
			string name = (string)modality.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName);

			//reads the message in its plain text format (automatically converted)
			string message = e.Text;

			//write message to log
			using (FileStream stream = File.Open(fileLog, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
			{
				using (StreamWriter writer = new StreamWriter(stream))
				{
					if (name.Contains(_nameShortener))
					{
						name = name.Substring(name.IndexOf(_nameShortener) + _nameShortener.Length);
					}
					writer.WriteLine(LogMessage, name, DateTime.Now.ToString("HH:mm:ss"), message);
				}
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
