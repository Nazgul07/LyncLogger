using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Threading;
using AdysTech.CredentialManager;
using Microsoft.Exchange.WebServices.Data;
using Notifications.Wpf;
using Task = System.Threading.Tasks.Task;

namespace LyncLogger
{
	internal class Program
	{
		private const string Enable365 = "Enable Office 365 Integration";
		private const string Disable365 = "Disable Office 365 Integration";

		
		/// <summary>
		/// create directory if doesnt exist
		/// </summary>
		/// <param name="folder"></param>
		private static void CreateDirectoryIfMissing(String folder)
		{
			if (!Directory.Exists(folder))
			{
				Directory.CreateDirectory(folder);
			}
		}
		private static readonly Mutex Mutex = new Mutex(true, "{d871d1b7-1b3c-4434-b1fa-f7bde92dd8fd}");
		[STAThread]
		private static void Main(string[] args)
		{
			if (Mutex.WaitOne(TimeSpan.Zero, true))
			{
				Notifications.Dispatcher = Dispatcher.CurrentDispatcher;
				//folder to log conversations
				string logFolder = Environment.ExpandEnvironmentVariables(SettingsManager.ReadSetting("logfolder"));

				// create log directory if missing
				CreateDirectoryIfMissing(logFolder);

				//-- -- -- Handles Sound record operations

				InitializeAudioLoggerStatus();
				AudioLogger.Instance.Initialize(logFolder);

				//-- -- -- Add notification icon
				NotifyIconSystray.AddNotifyIcon("Lync Logger", new[]
				{
					new MenuItem("Lync History", (s, e) => { Process.Start(logFolder); }),
					new MenuItem($"Turn Notifications {(Notifications.Enabled ? "Off" : "On")}",
						(s, e) =>
						{
							Notifications.Enabled = !Notifications.Enabled;
							((MenuItem) s).Text = $"Turn Notifications {(Notifications.Enabled ? "Off" : "On")}";
						}),
					new MenuItem($"Switch Audio Logger {(AudioLogger.Instance.IsAllowedRecording ? "Off" : "On")}",
						(s, e) => { SwitchAudio((MenuItem) s); }),
					new MenuItem($@"{
							(Validate365Credentials(new NetworkCredential(SettingsManager.ReadSetting("office365username"),
								SecureCredentials.DecryptString(SettingsManager.ReadSetting("office365password"))))
								? "Disable"
								: "Enable")
						} Office 365 Integration", (s, e) =>
					{
						AuthenticateWithOffice365((MenuItem) s);
					})
				});


				Task.Factory.StartNew(() =>
				{
					try
					{
						LyncLogger.Run(logFolder);
					}
					catch
					{
						//exit app properly
						NotifyIconSystray.DisposeNotifyIcon();
					}
				});

				Application.Run();
				Mutex.ReleaseMutex();
			}
		}


		/// <summary>
		/// Set AudioLogger On/Off by checking the current status
		/// </summary>
		private static void InitializeAudioLoggerStatus()
		{
			string loggerStatus = SettingsManager.ReadSetting("AudioLoggerStatus");
			AudioLogger.Instance.IsAllowedRecording = loggerStatus.ToUpper() == "ON";
		}

		public static bool Validate365Credentials(NetworkCredential cred)
		{
			try
			{
				var ewsProxy = new ExchangeService() { Url = new Uri("https://outlook.office365.com/ews/exchange.asmx") };
				ewsProxy.Credentials = new NetworkCredential(cred.UserName, cred.Password);
				ewsProxy.FindFolders(WellKnownFolderName.Root, new FolderView(1));
				return true;
			}
			catch
			{
				return false;
			}
		}

		private static void AuthenticateWithOffice365(MenuItem menu, string message = "")
		{
			bool save = false;
			if (!Validate365Credentials(new NetworkCredential(SettingsManager.ReadSetting("office365username"),
				SecureCredentials.DecryptString(SettingsManager.ReadSetting("office365password")))))
			{
				var cred = CredentialManager.PromptForCredentials("Office 365", ref save, message,
					"Credentials for Office 365");
				if (cred == null) return;

				if (Validate365Credentials(cred))
				{
					SettingsManager.AddUpdateAppSettings("office365username", cred.UserName);
					SettingsManager.AddUpdateAppSettings("office365password",
						SecureCredentials.EncryptString(SecureCredentials.ToSecureString(cred.Password)));
					menu.Text = Disable365;
					
					 Notifications.Send("Connection to Office 365 Established", NotificationType.Success);
				}
				else
				{
					AuthenticateWithOffice365(menu, "Invalid Credentials. Try again.");
				}
			}
			else
			{
				SettingsManager.AddUpdateAppSettings("office365username", "");
				SettingsManager.AddUpdateAppSettings("office365password", "");
				menu.Text = Enable365;
				Notifications.Send("Office 365 Integration Disabled", NotificationType.Information);
			}
		}

		/// <summary>
		/// Activate or Deactivate audio recording
		/// </summary>
		private static void SwitchAudio(MenuItem menu)
		{
			AudioLogger.Instance.IsAllowedRecording = !AudioLogger.Instance.IsAllowedRecording;
			string status = (AudioLogger.Instance.IsAllowedRecording ? "On" : "Off");

			SettingsManager.AddUpdateAppSettings("AudioLoggerStatus", status);
			
			menu.Text = $"Switch Audio logger {(AudioLogger.Instance.IsAllowedRecording ? "Off" : "On")}";

			Notifications.Send($"Audio Logging {(AudioLogger.Instance.IsAllowedRecording ? "Enabled" : "Disabled")}", NotificationType.Information);
		}
	}
}
