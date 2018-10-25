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
				
				//-- -- -- Add notification icon
				NotifyIconSystray.AddNotifyIcon("Lync Logger", new[]
				{
					new MenuItem("View Lync History Folder", (s, e) => { Process.Start(logFolder); }),
					new MenuItem($"{(LyncLogger.TextLoggingEnabled ? "Disable" : "Enable")} Text File Logging",
						(s, e) =>
						{
							LyncLogger.TextLoggingEnabled = !LyncLogger.TextLoggingEnabled;
							((MenuItem) s).Text = $"{(LyncLogger.TextLoggingEnabled ? "Disable" : "Enable")} Text File Logging";
							Notifications.Send($"Text File Logging {(LyncLogger.TextLoggingEnabled ? "Enabled" : "Disabled")}", NotificationType.Information);
						}),
					new MenuItem($"{(Notifications.HyperlinksEnabled ? "Disable" : "Enable")} Hyperlink Notifications",
						(s, e) =>
						{
							Notifications.HyperlinksEnabled = !Notifications.HyperlinksEnabled;
							((MenuItem) s).Text = $"{(Notifications.HyperlinksEnabled ? "Disable" : "Enable")} Hyperlink Notifications";
						}),
					new MenuItem($"{(Notifications.Enabled ? "Disable" : "Enable")} Notifications",
						(s, e) =>
						{
							Notifications.Enabled = !Notifications.Enabled;
							((MenuItem) s).Text = $"{(Notifications.Enabled ? "Disable" : "Enable")} Notifications";
						}),
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
	}
}
