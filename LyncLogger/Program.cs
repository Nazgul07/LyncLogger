using System;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Win32;

namespace LyncLogger
{
	internal class Program
	{
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

		private static void Main(string[] args)
		{
			//folder to log conversations
			string logFolder = Environment.ExpandEnvironmentVariables(SettingsManager.ReadSetting("logfolder"));

			// create log directory if missing
			CreateDirectoryIfMissing(logFolder);

			//-- -- -- Add notification icon
			NotifyIconSystray.AddNotifyIcon("Lync Logger", new MenuItem[] {
				new MenuItem("Lync History", (s, e) => { Process.Start(logFolder); }),
				new MenuItem("Switch Audio logger On/Off", (s, e) => {  AudioLogger.Instance.Switch(); })
			});

			//-- -- -- Handles Sound record operations

			RegisterKey("Software\\LyncLogger", "Audio", "Activated");
			AudioLogger.Instance.Initialize(logFolder);

			//-- -- -- Handles LYNC operations
			System.Threading.Tasks.Task.Factory.StartNew(() =>
			{
				try
				{
					LyncLogger.Run(logFolder);
				}
				catch (Exception ex)
				{
					//exit app properly
					NotifyIconSystray.DisposeNotifyIcon();
				}
			});

			//prevent the application from exiting right away
			Application.Run();
		}


		/// <summary>
		/// Create registry key to keep settings of recording for audio
		/// If registry key already exists, set AudioLogger
		/// </summary>
		/// <param name="keyName"></param>
		/// <param name="valueName"></param>
		/// <param name="value"></param>
		private static void RegisterKey(string keyName, string valueName, string value)
		{
			RegistryKey key = Registry.CurrentUser;
			RegistryKey lyncLoggerKey = key.OpenSubKey(keyName);
			if (lyncLoggerKey != null)
			{
				AudioLogger.Instance.IsAllowedRecording = ((string)lyncLoggerKey.GetValue(valueName) == value);
				lyncLoggerKey.Close();
			}
			else
			{
				RegistryKey subkey = key.CreateSubKey(keyName);
				subkey.SetValue(valueName, value);
				subkey.Close();
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
	}
}
