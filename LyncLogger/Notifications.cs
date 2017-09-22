using System;
using System.Diagnostics;
using System.Windows.Threading;
using Notifications.Wpf;

namespace LyncLogger
{
	internal class Notifications
	{
		internal static Dispatcher Dispatcher { get; set; }
		internal static bool Enabled
		{
			get => SettingsManager.ReadSetting("ShowNotifications").ToUpper() == "TRUE";
			set => SettingsManager.AddUpdateAppSettings("ShowNotifications", value ? "True" : "False");
		}

		internal static bool HyperlinksEnabled
		{
			get => SettingsManager.ReadSetting("ShowHyperlinkNotifications").ToUpper() == "TRUE";
			set => SettingsManager.AddUpdateAppSettings("ShowHyperlinkNotifications", value ? "True" : "False");
		}

		internal static void Send(string message, NotificationType type)
		{
			if (Enabled)
			{
				SendNoCheck(message, type);
			}
		}

		private static void SendNoCheck(string message, NotificationType type)
		{
			Dispatcher.Invoke(() =>
			{
				try
				{
					NotificationManager notificationManager = new NotificationManager();
					notificationManager.Show(new NotificationContent
					{
						Title = "Lync Logger",
						Message = message,
						Type = type,
					}, "", TimeSpan.FromMilliseconds(2500));
				}
				catch
				{
					//ignore
				}
			});
		}

		internal static void SendHyperlink(string link)
		{
			if (HyperlinksEnabled)
			{
				try
				{
					Dispatcher.Invoke(() =>
					{
						NotificationManager notificationManager = new NotificationManager();
						notificationManager.Show(new NotificationContent
							{
								Title = "Lync Logger Hyperlink",
								Message = link,
							}, "",
							TimeSpan.FromMilliseconds(Convert.ToInt16(SettingsManager.ReadSetting("HyperlinkNotificationsTimeout")) * 1000),
							() =>
							{
								System.Threading.Tasks.Task.Factory.StartNew(() => { 
									try { 
										Process.Start(link);
									}
									catch (Exception e)
									{
										SendNoCheck(e.Message, NotificationType.Error);
								}
								});
							});
					});
				}
				catch
				{
					//ignore
				}
			}
		}
	}
}
