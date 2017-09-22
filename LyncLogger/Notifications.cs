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
		}

		internal static void SendHyperlink(string link)
		{
			if (HyperlinksEnabled)
			{
				Dispatcher.Invoke(() => {
					NotificationManager notificationManager = new NotificationManager();
					notificationManager.Show(new NotificationContent
					{
						Title = "Lync Logger Hyperlink",
						Message = link,
					}, "", TimeSpan.FromMilliseconds(Convert.ToInt16(SettingsManager.ReadSetting("HyperlinkNotificationsTimeout")) * 1000), () => { Process.Start(link); });
				});
			}
		}
	}
}
