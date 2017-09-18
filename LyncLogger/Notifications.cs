using System;
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
	}
}
