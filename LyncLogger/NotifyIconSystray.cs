using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace LyncLogger
{
	/// <summary>
	/// Handles the systray icon
	/// </summary>
	internal static class NotifyIconSystray
	{
		private static NotifyIcon _notifyIcon;
		public delegate void Status(bool status);
		private static string _name;
		public delegate void CallbackQuit();
		public static event CallbackQuit OnQuit;

		private static string _onText = "on";
		private static string _onImage = "icon.ico";
		public static void SetOnIcon(string text, string imageName)
		{
			_onText = text;
			_onImage = imageName;
		}

		private static string _offText = "off";
		private static string _offImage = "icon_off.ico";
		public static void SetOffIcon(string text, string imageName)
		{
			_offText = text;
			_offImage = imageName;
		}

		/// <summary>
		/// This method allows to change the state of icon and tooltip
		/// true = Log Active: the logger detected the client and is active.
		/// </summary>
		/// <param name="status"></param>
		public static void Status_DelegateMethod(bool status)
		{
			string text = $"{_name}\nstatus: {(status ? _onText : _offText)}";

			string iconName = status ? _onImage : _offImage;

			SetNotifyIcon(iconName, text);
		}

		/// <summary>
		/// This delegate allows us to call Status_DelegateMethod in the backgroundworker
		/// It changes the indicator that displays the state of the app.
		/// </summary>
		public static Status ChangeStatus = Status_DelegateMethod;

		/// <summary>
		/// set text and icon for the taskbar
		/// Icon must be in the project as embedded resource
		/// </summary>
		/// <param name="nameIcon"></param>
		/// <param name="text"></param>
		public static void SetNotifyIcon(string iconName, string text)
		{
			//set text that support 128 char instead of 64
			Fixes.SetNotifyIconText(_notifyIcon, text);

			//get icon by its name. Icon must be in the project as embedded resource
			Assembly assembly = Assembly.GetExecutingAssembly();
			string ns = assembly.EntryPoint.DeclaringType.Namespace;
			using (Stream iconStream = assembly.GetManifestResourceStream($"{ns}.{iconName}"))
				_notifyIcon.Icon = new Icon(iconStream);
		}


		/// <summary>
		/// add notification icon to system tray bar (near the clock)
		/// quit option is available by default
		/// </summary>
		/// <param name="name">name displayed on mouse hover</param>
		/// <param name="items">items to add to the context menu</param>
		public static void AddNotifyIcon(String name, MenuItem[] items = null)
		{
			_name = name;
			_notifyIcon = new NotifyIcon();
			_notifyIcon.Visible = true;

			Status_DelegateMethod(false); //set name and icon

			ContextMenu contextMenu1 = new ContextMenu();
			if (items != null)
			{
				contextMenu1.MenuItems.AddRange(items);
			}

			contextMenu1.MenuItems.Add(new MenuItem("Quit", (s, e) =>
			{
				OnQuit?.Invoke();
				DisposeNotifyIcon();
			}));
			_notifyIcon.ContextMenu = contextMenu1;
		}
		
		public static void DisposeNotifyIcon()
		{
			_notifyIcon.Dispose();
			foreach (ProcessThread thread in Process.GetCurrentProcess().Threads)
			{
				thread.Dispose();
			}
			Process.GetCurrentProcess().Kill();
		}
	}

	public class Fixes
	{
		/// <summary>
		/// Set text tooltip to 128 char limit instead of 64
		/// http://stackoverflow.com/questions/579665/how-can-i-show-a-systray-tooltip-longer-than-63-chars
		/// </summary>
		/// <param name="ni"></param>
		/// <param name="text"></param>
		public static void SetNotifyIconText(NotifyIcon ni, string text)
		{
			if (text.Length >= 128) throw new ArgumentOutOfRangeException("Text limited to 127 characters");

			Type t = typeof(NotifyIcon);
			BindingFlags hidden = BindingFlags.NonPublic | BindingFlags.Instance;
			t.GetField("text", hidden).SetValue(ni, text);

			if ((bool)t.GetField("added", hidden).GetValue(ni))
				t.GetMethod("UpdateIcon", hidden).Invoke(ni, new object[] { true });
		}
	}
}
