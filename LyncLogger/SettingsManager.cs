﻿using System;
using System.Configuration;
using System.Reflection;

namespace LyncLogger
{
	public class SettingsManager
	{
		/// <summary>
		/// read setting from app.config
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public static string ReadSetting(string key)
		{
			string result = string.Empty;
			try
			{
				var appSettings = ConfigurationManager.AppSettings;
				if (appSettings[key] != null)
				{
					result = appSettings[key];
				}
			}
			catch (ConfigurationErrorsException)
			{
				
			}
			return result;
		}

		/// <summary>
		/// add or update setting to app.config
		/// </summary>
		/// <param name="key"></param>
		/// <param name="value"></param>
		public static void AddUpdateAppSettings(string key, string value)
		{
			try
			{
				var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
				var settings = configFile.AppSettings.Settings;
				if (settings[key] == null)
				{
					settings.Add(key, value);
				}
				else
				{
					settings[key].Value = value;
				}
				configFile.Save(ConfigurationSaveMode.Modified);
				ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
			}
			catch (ConfigurationErrorsException)
			{
				Console.WriteLine("Error writing app settings");
			}
		}
	}
}
