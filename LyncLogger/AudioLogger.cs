using System;
using LyncLogger.SoundManager;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
using log4net;
using System.Reflection;

namespace LyncLogger
{
	internal class AudioLogger
	{
		private static readonly string TempFolder = Environment.ExpandEnvironmentVariables("%temp%\\lyncloggeraudio");
		private static readonly string WaveFilename = TempFolder + "\\captureSpeakers.wav";
		private static readonly string MicFilename = TempFolder + "\\captureMic.wav";
		private static string _folderLog = "";
		private static string _fileLog = "";

		private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		private SoundRecorder _soundRecorder;

		public bool IsAllowedRecording { get; set; }

		private static AudioLogger _instance;

		private AudioLogger() { }

		public static AudioLogger Instance
		{
			get
			{
				if (_instance == null)
				{
					_instance = new AudioLogger();
				}
				return _instance;
			}
		}

		internal void Initialize(string folderLog)
		{
			_folderLog = folderLog;

			_soundRecorder = new SoundRecorder();
		}

		/// <summary>
		/// start recording
		/// </summary>
		/// <param name="fileLog"></param>
		public void Start(string fileLog)
		{
			if (!IsAllowedRecording)
			{
				Log.Warn("recording not active (known)");
				return;
			}
			_fileLog = fileLog;

			if (!Directory.Exists(TempFolder))
				Directory.CreateDirectory(TempFolder);

			try
			{
				_soundRecorder.CaptureSpeakersToWave(WaveFilename, true);

				_soundRecorder.CaptureMicToWave(MicFilename);
			}
			catch (Exception ex)
			{
				Log.Error("error starting audio recording", ex);
			}
		}

		/// <summary>
		/// stop audio recording
		/// </summary>
		public void Stop()
		{
			if (!IsAllowedRecording)
				return;

			try
			{
				_soundRecorder.UnCaptureSpeakersToWave();
				_soundRecorder.UnCaptureMicToWave();
			}
			catch (Exception ex)
			{
				Log.Error("error canceling audio recording", ex);
			}

			try
			{
				Log.Info($"build audio record: {Path.Combine(_folderLog, _fileLog)} :");
				_soundRecorder.MixerWave(TempFolder, Path.Combine(_folderLog, _fileLog));
			}
			catch (Exception ex)
			{
				Log.Error($"error building audio record file {Path.Combine(_folderLog, _fileLog)} :", ex);
			}

			try
			{
				Directory.Delete(TempFolder, true);
			}
			catch (Exception ex)
			{
				Log.Error("error deleting record temp folder", ex);
			}
		}

		/// <summary>
		/// Activate or Deactivate audio recording
		/// </summary>
		internal void Switch()
		{
			IsAllowedRecording = !IsAllowedRecording;
			string status = (IsAllowedRecording ? "Activated" : "Deactivated");

			RegistryKey lyncLoggerKey = Registry.CurrentUser.OpenSubKey("LyncLogger");
			if (lyncLoggerKey != null)
			{
				lyncLoggerKey.SetValue("Audio", status);
				lyncLoggerKey.Close();
			}

			MessageBox.Show("Audio logger is " + status);
		}
	}
}