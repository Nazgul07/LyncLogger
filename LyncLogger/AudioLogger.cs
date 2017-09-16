using System;
using LyncLogger.SoundManager;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
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
			catch
			{
				// ignored
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
			catch
			{
				// ignored
			}

			try
			{
				_soundRecorder.MixerWave(TempFolder, Path.Combine(_folderLog, _fileLog));
			}
			catch
			{
				// ignored
			}

			try
			{
				Directory.Delete(TempFolder, true);
			}
			catch
			{
				// ignored
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