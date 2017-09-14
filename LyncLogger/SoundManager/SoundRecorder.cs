using CSCore;
using CSCore.Codecs;
using CSCore.Codecs.WAV;
using CSCore.SoundIn;
using CSCore.SoundOut;
using CSCore.Streams;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace LyncLogger.SoundManager
{
	public class SoundRecorder
	{
		//soundcard
		private WasapiCapture _capture;

		private WaveWriter _waveWriter;
		private ISoundOut _soundSilenceOut;
		private IWaveSource _soundSilenceSource;

		//mic
		private WaveIn _waveIn;

		private byte[] _writerBuffer;
		private IWriteable _writer;
		private IWaveSource _source;

		/// <summary>
		/// Record sound made in Mic and save it to a wave file
		/// </summary>
		/// <param name="wavefile">name of the wave file with extension</param>
		public void CaptureMicToWave(string wavefile)
		{
			int i = 0;
			string extension = ".wav";

			foreach (var device in WaveIn.Devices)
			{
				_waveIn = new WaveInEvent(new WaveFormat(44100, 16, device.Channels));
				_waveIn.Device = i++;

				_waveIn.Initialize();
				_waveIn.Start();

				var waveInToSource = new SoundInSource(_waveIn);

				_source = waveInToSource;
				var notifyStream = new SingleBlockNotificationStream(_source);


				_source = notifyStream.ToWaveSource(16);
				_writerBuffer = new byte[_source.WaveFormat.BytesPerSecond];

				wavefile = $"{wavefile.Remove(wavefile.LastIndexOf(extension) - (i > 1 ? 1 : 0))}{i}{extension}";
				_writer = new WaveWriter(wavefile, _source.WaveFormat);
				waveInToSource.DataAvailable += (s, e) =>
				{
					int read = 0;
					while ((read = _source.Read(_writerBuffer, 0, _writerBuffer.Length)) > 0)
					{
						_writer.Write(_writerBuffer, 0, read);
					}
				};
			}
		}

		/// <summary>
		/// Stop recording of the Mic started by CaptureMicToWave
		/// </summary>
		public void UnCaptureMicToWave()
		{
			_waveIn.Stop();
			_source.Dispose();
			_waveIn.Dispose();
			if (_writer is IDisposable)
				((IDisposable)_writer).Dispose();

			_waveIn = null;
			_source = null;
			_writer = null;
		}

		/// <summary>
		/// Capture the audio outputed by the soundcard to a wave file.
		/// The sound you ear is recorded and saved to [wavefile].wav
		/// </summary>
		/// <param name="wavefile">name of the wave file with extension</param>
		/// <param name="captureSilence">if true record blank sounds</param>
		public void CaptureSpeakersToWave(string wavefile, bool captureSilence)
		{
			_capture = new WasapiLoopbackCapture();

			//initialize the selected device for recording
			_capture.Initialize();

			//create a wavewriter to write the data to
			_waveWriter = new WaveWriter(wavefile, _capture.WaveFormat);

			//setup an eventhandler to receive the recorded data
			_capture.DataAvailable += (s, e) =>
			{
				//save the recorded audio
				_waveWriter.Write(e.Data, e.Offset, e.ByteCount);
			};

			//start recording
			_capture.Start();

			if (captureSilence)
			{
				CaptureSilence();
			}
		}

		/// <summary>
		/// stop audio capture started with CaptureAudioToWave
		/// </summary>
		public void UnCaptureSpeakersToWave()
		{
			//Stop silence recording
			if (_soundSilenceOut != null)
			{
				_soundSilenceOut.Stop();
				_soundSilenceOut.Dispose();
				_soundSilenceSource.Dispose();
				_soundSilenceOut = null;
				_soundSilenceSource = null;
			}

			//stop recording
			_capture.Stop();
			_waveWriter.Dispose();
			_waveWriter = null;

			_capture.Dispose();
			_capture = null;
		}


		/// <summary>
		/// Play a blank sound. It keeps recording even when no sound is played.
		/// Loopback callback is triggered only when a sound is played.
		/// Playing a blank sound keep the recorder running even when there's no sound playing
		/// </summary>
		private void CaptureSilence()
		{
			_soundSilenceSource = new SilenceGenerator();

			_soundSilenceOut = GetSoundOut();

			_soundSilenceOut.Initialize(_soundSilenceSource);
			_soundSilenceOut.Play();
		}

		private ISoundOut GetSoundOut()
		{
			if (WasapiOut.IsSupportedOnCurrentPlatform)
				return new WasapiOut();
			else
				return new DirectSoundOut();
		}

		private const int MixerSampleRate = 44100; //44.1kHz

		public void MixerWave(string filesFolder, string fileMix)
		{
			string[] waveFiles = Directory.GetFiles(filesFolder);

			var mixer = new SimpleMixer(waveFiles.Count(), MixerSampleRate) //output: stereo, 44,1kHz
			{
				FillWithZeros = false,
				DivideResult = false //you may play around with this
			};


			foreach (string waveFile in waveFiles)
			{
				try
				{
					mixer.AddSource(
						CodecFactory.Instance.GetCodec(waveFile)
						.ChangeSampleRate(MixerSampleRate)
						.ToStereo()
						.ToSampleSource());
				}
				catch (Exception e)
				{
					Debug.Write("AddSource to mix func error: " + e.Message);
				}
			}

			mixer.ToWaveSource().WriteToFile(fileMix);

			mixer.Dispose();
			mixer = null;

		}


	}
}
