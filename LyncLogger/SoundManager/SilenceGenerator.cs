using CSCore;
using System;

namespace LyncLogger.SoundManager
{
	public class SilenceGenerator : IWaveSource
	{
		public int Read(byte[] buffer, int offset, int count)
		{
			Array.Clear(buffer, offset, count);
			return count;
		}

		public WaveFormat WaveFormat { get; } = new WaveFormat(44100, 16, 2);

		public long Position
		{
			get => -1;
			set => throw new InvalidOperationException();
		}

		public long Length => -1;

		public void Dispose()
		{
			//do nothing
		}
	}
}
