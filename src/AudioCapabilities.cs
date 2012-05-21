using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	public class AudioCapabilities
	{

		/// <summary>Gets the minimum number of audio channels. </summary>
		public int MinimumChannels
		{
			get;
			private set;
		}

		/// <summary>Gets the maximum number of audio channels. </summary>
		public int MaximumChannels
		{
			get;
			private set;
		}

		/// <summary>Gets the granularity of the channels. For example, channels 2 through 4, in steps of 2. </summary>
		public int ChannelsGranularity
		{
			get;
			private set;
		}

		/// <summary>Gets the minimum number of bits per sample. </summary>
		public int MinimumSampleSize
		{
			get;
			private set;
		}

		/// <summary> Maximum number of bits per sample. </summary>
		public int MaximumSampleSize
		{
			get;
			private set;
		}

		/// <summary>Gets the granularity of the bits per sample. For example, 8 bits per sample through 32 bits per sample, in steps of 8. </summary>
		public int SampleSizeGranularity
		{
			get;
			private set;
		}

		/// <summary>Gets the minimum sample frequency. </summary>
		public int MinimumSamplingRate
		{
			get;
			private set;
		}

		/// <summary>Gets the maximum sample frequency. </summary>
		public int MaximumSamplingRate
		{
			get;
			private set;
		}

		/// <summary>Gets the granularity of the frequency. For example, 11025 Hz to 44100 Hz, in steps of 11025 Hz. </summary>
		public int SamplingRateGranularity
		{
			get;
			private set;
		}

		/// <summary> Retrieve capabilities of an audio device </summary>
		internal AudioCapabilities(IAMStreamConfig audioStreamConfig)
		{
			if (audioStreamConfig == null)
				throw new ArgumentNullException("audioStreamConfig");

			AMMediaType mediaType = null;
			AudioStreamConfigCaps caps = null;
			IntPtr pCaps = IntPtr.Zero;
			//IntPtr pMediaType;
			try
			{
				// Ensure this device reports capabilities
				int c, size;
				int hr = audioStreamConfig.GetNumberOfCapabilities(out c, out size);
				if (hr != 0) Marshal.ThrowExceptionForHR(hr);
				if (c <= 0)
					throw new NotSupportedException("This audio device does not report capabilities.");
				if (size > Marshal.SizeOf(typeof(AudioStreamConfigCaps)))
					throw new NotSupportedException("Unable to retrieve audio device capabilities. This audio device requires a larger AudioStreamConfigCaps structure.");
				if (c > 1)
					Debug.WriteLine("WARNING: This audio device supports " + c + " capability structures. Only the first structure will be used.");

				// Alloc memory for structure
				pCaps = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(AudioStreamConfigCaps)));

				// Retrieve first (and hopefully only) capabilities struct
				hr = audioStreamConfig.GetStreamCaps(0, out mediaType, pCaps);
				if (hr != 0) Marshal.ThrowExceptionForHR(hr);

				// Convert pointers to managed structures
				//mediaType = (AMMediaType)Marshal.PtrToStructure(pMediaType, typeof(AMMediaType));
				caps = (AudioStreamConfigCaps)Marshal.PtrToStructure(pCaps, typeof(AudioStreamConfigCaps));

				// Extract info
				MinimumChannels = caps.MinimumChannels;
				MaximumChannels = caps.MaximumChannels;
				ChannelsGranularity = caps.ChannelsGranularity;
				MinimumSampleSize = caps.MinimumBitsPerSample;
				MaximumSampleSize = caps.MaximumBitsPerSample;
				SampleSizeGranularity = caps.BitsPerSampleGranularity;
				MinimumSamplingRate = caps.MinimumSampleFrequency;
				MaximumSamplingRate = caps.MaximumSampleFrequency;
				SamplingRateGranularity = caps.SampleFrequencyGranularity;

			}
			finally
			{
				if (pCaps != IntPtr.Zero)
					Marshal.FreeCoTaskMem(pCaps); pCaps = IntPtr.Zero;
				if (mediaType != null)
					DsUtils.FreeAMMediaType(mediaType); mediaType = null;
			}
		}
	}
}
