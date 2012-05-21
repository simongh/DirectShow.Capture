using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	public class VideoCapabilities
	{
		/// <summary>Gets the native size of the incoming video signal. This is the largest signal the filter can digitize with every pixel remaining unique.</summary>
		public Size InputSize
		{
			get;
			private set;
		}

		/// <summary>Gets the minimum supported frame size.</summary>
		public Size MinFrameSize
		{
			get;
			private set;
		}

		/// <summary>Gets the maximum supported frame size.</summary>
		public Size MaxFrameSize
		{
			get;
			private set;
		}

		/// <summary>Gets the granularity of the output width. This value specifies the increments that are valid between MinFrameSize and MaxFrameSize.</summary>
		public int FrameSizeGranularityX
		{
			get;
			private set;
		}

		/// <summary>Gets the granularity of the output height. This value specifies the increments that are valid between MinFrameSize and MaxFrameSize.</summary>
		public int FrameSizeGranularityY
		{
			get;
			private set;
		}

		/// <summary>Gets the minimum supported frame rate.</summary>
		public double MinFrameRate
		{
			get;
			private set;
		}

		/// <summary>Gets the maximum supported frame rate.</summary>
		public double MaxFrameRate
		{
			get;
			private set;
		}

		/// <summary> Retrieve capabilities of a video device </summary>
		internal VideoCapabilities(IAMStreamConfig videoStreamConfig)
		{
			if (videoStreamConfig == null)
				throw new ArgumentNullException("videoStreamConfig");

			AMMediaType mediaType = null;
			VideoStreamConfigCaps caps = null;
			IntPtr pCaps = IntPtr.Zero;
			//IntPtr pMediaType;
			try
			{
				// Ensure this device reports capabilities
				int c, size;
				int hr = videoStreamConfig.GetNumberOfCapabilities(out c, out size);
				Marshal.ThrowExceptionForHR(hr);

				if (c <= 0)
					throw new NotSupportedException("This video device does not report capabilities.");
				if (size > Marshal.SizeOf(typeof(VideoStreamConfigCaps)))
					throw new NotSupportedException("Unable to retrieve video device capabilities. This video device requires a larger VideoStreamConfigCaps structure.");
				if (c > 1)
					Debug.WriteLine("This video device supports " + c + " capability structures. Only the first structure will be used.");

				// Alloc memory for structure
				pCaps = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(VideoStreamConfigCaps)));

				// Retrieve first (and hopefully only) capabilities struct
				hr = videoStreamConfig.GetStreamCaps(0, out mediaType, pCaps);
				Marshal.ThrowExceptionForHR(hr);

				// Convert pointers to managed structures
				//mediaType = (AMMediaType)Marshal.PtrToStructure(pMediaType, typeof(AMMediaType));
				caps = (VideoStreamConfigCaps)Marshal.PtrToStructure(pCaps, typeof(VideoStreamConfigCaps));

				// Extract info
				InputSize = caps.InputSize;
				MinFrameSize = caps.MinOutputSize;
				MaxFrameSize = caps.MaxOutputSize;
				FrameSizeGranularityX = caps.OutputGranularityX;
				FrameSizeGranularityY = caps.OutputGranularityY;
				MinFrameRate = (double)10000000 / caps.MaxFrameInterval;
				MaxFrameRate = (double)10000000 / caps.MinFrameInterval;
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
