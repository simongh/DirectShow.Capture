using System;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Control and query a hardware TV Tuner.
	/// </summary>
	public class Tuner : IDisposable
	{
		protected IAMTVTuner tvTuner = null;
		private bool _IsDisposed;

		/// <summary>
		///  Get or set the TV Tuner channel.
		/// </summary>
		public int Channel
		{
			get
			{
				int channel;
				AMTunerSubChannel v, a;
				int hr = tvTuner.get_Channel(out channel, out v, out a);
				Marshal.ThrowExceptionForHR(hr);

				return channel;
			}
			set
			{
				int hr = tvTuner.put_Channel(value, AMTunerSubChannel.Default, AMTunerSubChannel.Default);
				Marshal.ThrowExceptionForHR(hr);
			}
		}

		/// <summary>
		///  Get or set the tuner frequency (cable or antenna).
		/// </summary>
		public TunerInputType InputType
		{
			get
			{
				DirectShowLib.TunerInputType t;
				int hr = tvTuner.get_InputType(0, out t);
				Marshal.ThrowExceptionForHR(hr);

				return (TunerInputType)t;
			}
			set
			{
				var t = (DirectShowLib.TunerInputType)value;
				int hr = tvTuner.put_InputType(0, t);
				Marshal.ThrowExceptionForHR(hr);
			}
		}

		/// <summary>
		///  Indicates whether a signal is present on the current channel.
		///  If the signal strength cannot be determined, a NotSupportedException
		///  is thrown.
		/// </summary>
		public bool SignalPresent
		{
			get
			{
				AMTunerSignalStrength sig;
				int hr = tvTuner.SignalPresent(out sig);
				Marshal.ThrowExceptionForHR(hr);

				if (sig == AMTunerSignalStrength.HasNoSignalStrength) 
					throw new NotSupportedException("Signal strength not available.");
				
				return sig == AMTunerSignalStrength.SignalPresent;
			}
		}

		/// <summary> Initialize this object with a DirectShow tuner </summary>
		public Tuner(IAMTVTuner tuner)
		{
			tvTuner = tuner;
		}

		~Tuner()
		{
			if (!_IsDisposed)
				Dispose(false);
		}

		public void Dispose()
		{
			if (!_IsDisposed)
			{
				Dispose(true);
				_IsDisposed = true;
			}
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (tvTuner != null)
				Marshal.ReleaseComObject(tvTuner); tvTuner = null;
		}
	}
}
