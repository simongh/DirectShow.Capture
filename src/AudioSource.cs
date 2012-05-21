using System;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Represents a physical connector or source on an 
	///  audio device. This class is used on filters that
	///  support the IAMAudioInputMixer interface such as 
	///  source cards.
	/// </summary>
	public class AudioSource : Source
	{
		private IPin _Pin;			// audio mixer interface (COM object)

		/// <summary> Enable or disable this source. For audio sources it is 
		/// usually possible to enable several sources. When setting Enabled=true,
		/// set Enabled=false on all other audio sources. </summary>
		public override bool Enabled
		{
			get
			{
				IAMAudioInputMixer mix = (IAMAudioInputMixer)_Pin;
				bool e;
				mix.get_Enable(out e);
				return (e);
			}
			set
			{
				IAMAudioInputMixer mix = (IAMAudioInputMixer)_Pin;
				mix.put_Enable(value);
			}
		}

		internal AudioSource(IPin pin)
		{
			if ((pin as IAMAudioInputMixer) == null)
				throw new NotSupportedException("The input pin does not support the IAMAudioInputMixer interface");
			
			_Pin = pin;
			Name = getName(pin);
		}

		/// <summary> Retrieve the friendly name of a connectorType. </summary>
		private string getName(IPin pin)
		{
			string s = "Unknown pin";
			PinInfo pinInfo = new PinInfo();

			// Direction matches, so add pin name to listbox
			int hr = pin.QueryPinInfo(out pinInfo);
			Marshal.ThrowExceptionForHR(hr);
			s = pinInfo.name + "";

			// The pininfo structure contains a reference to an IBaseFilter,
			// so you must release its reference to prevent resource a leak.
			if (pinInfo.filter != null)
				Marshal.ReleaseComObject(pinInfo.filter); 
			pinInfo.filter = null;

			return (s);
		}

		/// <summary> Release unmanaged resources. </summary>
		protected override void Dispose(bool disposing)
		{
			if (_Pin != null)
				Marshal.ReleaseComObject(_Pin);
			_Pin = null;
			base.Dispose(disposing);
		}
	}
}
