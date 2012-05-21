using System;

namespace DirectShow.Capture
{
	/// <summary>
	///  Exception thrown when the device cannot be rendered or started.
	/// </summary>
	public class DeviceInUseException : Exception
	{
		// Initializes a new instance with the specified HRESULT
		public DeviceInUseException(string deviceName, int hResult)
			: base(deviceName + " is in use or cannot be rendered. HRESULT = " + hResult + "")
		{
		}
	}
}
