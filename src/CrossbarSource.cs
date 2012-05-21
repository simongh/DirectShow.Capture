using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Represents a physical connector or source on an 
	///  audio/video device. This class is used on filters that
	///  support the IAMCrossbar interface such as TV Tuners.
	/// </summary>
	public class CrossbarSource : Source
	{
		private int _InputPin;			// input pin number on the crossbar
		private PhysicalConnectorType _ConnectorType;		// type of the connector

		/// <summary>
		/// Gets the crossbar filter
		/// </summary>
		internal IAMCrossbar Crossbar
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the output pin
		/// </summary>
		internal int OutputPin
		{
			get;
			private set;
		}

		/// <summary>Gets or sets whether this source is enabled.</summary>
		public override bool Enabled
		{
			get
			{
				int i;
				return Crossbar.get_IsRoutedTo(OutputPin, out i) == 0 && _InputPin == i;
			}
			set
			{
				if (value)
				{
					// Enable this route
					int hr = Crossbar.Route(OutputPin, _InputPin);
					Marshal.ThrowExceptionForHR(hr);
				}
				else
				{
					// Disable this route by routing the output
					// pin to input pin -1
					int hr = Crossbar.Route(OutputPin, -1);
					Marshal.ThrowExceptionForHR(hr);
				}
			}
		}

		internal CrossbarSource(IAMCrossbar crossbar, int outputPin, int inputPin, PhysicalConnectorType connectorType)
		{
			Crossbar = crossbar;
			OutputPin = outputPin;
			_InputPin = inputPin;
			_ConnectorType = connectorType;
			this.Name = getName(connectorType);
		}

		/// <summary> Retrieve the friendly name of a connectorType. </summary>
		private string getName(PhysicalConnectorType connectorType)
		{
			string name;
			switch (connectorType)
			{
				case PhysicalConnectorType.Video_Tuner:
					name = "Video Tuner";
					break;
				case PhysicalConnectorType.Video_Composite:
					name = "Video Composite";
					break;
				case PhysicalConnectorType.Video_SVideo:
					name = "Video S-Video";
					break;
				case PhysicalConnectorType.Video_RGB:
					name = "Video RGB";
					break;
				case PhysicalConnectorType.Video_YRYBY:
					name = "Video YRYBY";
					break;
				case PhysicalConnectorType.Video_SerialDigital:
					name = "Video Serial Digital";
					break;
				case PhysicalConnectorType.Video_ParallelDigital:
					name = "Video Parallel Digital";
					break;
				case PhysicalConnectorType.Video_SCSI:
					name = "Video SCSI";
					break;
				case PhysicalConnectorType.Video_AUX:
					name = "Video AUX";
					break;
				case PhysicalConnectorType.Video_1394:
					name = "Video Firewire";
					break;
				case PhysicalConnectorType.Video_USB:
					name = "Video USB";
					break;
				case PhysicalConnectorType.Video_VideoDecoder:
					name = "Video Decoder";
					break;
				case PhysicalConnectorType.Video_VideoEncoder:
					name = "Video Encoder";
					break;
				case PhysicalConnectorType.Video_SCART:
					name = "Video SCART";
					break;

				case PhysicalConnectorType.Audio_Tuner:
					name = "Audio Tuner";
					break;
				case PhysicalConnectorType.Audio_Line:
					name = "Audio Line In";
					break;
				case PhysicalConnectorType.Audio_Mic:
					name = "Audio Mic";
					break;
				case PhysicalConnectorType.Audio_AESDigital:
					name = "Audio AES Digital";
					break;
				case PhysicalConnectorType.Audio_SPDIFDigital:
					name = "Audio SPDIF Digital";
					break;
				case PhysicalConnectorType.Audio_SCSI:
					name = "Audio SCSI";
					break;
				case PhysicalConnectorType.Audio_AUX:
					name = "Audio AUX";
					break;
				case PhysicalConnectorType.Audio_1394:
					name = "Audio Firewire";
					break;
				case PhysicalConnectorType.Audio_USB:
					name = "Audio USB";
					break;
				case PhysicalConnectorType.Audio_AudioDecoder:
					name = "Audio Decoder";
					break;

				default:
					name = "Unknown Connector";
					break;
			}
			return name;
		}

		/// <summary> Release unmanaged resources. </summary>
		protected override void Dispose(bool disposing)
		{
			base.Dispose(disposing);

			if (Crossbar != null)
				Marshal.ReleaseComObject(Crossbar);
			Crossbar = null;
		}
	}
}
