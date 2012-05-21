using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	public class SourceCollection : System.Collections.ObjectModel.ReadOnlyCollection<Source>, IDisposable
	{
		private bool _IsDisposed;

		/// <summary>
		///  Gets or sets the source/physical connector currently in use.
		///  This is marked internal so that the Capture class can control
		///  how and when the source is changed.
		/// </summary>
		internal Source CurrentSource
		{
			get { return this.FirstOrDefault(x => x.Enabled); }
			set
			{
				if (value == null)
				{
					// Disable all sources
					foreach (Source s in this)
						s.Enabled = false;
				}
				else if (value is CrossbarSource)
				{
					// Enable this source
					// (this will automatically disable all other sources)
					value.Enabled = true;
				}
				else
				{
					// Disable all sources
					// Enable selected source
					foreach (Source s in this)
						s.Enabled = false;
					value.Enabled = true;
				}
			}
		}

		/// <summary> Initialize collection with no sources. </summary>
		internal SourceCollection()
			: base(new List<Source>())
		{ }

		/// <summary> Initialize collection with sources from graph. </summary>
		internal SourceCollection(ICaptureGraphBuilder2 graphBuilder, IBaseFilter deviceFilter, bool isVideoDevice)
			: base(new List<Source>(addFromGraph(graphBuilder, deviceFilter, isVideoDevice)))
		{ }
		
		~SourceCollection()
		{
			Dispose(false);
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (_IsDisposed)
				return;

			foreach (var item in this)
			{
				item.Dispose();
			}

			_IsDisposed = true;
		}

		/// <summary> Populate the collection from a filter graph. </summary>
		protected static IEnumerable<Source> addFromGraph(ICaptureGraphBuilder2 graphBuilder, IBaseFilter deviceFilter, bool isVideoDevice)
		{
			Trace.Assert(graphBuilder != null);

			var result = new List<Source>();

			IEnumerable<object> crossbars = findCrossbars(graphBuilder, deviceFilter);
			foreach (IAMCrossbar crossbar in crossbars)
			{
				result.AddRange(findCrossbarSources(graphBuilder, crossbar, isVideoDevice));
			}

			if (!isVideoDevice)
			{
				if (result.Count == 0)
				{
					result.AddRange(findAudioSources(graphBuilder, deviceFilter));
				}
			}

			return result;
		}

		/// <summary>
		///	 Retrieve a list of crossbar filters in the graph.
		///  Most hardware devices should have a maximum of 2 crossbars, 
		///  one for video and another for audio.
		/// </summary>
		protected static IEnumerable<IAMCrossbar> findCrossbars(ICaptureGraphBuilder2 graphBuilder, IBaseFilter deviceFilter)
		{
			//var crossbars = new List<IAMCrossbar>();

			DsGuid category = DsGuid.FromGuid(FindDirection.UpstreamOnly);
			DsGuid type = new DsGuid();
			Guid riid = typeof(IAMCrossbar).GUID;

			object comObj = null;
			object comObjNext = null;

			// Find the first interface, look upstream from the selected device
			int hr = graphBuilder.FindInterface(category, type, deviceFilter, riid, out comObj);
			while ((hr == 0) && (comObj != null))
			{
				// If found, add to the list
				if (comObj is IAMCrossbar)
				{
					yield return comObj as IAMCrossbar;
					//crossbars.Add(comObj as IAMCrossbar);

					// Find the second interface, look upstream from the next found crossbar
					hr = graphBuilder.FindInterface(category, type, comObj as IBaseFilter, riid, out comObjNext);
					comObj = comObjNext;
				}
				else
					comObj = null;
			}

			//return (crossbars);
		}

		/// <summary>
		///  Populate the internal InnerList with sources/physical connectors
		///  found on the crossbars. Each instance of this class is limited
		///  to video only or audio only sources ( specified by the isVideoDevice
		///  parameter on the constructor) so we check each source before adding
		///  it to the list.
		/// </summary>
		protected static IEnumerable<CrossbarSource> findCrossbarSources(ICaptureGraphBuilder2 graphBuilder, IAMCrossbar crossbar, bool isVideoDevice)
		{
			//ArrayList sources = new ArrayList();
			var sources = new List<CrossbarSource>();
			int numOutPins;
			int numInPins;
			int hr = crossbar.get_PinCounts(out numOutPins, out numInPins);
			Marshal.ThrowExceptionForHR(hr);

			// We loop through every combination of output and input pin
			// to see which combinations match.

			// Loop through output pins
			for (int cOut = 0; cOut < numOutPins; cOut++)
			{
				// Loop through input pins
				for (int cIn = 0; cIn < numInPins; cIn++)
				{
					// Can this combination be routed?
					hr = crossbar.CanRoute(cOut, cIn);
					if (hr == 0)
					{
						// Yes, this can be routed
						int relatedPin;
						PhysicalConnectorType connectorType;
						hr = crossbar.get_CrossbarPinInfo(true, cIn, out relatedPin, out connectorType);
						if (hr < 0)
							Marshal.ThrowExceptionForHR(hr);

						// Is this the correct type?, If so add to the InnerList
						CrossbarSource source = new CrossbarSource(crossbar, cOut, cIn, connectorType);
						if (connectorType < PhysicalConnectorType.Audio_Tuner)
							if (isVideoDevice)
								sources.Add(source);
							else
								if (!isVideoDevice)
									sources.Add(source);
					}
				}
			}

			// Some silly drivers (*cough* Nvidia *cough*) add crossbars
			// with no real choices. Every input can only be routed to
			// one output. Loop through every Source and see if there
			// at least one other Source with the same output pin.
			int refIndex = 0;
			while (refIndex < sources.Count)
			{
				bool found = false;
				CrossbarSource refSource = (CrossbarSource)sources[refIndex];
				for (int c = 0; c < sources.Count; c++)
				{
					CrossbarSource s = (CrossbarSource)sources[c];
					if ((refSource.OutputPin == s.OutputPin) && (refIndex != c))
					{
						found = true;
						break;
					}
				}
				if (found)
					refIndex++;
				else
					sources.RemoveAt(refIndex);
			}

			return sources;
		}

		protected static IEnumerable<Source> findAudioSources(ICaptureGraphBuilder2 graphBuilder, IBaseFilter deviceFilter)
		{
			var sources = new List<Source>();
			IAMAudioInputMixer audioInputMixer = deviceFilter as IAMAudioInputMixer;
			if (audioInputMixer != null)
			{
				// Get a pin enumerator off the filter
				IEnumPins pinEnum;
				int hr = deviceFilter.EnumPins(out pinEnum);
				pinEnum.Reset();
				
				if ((hr == 0) && (pinEnum != null))
				{
					// Loop through each pin
					IPin[] pins = new IPin[1];
					IntPtr f = IntPtr.Zero;
					do
					{
						// Get the next pin
						hr = pinEnum.Next(1, pins, f);
						if ((hr == 0) && (pins[0] != null))
						{
							// Is this an input pin?
							PinDirection dir = PinDirection.Output;
							hr = pins[0].QueryDirection(out dir);
							if ((hr == 0) && (dir == (PinDirection.Input)))
							{
								// Add the input pin to the sources list
								AudioSource source = new AudioSource(pins[0]);
								sources.Add(source);
							}
							pins[0] = null;
						}
					}
					while (hr == 0);

					Marshal.ReleaseComObject(pinEnum); pinEnum = null;
				}
			}

			// If there is only one source, don't return it
			// because there is nothing for the user to choose.
			// (Hopefully that single source is already enabled).
			if (sources.Count == 1)
				sources.Clear();

			return sources;
		}
	}
}
