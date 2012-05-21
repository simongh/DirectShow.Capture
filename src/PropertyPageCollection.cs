using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	public class PropertyPageCollection : System.Collections.ObjectModel.ReadOnlyCollection<PropertyPage>, IDisposable
	{
		private bool _IsDisposed;

		internal PropertyPageCollection()
			:base(new List<PropertyPage>())
		{}

				/// <summary> Initialize collection with property pages from existing graph. </summary>
		internal PropertyPageCollection(ICaptureGraphBuilder2 graphBuilder,
										IBaseFilter videoDeviceFilter,
										IBaseFilter audioDeviceFilter,
										IBaseFilter videoCompressorFilter,
										IBaseFilter audioCompressorFilter,
										SourceCollection videoSources,
										SourceCollection audioSources)
			: base(new List<PropertyPage>(addFromGraph(graphBuilder, videoDeviceFilter, audioDeviceFilter, videoCompressorFilter, audioCompressorFilter, videoSources, audioSources)))
		{ }

		~PropertyPageCollection()
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
			foreach (var item in Items)
			{
				item.Dispose();
			}
		}

		/// <summary> Populate the collection by looking for commonly implemented property pages. </summary>
		protected static IEnumerable<PropertyPage> addFromGraph(ICaptureGraphBuilder2 graphBuilder,
																IBaseFilter videoDeviceFilter, 
																IBaseFilter audioDeviceFilter,
																IBaseFilter videoCompressorFilter, 
																IBaseFilter audioCompressorFilter,
																SourceCollection videoSources, 
																SourceCollection audioSources)
		{
			var result = new List<PropertyPage>();

			object filter = null;

			Trace.Assert(graphBuilder != null);

			// 1. the video capture filter
			addIfSupported(result, videoDeviceFilter, "Video Capture Device");

			// 2. the video capture pin
			DsGuid cat = DsGuid.FromGuid(PinCategory.Capture);
			DsGuid med = DsGuid.FromGuid(MediaType.Interleaved);
			Guid iid = typeof(IAMStreamConfig).GUID;
			int hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
			//int hr = graphBuilder.FindInterface(ref cat, ref med, videoDeviceFilter, ref iid, out filter);
			if (hr != 0)
			{
				med = DsGuid.FromGuid(MediaType.Video);
				hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
				if (hr != 0)
					filter = null;
			}
			addIfSupported(result, filter, "Video Capture Pin");

			// 3. the video preview pin
			cat = DsGuid.FromGuid(PinCategory.Preview);
			med = DsGuid.FromGuid(MediaType.Interleaved);
			iid = typeof(IAMStreamConfig).GUID;
			hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
			if (hr != 0)
			{
				med = DsGuid.FromGuid(MediaType.Video);
				hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
				if (hr != 0)
					filter = null;
			}
			addIfSupported(result, filter, "Video Preview Pin");

			// 4. the video crossbar(s)
			var crossbars = new List<IAMCrossbar>();
			int num = 1;
			for (int c = 0; c < videoSources.Count; c++)
			{
				CrossbarSource s = videoSources[c] as CrossbarSource;
				if (s != null)
				{
					if (crossbars.IndexOf(s.Crossbar) < 0)
					{
						crossbars.Add(s.Crossbar);
						if (addIfSupported(result, s.Crossbar, "Video Crossbar " + (num == 1 ? "" : num.ToString())))
							num++;
					}
				}
			}
			crossbars.Clear();

			// 5. the video compressor
			addIfSupported(result, videoCompressorFilter, "Video Compressor");

			// 6. the video TV tuner
			cat = DsGuid.FromGuid(PinCategory.Capture);
			med = DsGuid.FromGuid(MediaType.Interleaved);
			iid = typeof(IAMTVTuner).GUID;
			hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
			if (hr != 0)
			{
				med = DsGuid.FromGuid(MediaType.Video);
				hr = graphBuilder.FindInterface(cat, med, videoDeviceFilter, iid, out filter);
				if (hr != 0)
					filter = null;
			}
			addIfSupported(result, filter, "TV Tuner");

			// 7. the video compressor (VFW)
			IAMVfwCompressDialogs compressDialog = videoCompressorFilter as IAMVfwCompressDialogs;
			if (compressDialog != null)
			{
				VfwCompressorPropertyPage page = new VfwCompressorPropertyPage("Video Compressor", compressDialog);
				result.Add(page);
			}

			// 8. the audio capture filter
			addIfSupported(result, audioDeviceFilter, "Audio Capture Device");

			// 9. the audio capture pin
			cat = DsGuid.FromGuid(PinCategory.Capture);
			med = DsGuid.FromGuid(MediaType.Audio);
			iid = typeof(IAMStreamConfig).GUID;
			hr = graphBuilder.FindInterface(cat, med, audioDeviceFilter, iid, out filter);
			if (hr != 0)
			{
				filter = null;
			}
			addIfSupported(result, filter, "Audio Capture Pin");

			// 9. the audio preview pin
			cat = DsGuid.FromGuid(PinCategory.Preview);
			med = DsGuid.FromGuid(MediaType.Audio);
			iid = typeof(IAMStreamConfig).GUID;
			hr = graphBuilder.FindInterface(cat, med, audioDeviceFilter, iid, out filter);
			if (hr != 0)
			{
				filter = null;
			}
			addIfSupported(result, filter, "Audio Preview Pin");

			// 10. the audio crossbar(s)
			num = 1;
			for (int c = 0; c < audioSources.Count; c++)
			{
				CrossbarSource s = audioSources[c] as CrossbarSource;
				if (s != null)
				{
					if (crossbars.IndexOf(s.Crossbar) < 0)
					{
						crossbars.Add(s.Crossbar);
						if (addIfSupported(result, s.Crossbar, "Audio Crossbar " + (num == 1 ? "" : num.ToString())))
							num++;
					}
				}
			}
			crossbars.Clear();

			// 11. the audio compressor
			addIfSupported(result, audioCompressorFilter, "Audio Compressor");

			return result;
		}

		/// <summary> 
		///  Returns the object as an ISpecificPropertyPage
		///  if the object supports the ISpecificPropertyPage
		///  interface and has at least one property page.
		/// </summary>
		protected static bool addIfSupported(ICollection<PropertyPage>list, object o, string name)
		{
			ISpecifyPropertyPages specifyPropertyPages = null;
			DsCAUUID cauuid = new DsCAUUID();
			bool wasAdded = false;

			// Determine if the object supports the interface
			// and has at least 1 property page
			try
			{
				specifyPropertyPages = o as ISpecifyPropertyPages;
				if (specifyPropertyPages != null)
				{
					int hr = specifyPropertyPages.GetPages(out cauuid);
					if ((hr != 0) || (cauuid.cElems <= 0))
						specifyPropertyPages = null;
				}
			}
			finally
			{
				if (cauuid.pElems != IntPtr.Zero)
					Marshal.FreeCoTaskMem(cauuid.pElems);
			}

			// Add the page to the internal collection
			if (specifyPropertyPages != null)
			{
				DirectShowPropertyPage p = new DirectShowPropertyPage(name, specifyPropertyPages);
				list.Add(p);
				wasAdded = true;
			}
			return wasAdded;
		}
	}
}
