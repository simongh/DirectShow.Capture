using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Use the Capture class to capture audio and video to AVI files.
	/// </summary>
	/// <remarks>
	///  This is the core class of the Capture Class Library. The following 
	///  sections introduce the Capture class and how to use this library.
	///  
	/// <br/><br/>
	/// <para><b>Basic Usage</b></para>
	/// 
	/// <para>
	///  The Capture class only requires a video device and/or audio device
	///  to begin capturing. The <see cref="Filters"/> class provides
	///  lists of the installed video and audio devices. 
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  // Remember to add a reference to DirectShow.Capture.dll
	///  using DirectShow.Capture
	///  ...
	///  Capture capture = new Capture( Filters.VideoInputDevices[0], 
	///                                 Filters.AudioInputDevices[0] );
	///  capture.Start();
	///  ...
	///  capture.Stop();
	/// </div></code>
	/// <para>
	///  This will capture video and audio using the first video and audio devices
	///  installed on the system. To capture video only, pass a null as the second
	///  parameter of the constructor.
	/// </para>
	/// <para> 
	///  The class is initialized to a valid temporary file in the Windows temp
	///  folder. To capture to a different file, set the 
	///  <see cref="Capture.Filename"/> property before you begin
	///  capturing. Remember to add DirectShow.Capture.dll to 
	///  your project references.
	/// </para>
	///
	/// <br/>
	/// <para><b>Setting Common Properties</b></para>
	/// 
	/// <para>
	///  The example below shows how to change video and audio settings. 
	///  Properties such as <see cref="Capture.FrameRate"/> and 
	///  <see cref="AudioSampleSize"/> allow you to programmatically adjust
	///  the capture. Use <see cref="Capture.VideoCaps"/> and 
	///  <see cref="Capture.AudioCaps"/> to determine valid values for these
	///  properties.
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  Capture capture = new Capture( Filters.VideoInputDevices[0], 
	///                                 Filters.AudioInputDevices[1] );
	///  capture.VideoCompressor = Filters.VideoCompressors[0];
	///  capture.AudioCompressor = Filters.AudioCompressors[0];
	///  capture.FrameRate = 29.997;
	///  capture.FrameSize = new Size( 640, 480 );
	///  capture.AudioSamplingRate = 44100;
	///  capture.AudioSampleSize = 16;
	///  capture.Filename = "C:\MyVideo.avi";
	///  capture.Start();
	///  ...
	///  capture.Stop();
	/// </div></code>
	/// <para>
	///  The example above also shows the use of video and audio compressors. In most 
	///  cases you will want to use compressors. Uncompressed video can easily
	///  consume over a 1GB of disk space per minute. Whenever possible, set 
	///  the <see cref="Capture.VideoCompressor"/> and <see cref="Capture.AudioCompressor"/>
	///  properties as early as possible. Changing them requires the internal filter
	///  graph to be rebuilt which often causes most of the other properties to
	///  be reset to default values.
	/// </para>
	///
	/// <br/>
	/// <para><b>Listing Devices</b></para>
	/// 
	/// <para>
	///  Use the <see cref="Filters.VideoInputDevices"/> collection to list
	///  video capture devices installed on the system.
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  foreach ( Filter f in Filters.VideoInputDevices )
	///  {
	///		Debug.WriteLine( f.Name );
	///  }
	/// </div></code>
	/// The <see cref="Filters"/> class also provides collections for audio 
	/// capture devices, video compressors and audio compressors.
	///
	/// <br/>
	/// <para><b>Preview</b></para>
	/// 
	/// <para>
	///  Video preview is controled with the <see cref="Capture.PreviewWindow"/>
	///  property. Setting this property to a visible control will immediately 
	///  begin preview. Set to null to stop the preview. 
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  // Enable preview
	///  capture.PreviewWindow = myPanel;
	///  // Disable preview
	///  capture.PreviewWindow = null;
	/// </div></code>
	/// <para>
	///  The control used must have a window handle (HWND), good controls to 
	///  use are the Panel or the form itself.
	/// </para>
	/// <para>
	///  Retrieving or changing video/audio settings such as FrameRate, 
	///  FrameSize, AudioSamplingRate, and AudioSampleSize will cause
	///  the preview window to flash. This is beacuse the preview must be
	///  temporarily stopped. Disable the preview if you need to access
	///  several properties at the same time.
	/// </para>
	/// 
	/// <br/>
	/// <para><b>Property Pages</b></para>
	///  
	/// <para>
	///  Property pages exposed by the devices and compressors are
	///  available through the <see cref="Capture.PropertyPages"/> 
	///  collection.
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  // Display the first property page
	///  capture.PropertyPages[0].Show();
	/// </div></code>
	/// <para>
	///  The property pages will often expose more settings than
	///  the Capture class does directly. Some examples are brightness, 
	///  color space, audio balance and bass boost. The disadvantage
	///  to using the property pages is the user's choices cannot be
	///  saved and later restored. The exception to this is the video
	///  and audio compressor property pages. Most compressors support
	///  the saving and restoring state, see the 
	///  <see cref="PropertyPage.State"/> property for more information.
	/// </para>
	/// <para>
	///  Changes made in the property page will be reflected 
	///  immediately in the Capture class properties (e.g. Capture.FrameSize).
	///  However, the reverse is not always true. A change made directly to
	///  FrameSize, for example, may not be reflected in the associated
	///  property page. Fortunately, the filter will use requested FrameSize
	///  even though the property page shows otherwise.
	/// </para>
	/// 
	/// <br/>
	/// <para><b>Saving and Restoring Settings</b></para>
	///  
	/// <para>
	///  To save the user's choice of devices and compressors, 
	///  save <see cref="Filter.MonikerString"/> and user it later
	///  to recreate the Filter object.
	/// </para>
	/// <para>
	///  To save a user's choices from a property page use
	///  <see cref="PropertyPage.State"/>. However, only the audio
	///  and video compressor property pages support this.
	/// </para>
	/// <para>
	///  The last items to save are the video and audio settings such
	///  as FrameSize and AudioSamplingRate. When restoring, remember
	///  to restore these properties after setting the video and audio
	///  compressors.
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  // Disable preview
	///  capture.PreviewWindow = null;
	///  
	///  // Save settings
	///  string videoDevice = capture.VideoDevice.MonikerString;
	///  string audioDevice = capture.AudioDevice.MonikerString;
	///  string videoCompressor = capture.VideoCompressor.MonikerString;
	///  string audioCompressor = capture.AudioCompressor.MonikerString;
	///  double frameRate = capture.FrameRate;
	///  Size frameSize = capture.FrameSize;
	///  short audioChannels = capture.AudioChannels;
	///  short audioSampleSize = capture.AudioSampleSize;
	///  int audioSamplingRate = capture.AudioSamplingRate;
	///  ArrayList pages = new ArrayList();
	///  foreach ( PropertyPage p in capture.PropertyPages )
	///  {
	///		if ( p.SupportsPersisting )
	///			pages.Add( p.State );
	///	}
	///			
	///  
	///  // Restore settings
	///  Capture capture = new Capture( new Filter( videoDevice), 
	///				new Filter( audioDevice) );
	///  capture.VideoCompressor = new Filter( videoCompressor );
	///  capture.AudioCompressor = new Filter( audioCompressor );
	///  capture.FrameRate = frameRate;
	///  capture.FrameSize = frameSize;
	///  capture.AudioChannels = audioChannels;
	///  capture.AudioSampleSize = audioSampleSize;
	///  capture.AudioSamplingRate = audioSamplingRate;
	///  foreach ( PropertyPage p in capture.PropertyPages )
	///  {
	///		if ( p.SupportsPersisting )
	///		{
	///			p.State = (byte[]) pages[0]
	///			pages.RemoveAt( 0 );
	///		}
	///	 }
	///  // Enable preview
	///  capture.PreviewWindow = myPanel;
	/// </div></code>
	///  
	/// <br/>
	/// <para><b>TV Tuner</b></para>
	/// 
	/// <para>
	///  To access the TV Tuner, use the <see cref="Capture.Tuner"/> property.
	///  If the device does not have a TV tuner, this property will be null.
	///  See <see cref="DirectShow.Capture.Tuner.Channel"/>, 
	///  <see cref="DirectShow.Capture.Tuner.InputType"/> and 
	///  <see cref="DirectShow.Capture.Tuner.SignalPresent"/> 
	///  for more information.
	/// </para>
	/// <code><div style="background-color:whitesmoke;">
	///  // Change to channel 5
	///  capture.Tuner.Channel = 5;
	/// </div></code>
	/// 
	/// <br/>
	/// <para><b>Troubleshooting</b></para>
	/// 
	/// <para>
	///  This class library uses COM Interop to access the full
	///  capabilities of DirectShow, so if there is another
	///  application that can successfully use a hardware device
	///  then it should be possible to modify this class library
	///  to use the device.
	/// </para>
	/// <para>
	///  Try the <b>AMCap</b> sample from the DirectShow SDK 
	///  (DX9\Samples\C++\DirectShow\Bin\AMCap.exe) or 
	///  <b>Virtual VCR</b> from http://www.DigTV.ws 
	/// </para>
	/// 
	/// <br/>
	/// <para><b>Credits</b></para>
	/// 
	/// <para>
	///  This class library would not be possible without the
	///  DShowNET project by NETMaster: 
	///  http://www.codeproject.com/useritems/directshownet.asp
	/// </para>
	/// <para>
	///  Documentation is generated by nDoc available at
	///  http://ndoc.sourceforge.net
	/// </para>
	/// 
	/// <br/>
	/// <para><b>Feedback</b></para>
	/// 
	///  Feel free to send comments and questions to me at
	///  mportobello@hotmail.com. If the the topic may be of interest
	///  to others, post your question on the www.codeproject.com
	///  page for DirectShow.Capture.
	/// </remarks>
	public class Capture : IDisposable
	{
		private bool _IsDisposed;
		private GraphState _graphState;
		private bool _isPreviewRendered;
		private bool _isCaptureRendered;
		private bool _wantPreviewRendered;
		private bool _wantCaptureRendered;

		private int _rotCookie;						// Cookie into the Running Object Table
		private Filter _videoCompressor;
		private Filter _audioCompressor;
		private string _filename = "";
		private System.Windows.Forms.Control _previewWindow;
		private VideoCapabilities _videoCaps;
		private AudioCapabilities _audioCaps;
		private SourceCollection _videoSources;
		private SourceCollection _audioSources;
		private PropertyPageCollection _propertyPages;

		private IGraphBuilder _graphBuilder;						// DShow Filter: Graph builder 
		private IMediaControl _mediaControl;
		private IVideoWindow _videoWindow;						// DShow Filter: Control preview window -> copy of graphBuilder
		private ICaptureGraphBuilder2 _captureGraphBuilder;	// DShow Filter: building graphs for capturing video
		private IAMStreamConfig _videoStreamConfig;
		private IAMStreamConfig _audioStreamConfig;
		private IBaseFilter _videoDeviceFilter;
		private IBaseFilter _videoCompressorFilter;
		private IBaseFilter _audioDeviceFilter;
		private IBaseFilter _audioCompressorFilter;
		private IBaseFilter _muxFilter;
		private IFileSinkFilter _fileWriterFilter;

		/// <summary> 
		/// Fired when a capture is completed (manually or automatically). 
		/// </summary>
		public event EventHandler CaptureComplete;

		/// <summary> 
		/// Possible states of the interal filter graph 
		/// </summary>
		protected enum GraphState
		{
			/// <summary>
			/// No filter graph at all
			/// </summary>
			Null,
			/// <summary>
			/// Filter graph created with device filters added
			/// </summary>
			Created,
			/// <summary>
			/// Filter complete built, ready to run (possibly previewing)
			/// </summary>
			Rendered,
			/// <summary>
			/// Filter is capturing
			/// </summary>
			Capturing
		}

		/// <summary>
		/// Gets whether the class currently capturing.
		/// </summary>
		public bool IsCapturing 
		{ 
			get { return _graphState == GraphState.Capturing; } 
		}

		/// <summary>
		/// Gets whether the class been cued to begin capturing.
		/// </summary>
		public bool IsCued 
		{ 
			get { return _isCaptureRendered && _graphState == GraphState.Rendered; } 
		}

		/// <summary>
		/// Gets whether the class currently stopped.
		/// </summary>
		public bool Stopped 
		{ 
			get { return _graphState != GraphState.Capturing; } 
		}

		/// <summary> 
		///  Name of file to capture to. Initially set to
		///  a valid temporary file.
		/// </summary>		
		/// <remarks>
		///  If the file does not exist, it will be created. If it does 
		///  exist, it will be overwritten. An overwritten file will 
		///  not be shortened if the captured data is smaller than the 
		///  original file. The file will be valid, it will just contain 
		///  extra, unused, data after the audio/video data. 
		/// 
		/// <para>
		///  A future version of this class will provide a method to copy 
		///  only the valid audio/video data to a new file. </para>
		/// 
		/// <para>
		///  This property cannot be changed while capturing or cued. </para>
		/// </remarks> 
		public string Filename
		{
			get { return _filename; }
			set
			{
				assertStopped();
				if (IsCued)
					throw new InvalidOperationException("The Filename cannot be changed once cued. Use Stop() before changing the filename.");
				
				_filename = value;
				
				if (_fileWriterFilter != null)
				{
					string s;
					AMMediaType mt = new AMMediaType();
					Marshal.ThrowExceptionForHR(_fileWriterFilter.GetCurFile(out s, mt));
					
					if (mt.formatSize > 0)
						Marshal.FreeCoTaskMem(mt.formatPtr);

					Marshal.ThrowExceptionForHR(_fileWriterFilter.SetFileName(_filename, mt));
				}
			}
		}

		/// <summary>
		/// Gets the control that will host the preview window. 
		/// </summary>
		/// <remarks>
		///  Setting this property will begin video preview
		///  immediately. Set this property after setting all
		///  other properties to avoid unnecessary changes
		///  to the internal filter graph (some properties like
		///  FrameSize require the internal filter graph to be 
		///  stopped and disconnected before the property
		///  can be retrieved or set).
		///  
		/// <para>
		///  To stop video preview, set this property to null. </para>
		/// </remarks>
		public System.Windows.Forms.Control PreviewWindow
		{
			get { return _previewWindow; }
			set
			{
				assertStopped();
				derenderGraph();
				_previewWindow = value;
				_wantPreviewRendered = ((_previewWindow != null) && (VideoDevice != null));
				renderGraph();
				startPreviewIfNeeded();
			}
		}

		/// <summary>
		/// Gets the capabilities of the video device.
		/// </summary>
		/// <remarks>
		///  It may be required to cue the capture (see <see cref="Cue"/>) 
		///  before all capabilities are correctly reported. If you 
		///  have such a device, the developer would be interested to
		///  hear from you.
		/// 
		/// <para>
		///  The information contained in this property is retrieved and
		///  cached the first time this property is accessed. Future
		///  calls to this property use the cached results. This was done 
		///  for performance. </para>
		///  
		/// <para>
		///  However, this means <b>you may get different results depending 
		///  on when you access this property first</b>. If you are experiencing 
		///  problems, try accessing the property immediately after creating 
		///  the Capture class or immediately after setting the video and 
		///  audio compressors. Also, inform the developer. </para>
		/// </remarks>
		public VideoCapabilities VideoCaps
		{
			get
			{
				if (_videoCaps == null && _videoStreamConfig != null)
				{
					try
					{
						_videoCaps = new VideoCapabilities(_videoStreamConfig);
					}
					catch (Exception ex)
					{
						Debug.WriteLine("VideoCaps: unable to create videoCaps." + ex.ToString());
					}
				}
				
				return _videoCaps;
			}
		}

		/// <summary>
		/// Gets the capabilities of the audio device.
		/// </summary>
		/// <remarks>
		///  It may be required to cue the capture (see <see cref="Cue"/>) 
		///  before all capabilities are correctly reported. If you 
		///  have such a device, the developer would be interested to
		///  hear from you.
		/// 
		/// <para>
		///  The information contained in this property is retrieved and
		///  cached the first time this property is accessed. Future
		///  calls to this property use the cached results. This was done 
		///  for performance. </para>
		///  
		/// <para>
		///  However, this means <b>you may get different results depending 
		///  on when you access this property first</b>. If you are experiencing 
		///  problems, try accessing the property immediately after creating 
		///  the Capture class or immediately after setting the video and 
		///  audio compressors. Also, inform the developer. </para>
		/// </remarks>
		public AudioCapabilities AudioCaps
		{
			get
			{
				if (_audioCaps == null && _audioStreamConfig != null)
				{
					try
					{
						_audioCaps = new AudioCapabilities(_audioStreamConfig);
					}
					catch (Exception ex) 
					{ 
						Debug.WriteLine("AudioCaps: unable to create audioCaps." + ex.ToString()); 
					}
				}
				
				return _audioCaps;
			}
		}

		/// <summary> 
		///  The video capture device filter. Read-only. To use a different 
		///  device, dispose of the current Capture instance and create a new 
		///  instance with the desired device. 
		/// </summary>
		public Filter VideoDevice
		{
			get;
			private set;
		}

		/// <summary> 
		///  The audio capture device filter. Read-only. To use a different 
		///  device, dispose of the current Capture instance and create a new 
		///  instance with the desired device. 
		/// </summary>
		public Filter AudioDevice
		{
			get;
			private set;
		}

		/// <summary> 
		///  The video compression filter. When this property is changed 
		///  the internal filter graph is rebuilt. This means that some properties
		///  will be reset. Set this property as early as possible to avoid losing 
		///  changes. This property cannot be changed while capturing.
		/// </summary>
		public Filter VideoCompressor
		{
			get { return _videoCompressor; }
			set
			{
				assertStopped();
				destroyGraph();
				_videoCompressor = value;
				renderGraph();
				startPreviewIfNeeded();
			}
		}

		/// <summary> 
		///  The audio compression filter. 
		/// </summary>
		/// <remarks>
		///  When this property is changed 
		///  the internal filter graph is rebuilt. This means that some properties
		///  will be reset. Set this property as early as possible to avoid losing 
		///  changes. This property cannot be changed while capturing.
		/// </remarks>
		public Filter AudioCompressor
		{
			get { return _audioCompressor; }
			set
			{
				assertStopped();
				destroyGraph();
				_audioCompressor = value;
				renderGraph();
				startPreviewIfNeeded();
			}
		}

		/// <summary> 
		///  The current video source. Use Capture.VideoSources to 
		///  list available sources. Set to null to disable all 
		///  sources (mute).
		/// </summary>
		public Source VideoSource
		{
			get { return VideoSources.CurrentSource; }
			set { VideoSources.CurrentSource = value; }
		}

		/// <summary> 
		///  The current audio source. Use Capture.AudioSources to 
		///  list available sources. Set to null to disable all 
		///  sources (mute).
		/// </summary>
		public Source AudioSource
		{
			get { return AudioSources.CurrentSource; }
			set { AudioSources.CurrentSource = value; }
		}

		/// <summary> 
		///  Collection of available video sources/physical connectors 
		///  on the current video device. 
		/// </summary>
		/// <remarks>
		///  In most cases, if the device has only one source, 
		///  this collection will be empty. 
		/// 
		/// <para>
		///  The information contained in this property is retrieved and
		///  cached the first time this property is accessed. Future
		///  calls to this property use the cached results. This was done 
		///  for performance. </para>
		///  
		/// <para>
		///  However, this means <b>you may get different results depending 
		///  on when you access this property first</b>. If you are experiencing 
		///  problems, try accessing the property immediately after creating 
		///  the Capture class or immediately after setting the video and 
		///  audio compressors. Also, inform the developer. </para>
		/// </remarks>
		public SourceCollection VideoSources
		{
			get
			{
				if (_videoSources == null)
				{
					try
					{
						if (VideoDevice != null)
							_videoSources = new SourceCollection(_captureGraphBuilder, _videoDeviceFilter, true);
						else
							_videoSources = new SourceCollection();
					}
					catch (Exception ex) 
					{ 
						Debug.WriteLine("VideoSources: unable to create VideoSources." + ex.ToString()); 
					}
				}
				return _videoSources;
			}
		}

		/// <summary> 
		///  Collection of available audio sources/physical connectors 
		///  on the current audio device. 
		/// </summary>
		/// <remarks>
		///  In most cases, if the device has only one source, 
		///  this collection will be empty. For audio
		///  there are 2 different methods for enumerating audio sources
		///  an audio crossbar (usually TV tuners?) or an audio mixer 
		///  (usually sound cards?). This class will first look for an 
		///  audio crossbar. If no sources or only one source is available
		///  on the crossbar, this class will then look for an audio mixer.
		///  This class does not support both methods.
		/// 
		/// <para>
		///  The information contained in this property is retrieved and
		///  cached the first time this property is accessed. Future
		///  calls to this property use the cached results. This was done 
		///  for performance. </para>
		///  
		/// <para>
		///  However, this means <b>you may get different results depending 
		///  on when you access this property first</b>. If you are experiencing 
		///  problems, try accessing the property immediately after creating 
		///  the Capture class or immediately after setting the video and 
		///  audio compressors. Also, inform the developer. </para>
		///  </remarks>
		public SourceCollection AudioSources
		{
			get
			{
				if (_audioSources == null)
				{
					try
					{
						if (AudioDevice != null)
							_audioSources = new SourceCollection(_captureGraphBuilder, _audioDeviceFilter, false);
						else
							_audioSources = new SourceCollection();
					}
					catch (Exception ex) 
					{ 
						Debug.WriteLine("AudioSources: unable to create AudioSources." + ex.ToString()); 
					}
				}
				return _audioSources;
			}
		}

		/// <summary>
		///  Available property pages. 
		/// </summary>
		/// <remarks>
		///  These are property pages exposed by the DirectShow filters. 
		///  These property pages allow users modify settings on the 
		///  filters directly. 
		/// 
		/// <para>
		///  The information contained in this property is retrieved and
		///  cached the first time this property is accessed. Future
		///  calls to this property use the cached results. This was done 
		///  for performance. </para>
		///  
		/// <para>
		///  However, this means <b>you may get different results depending 
		///  on when you access this property first</b>. If you are experiencing 
		///  problems, try accessing the property immediately after creating 
		///  the Capture class or immediately after setting the video and 
		///  audio compressors. Also, inform the developer. </para>
		/// </remarks>
		public PropertyPageCollection PropertyPages
		{
			get
			{
				if (_propertyPages == null)
				{
					try
					{
						_propertyPages = new PropertyPageCollection(
							_captureGraphBuilder,
							_videoDeviceFilter, 
							_audioDeviceFilter,
							_videoCompressorFilter, 
							_audioCompressorFilter,
							VideoSources, 
							AudioSources);
					}
					catch (Exception ex) 
					{ 
						Debug.WriteLine("PropertyPages: unable to get property pages." + ex.ToString()); 
					}

				}
				return _propertyPages;
			}
		}

		/// <summary>
		///  The TV Tuner or null if the current video device 
		///  does not have a TV Tuner.
		/// </summary>
		public Tuner Tuner
		{
			get;
			private set;
		}

		/// <summary>
		///  Gets or sets the frame rate used to capture video.
		/// </summary>
		/// <remarks>
		///  Common frame rates: 24 fps for film, 25 for PAL, 29.997
		///  for NTSC. Not all NTSC capture cards can capture at 
		///  exactly 29.997 fps. Not all frame rates are supported. 
		///  When changing the frame rate, the closest supported 
		///  frame rate will be used. 
		///  
		/// <para>
		///  Not all devices support getting/setting this property.
		///  If this property is not supported, accessing it will
		///  throw and exception. </para>
		///  
		/// <para>
		///  This property cannot be changed while capturing. Changing 
		///  this property while preview is enabled will cause some 
		///  fickering while the internal filter graph is partially
		///  rebuilt. Changing this property while cued will cancel the
		///  cue. Call Cue() again to re-cue the capture. </para>
		/// </remarks>
		public double FrameRate
		{
			get
			{
				long avgTimePerFrame = (long)getStreamConfigSetting(_videoStreamConfig, "AvgTimePerFrame");
				return (double)10000000 / avgTimePerFrame;
			}
			set
			{
				long avgTimePerFrame = (long)(10000000 / value);
				setStreamConfigSetting(_videoStreamConfig, "AvgTimePerFrame", avgTimePerFrame);
			}
		}

		/// <summary>
		///  Gets or sets the frame size used to capture video.
		/// </summary>
		/// <remarks>
		///  To change the frame size, assign a new Size object 
		///  to this property <code>capture.Size = new Size( w, h );</code>
		///  rather than modifying the size in place 
		///  (capture.Size.Width = w;). Not all frame
		///  rates are supported.
		///  
		/// <para>
		///  Not all devices support getting/setting this property.
		///  If this property is not supported, accessing it will
		///  throw and exception. </para>
		/// 
		/// <para> 
		///  This property cannot be changed while capturing. Changing 
		///  this property while preview is enabled will cause some 
		///  fickering while the internal filter graph is partially
		///  rebuilt. Changing this property while cued will cancel the
		///  cue. Call Cue() again to re-cue the capture. </para>
		/// </remarks>
		public Size FrameSize
		{
			get
			{
				BitmapInfoHeader bmiHeader;
				bmiHeader = (BitmapInfoHeader)getStreamConfigSetting(_videoStreamConfig, "BmiHeader");
				Size size = new Size(bmiHeader.Width, bmiHeader.Height);
				return size;
			}
			set
			{
				BitmapInfoHeader bmiHeader;
				bmiHeader = (BitmapInfoHeader)getStreamConfigSetting(_videoStreamConfig, "BmiHeader");
				bmiHeader.Width = value.Width;
				bmiHeader.Height = value.Height;
				setStreamConfigSetting(_videoStreamConfig, "BmiHeader", bmiHeader);
			}
		}

		/// <summary>
		///  Get or set the number of channels in the waveform-audio data. 
		/// </summary>
		/// <remarks>
		///  Monaural data uses one channel and stereo data uses two channels. 
		///  
		/// <para>
		///  Not all devices support getting/setting this property.
		///  If this property is not supported, accessing it will
		///  throw and exception. </para>
		///  
		/// <para>
		///  This property cannot be changed while capturing. Changing 
		///  this property while preview is enabled will cause some 
		///  fickering while the internal filter graph is partially
		///  rebuilt. Changing this property while cued will cancel the
		///  cue. Call Cue() again to re-cue the capture. </para>
		/// </remarks>
		public short AudioChannels
		{
			get { return (short)getStreamConfigSetting(_audioStreamConfig, "nChannels"); }
			set { setStreamConfigSetting(_audioStreamConfig, "nChannels", value); }
		}

		/// <summary>
		///  Get or set the number of audio samples taken per second.
		/// </summary>
		/// <remarks>
		///  Common sampling rates are 8.0 kHz, 11.025 kHz, 22.05 kHz, and 
		///  44.1 kHz. Not all sampling rates are supported.
		///  
		/// <para>
		///  Not all devices support getting/setting this property.
		///  If this property is not supported, accessing it will
		///  throw and exception. </para>
		///  
		/// <para>
		///  This property cannot be changed while capturing. Changing 
		///  this property while preview is enabled will cause some 
		///  fickering while the internal filter graph is partially
		///  rebuilt. Changing this property while cued will cancel the
		///  cue. Call Cue() again to re-cue the capture. </para>
		/// </remarks>
		public int AudioSamplingRate
		{
			get { return (int)getStreamConfigSetting(_audioStreamConfig, "nSamplesPerSec"); }
			set { setStreamConfigSetting(_audioStreamConfig, "nSamplesPerSec", value); }
		}

		/// <summary>
		///  Get or set the number of bits recorded per sample. 
		/// </summary>
		/// <remarks>
		///  Common sample sizes are 8 bit and 16 bit. Not all
		///  samples sizes are supported.
		///  
		/// <para>
		///  Not all devices support getting/setting this property.
		///  If this property is not supported, accessing it will
		///  throw and exception. </para>
		///  
		/// <para>
		///  This property cannot be changed while capturing. Changing 
		///  this property while preview is enabled will cause some 
		///  fickering while the internal filter graph is partially
		///  rebuilt. Changing this property while cued will cancel the
		///  cue. Call Cue() again to re-cue the capture. </para>
		/// </remarks>
		public short AudioSampleSize
		{
			get { return (short)getStreamConfigSetting(_audioStreamConfig, "wBitsPerSample"); }
			set { setStreamConfigSetting(_audioStreamConfig, "wBitsPerSample", value); }
		}

		/// <summary> 
		///  Create a new Capture object. 
		///  videoDevice and audioDevice can be null if you do not 
		///  wish to capture both audio and video. However at least
		///  one must be a valid device. Use the <see cref="Filters"/> 
		///  class to list available devices.
		///  </summary>
		public Capture(Filter videoDevice, Filter audioDevice)
		{
			if (videoDevice == null && audioDevice == null)
				throw new ArgumentException("The videoDevice and/or the audioDevice parameter must be set to a valid Filter.\n");
			VideoDevice = videoDevice;
			AudioDevice = audioDevice;
			Filename = getTempFilename();
			createGraph();
		}

		~Capture()
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

			_wantPreviewRendered = false;
			_wantCaptureRendered = false;
			CaptureComplete = null;

			try { 
				destroyGraph(); 
			}
			catch 
			{ }

			if (_videoSources != null)
				_videoSources.Dispose(); 
			_videoSources = null;
			if (_audioSources != null)
				_audioSources.Dispose(); 
			_audioSources = null;

			_IsDisposed = true;
		}

		/// <summary>
		///  Prepare for capturing. Use this method when capturing 
		///  must begin as quickly as possible. 
		/// </summary>
		/// <remarks>
		///  This will create/overwrite a zero byte file with 
		///  the name set in the Filename property. 
		///  
		/// <para>
		///  This will disable preview. Preview will resume
		///  once capture begins. This problem can be fixed
		///  if someone is willing to make the change. </para>
		///  
		/// <para>
		///  This method is optional. If Cue() is not called, 
		///  Start() will call it before capturing. This method
		///  cannot be called while capturing. </para>
		/// </remarks>
		public void Cue()
		{
			assertStopped();

			// We want the capture stream rendered
			_wantCaptureRendered = true;

			// Re-render the graph (if necessary)
			renderGraph();

			// Pause the graph
			Marshal.ThrowExceptionForHR(_mediaControl.Pause());
		}

		/// <summary> 
		/// Begin capturing. 
		/// </summary>
		public void Start()
		{
			assertStopped();

			// We want the capture stream rendered
			_wantCaptureRendered = true;

			// Re-render the graph (if necessary)
			renderGraph();

			// Start the filter graph: begin capturing
			Marshal.ThrowExceptionForHR(_mediaControl.Run());

			// Update the state
			_graphState = GraphState.Capturing;
		}

		/// <summary> 
		///  Stop the current capture capture. If there is no
		///  current capture, this method will succeed.
		/// </summary>
		public void Stop()
		{
			_wantCaptureRendered = false;

			// Stop the graph if it is running
			// If we have a preview running we should only stop the
			// capture stream. However, if we have a preview stream
			// we need to re-render the graph anyways because we 
			// need to get rid of the capture stream. To re-render
			// we need to stop the entire graph
			if (_mediaControl != null)
			{
				_mediaControl.Stop();
			}

			// Update the state
			if (_graphState == GraphState.Capturing)
			{
				_graphState = GraphState.Rendered;
				OnCaptureComplete();
			}

			// So we destroy the capture stream IF 
			// we need a preview stream. If we don't
			// this will leave the graph as it is.
			try
			{ 
				renderGraph(); 
			}
			catch 
			{ }
			try 
			{ 
				startPreviewIfNeeded(); 
			}
			catch 
			{ }
		}

		protected virtual void OnCaptureComplete()
		{
			if (CaptureComplete != null)
				CaptureComplete(this, new EventArgs());
		}

		/// <summary> 
		///  Create a new filter graph and add filters (devices, compressors, 
		///  misc), but leave the filters unconnected. Call renderGraph()
		///  to connect the filters.
		/// </summary>
		private void createGraph()
		{
			DsGuid cat;
			DsGuid med;

			// Ensure required properties are set
			if (VideoDevice == null && AudioDevice == null)
				throw new ArgumentException("The video and/or audio device have not been set. Please set one or both to valid capture devices.\n");

			// Skip if we are already created
			if ((int)_graphState < (int)GraphState.Created)
			{
				// Garbage collect, ensure that previous filters are released
				GC.Collect();

				// Make a new filter graph
				//_graphBuilder = (IGraphBuilder)Activator.CreateInstance(Type.GetTypeFromCLSID(Clsid.FilterGraph, true));
				_graphBuilder = (IGraphBuilder)new FilterGraph();

				// Get the Capture Graph Builder
				//Guid clsid = Clsid.CaptureGraphBuilder2;
				//Guid riid = typeof(ICaptureGraphBuilder2).GUID;
				//_captureGraphBuilder = (ICaptureGraphBuilder2)DsBugWO.CreateDsInstance(ref clsid, ref riid);
				_captureGraphBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();

				// Link the CaptureGraphBuilder to the filter graph
				Marshal.ThrowExceptionForHR(_captureGraphBuilder.SetFiltergraph(_graphBuilder));

				// Add the graph to the Running Object Table so it can be
				// viewed with GraphEdit
#if DEBUG
				//DsROT.AddGraphToRot(_graphBuilder, out _rotCookie);
#endif

				// Get the video device and add it to the filter graph
				if (VideoDevice != null)
				{
					_videoDeviceFilter = (IBaseFilter)Marshal.BindToMoniker(VideoDevice.MonikerString);
					Marshal.ThrowExceptionForHR(_graphBuilder.AddFilter(_videoDeviceFilter, "Video Capture Device"));
				}

				// Get the audio device and add it to the filter graph
				if (AudioDevice != null)
				{
					_audioDeviceFilter = (IBaseFilter)Marshal.BindToMoniker(AudioDevice.MonikerString);
					Marshal.ThrowExceptionForHR(_graphBuilder.AddFilter(_audioDeviceFilter, "Audio Capture Device"));
				}

				// Get the video compressor and add it to the filter graph
				if (VideoCompressor != null)
				{
					_videoCompressorFilter = (IBaseFilter)Marshal.BindToMoniker(VideoCompressor.MonikerString);
					Marshal.ThrowExceptionForHR(_graphBuilder.AddFilter(_videoCompressorFilter, "Video Compressor"));
				}

				// Get the audio compressor and add it to the filter graph
				if (AudioCompressor != null)
				{
					_audioCompressorFilter = (IBaseFilter)Marshal.BindToMoniker(AudioCompressor.MonikerString);
					Marshal.ThrowExceptionForHR(_graphBuilder.AddFilter(_audioCompressorFilter, "Audio Compressor"));
				}

				// Retrieve the stream control interface for the video device
				// FindInterface will also add any required filters
				// (WDM devices in particular may need additional
				// upstream filters to function).

				// Try looking for an interleaved media type
				object o;
				cat = DsGuid.FromGuid(PinCategory.Capture);
				med = DsGuid.FromGuid(MediaType.Interleaved);
				Guid iid = typeof(IAMStreamConfig).GUID;
				int hr = _captureGraphBuilder.FindInterface(cat, med, _videoDeviceFilter, iid, out o);
				//int hr = _captureGraphBuilder.FindInterface(ref cat, ref med, _videoDeviceFilter, ref iid, out o);

				if (hr != 0)
				{
					// If not found, try looking for a video media type
					med = MediaType.Video;
					hr = _captureGraphBuilder.FindInterface(cat, med, _videoDeviceFilter, iid, out o);
					//hr = _captureGraphBuilder.FindInterface(ref cat, ref med, _videoDeviceFilter, ref iid, out o);

					if (hr != 0)
						o = null;
				}
				_videoStreamConfig = o as IAMStreamConfig;

				// Retrieve the stream control interface for the audio device
				o = null;
				cat = DsGuid.FromGuid(PinCategory.Capture);
				med = DsGuid.FromGuid(MediaType.Audio);
				iid = typeof(IAMStreamConfig).GUID;
				hr = _captureGraphBuilder.FindInterface(cat, med, _audioDeviceFilter, iid, out o);
//				hr = _captureGraphBuilder.FindInterface(ref cat, ref med, _audioDeviceFilter, ref iid, out o);
				if (hr != 0)
					o = null;

				_audioStreamConfig = o as IAMStreamConfig;

				// Retreive the media control interface (for starting/stopping graph)
				_mediaControl = (IMediaControl)_graphBuilder;

				// Reload any video crossbars
				if (_videoSources != null) 
					_videoSources.Dispose(); 
				_videoSources = null;

				// Reload any audio crossbars
				if (_audioSources != null) 
					_audioSources.Dispose(); 
				_audioSources = null;

				// Reload any property pages exposed by filters
				if (_propertyPages != null) 
					_propertyPages.Dispose(); 
				_propertyPages = null;

				// Reload capabilities of video device
				_videoCaps = null;

				// Reload capabilities of video device
				_audioCaps = null;

				// Retrieve TV Tuner if available
				o = null;
				cat = DsGuid.FromGuid(PinCategory.Capture);
				med = DsGuid.FromGuid(MediaType.Interleaved);
				iid = typeof(IAMTVTuner).GUID;
				hr = _captureGraphBuilder.FindInterface(cat, med, _videoDeviceFilter, iid, out o);
				//hr = _captureGraphBuilder.FindInterface(ref cat, ref med, _videoDeviceFilter, ref iid, out o);
				if (hr != 0)
				{
					med = DsGuid.FromGuid(MediaType.Video);
					hr = _captureGraphBuilder.FindInterface(cat, med, _videoDeviceFilter, iid, out o);
					//hr = _captureGraphBuilder.FindInterface(ref cat, ref med, _videoDeviceFilter, ref iid, out o);
					if (hr != 0)
						o = null;
				}
				
				IAMTVTuner t = o as IAMTVTuner;
				if (t != null)
					Tuner = new Tuner(t);


				/*
							// ----------- VMR 9 -------------------
							//## check out samples\inc\vmrutil.h :: RenderFileToVMR9

							IBaseFilter vmr = null;
							if ( ( VideoDevice != null ) && ( previewWindow != null ) )
							{
								vmr = (IBaseFilter) Activator.CreateInstance( Type.GetTypeFromCLSID( Clsid.VideoMixingRenderer9, true ) ); 
								hr = graphBuilder.AddFilter( vmr, "VMR" );
								if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

								IVMRFilterConfig9 vmrFilterConfig = (IVMRFilterConfig9) vmr;
								hr = vmrFilterConfig.SetRenderingMode( VMRMode9.Windowless );
								if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

								IVMRWindowlessControl9 vmrWindowsless = (IVMRWindowlessControl9) vmr;	
								hr = vmrWindowsless.SetVideoClippingWindow( previewWindow.Handle );
								if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );
							}
							//------------------------------------------- 

							// ---------- SmartTee ---------------------

							IBaseFilter smartTeeFilter = (IBaseFilter) Activator.CreateInstance( Type.GetTypeFromCLSID( Clsid.SmartTee, true ) ); 
							hr = graphBuilder.AddFilter( smartTeeFilter, "Video Smart Tee" );
							if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

							// Video -> SmartTee
							cat = PinCategory.Capture;
							med = MediaType.Video;
							hr = captureGraphBuilder.RenderStream( ref cat, ref med, videoDeviceFilter, null, smartTeeFilter ); 
							if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

							// smarttee -> mux
							cat = PinCategory.Capture;
							med = MediaType.Video;
							hr = captureGraphBuilder.RenderStream( ref cat, ref med, smartTeeFilter, null, muxFilter ); 
							if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

							// smarttee -> vmr
							cat = PinCategory.Preview;
							med = MediaType.Video;
							hr = captureGraphBuilder.RenderStream( ref cat, ref med, smartTeeFilter, null, vmr ); 
							if( hr < 0 ) Marshal.ThrowExceptionForHR( hr );

							// -------------------------------------
				*/
				// Update the state now that we are done
				_graphState = GraphState.Created;
			}
		}

		/// <summary>
		///  Connects the filters of a previously created graph 
		///  (created by createGraph()). Once rendered the graph
		///  is ready to be used. This method may also destroy
		///  streams if we have streams we no longer want.
		/// </summary>
		private void renderGraph()
		{
			DsGuid cat;
			DsGuid med;
			int hr;
			bool didSomething = false;

			assertStopped();

			// Ensure required properties set
			if (_filename == null)
				throw new ArgumentException("The Filename property has not been set to a file.\n");

			// Stop the graph
			if (_mediaControl != null)
				_mediaControl.Stop();

			// Create the graph if needed (group should already be created)
			createGraph();

			// Derender the graph if we have a capture or preview stream
			// that we no longer want. We can't derender the capture and 
			// preview streams seperately. 
			// Notice the second case will leave a capture stream intact
			// even if we no longer want it. This allows the user that is
			// not using the preview to Stop() and Start() without
			// rerendering the graph.
			if (!_wantPreviewRendered && _isPreviewRendered)
				derenderGraph();
			if (!_wantCaptureRendered && _isCaptureRendered)
				if (_wantPreviewRendered)
					derenderGraph();


			// Render capture stream (only if necessary)
			if (_wantCaptureRendered && !_isCaptureRendered)
			{
				// Render the file writer portion of graph (mux -> file)
				Guid mediaSubType = MediaSubType.Avi;
				Marshal.ThrowExceptionForHR(_captureGraphBuilder.SetOutputFileName(mediaSubType, Filename, out _muxFilter, out _fileWriterFilter));

				// Render video (video -> mux)
				if (VideoDevice != null)
				{
					// Try interleaved first, because if the device supports it,
					// it's the only way to get audio as well as video
					cat = DsGuid.FromGuid(PinCategory.Capture);
					med = DsGuid.FromGuid(MediaType.Interleaved);
					hr = _captureGraphBuilder.RenderStream(cat, med, _videoDeviceFilter, _videoCompressorFilter, _muxFilter);
					//hr = _captureGraphBuilder.RenderStream(ref cat, ref med, _videoDeviceFilter, _videoCompressorFilter, _muxFilter);
					if (hr < 0)
					{
						med = DsGuid.FromGuid(MediaType.Video);
						hr = _captureGraphBuilder.RenderStream(cat, med, _videoDeviceFilter, _videoCompressorFilter, _muxFilter);
						//hr = _captureGraphBuilder.RenderStream(ref cat, ref med, _videoDeviceFilter, _videoCompressorFilter, _muxFilter);
						if (hr == -2147220969) 
							throw new DeviceInUseException("Video device", hr);
						
						Marshal.ThrowExceptionForHR(hr);
					}
				}

				// Render audio (audio -> mux)
				if (AudioDevice != null)
				{
					cat = DsGuid.FromGuid(PinCategory.Capture);
					med = DsGuid.FromGuid(MediaType.Audio);
					Marshal.ThrowExceptionForHR(_captureGraphBuilder.RenderStream(cat, med, _audioDeviceFilter, _audioCompressorFilter, _muxFilter));
					//Marshal.ThrowExceptionForHR(_captureGraphBuilder.RenderStream(ref cat, ref med, _audioDeviceFilter, _audioCompressorFilter, _muxFilter));
				}

				_isCaptureRendered = true;
				didSomething = true;
			}

			// Render preview stream (only if necessary)
			if (_wantPreviewRendered && !_isPreviewRendered)
			{
				// Render preview (video -> renderer)
				cat = DsGuid.FromGuid(PinCategory.Preview);
				med = DsGuid.FromGuid(MediaType.Video);
				Marshal.ThrowExceptionForHR(_captureGraphBuilder.RenderStream(cat,  med, _videoDeviceFilter, null, null));
//				Marshal.ThrowExceptionForHR(_captureGraphBuilder.RenderStream(ref cat, ref med, _videoDeviceFilter, null, null));

				// Get the IVideoWindow interface
				_videoWindow = (IVideoWindow)_graphBuilder;
				
				// Set the video window to be a child of the main window
				Marshal.ThrowExceptionForHR(_videoWindow.put_Owner(_previewWindow.Handle));

				// Set video window style
				Marshal.ThrowExceptionForHR(_videoWindow.put_WindowStyle(WindowStyle.Child | WindowStyle.ClipChildren | WindowStyle.ClipSiblings));
				//Marshal.ThrowExceptionForHR(_videoWindow.put_WindowStyle(WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS));

				// Position video window in client rect of owner window
				_previewWindow.Resize += new EventHandler(onPreviewWindowResize);
				onPreviewWindowResize(this, null);

				// Make the video window visible, now that it is properly positioned
				Marshal.ThrowExceptionForHR(_videoWindow.put_Visible(OABool.True));

				_isPreviewRendered = true;
				didSomething = true;
			}

			if (didSomething)
				_graphState = GraphState.Rendered;
		}

		/// <summary>
		///  Setup and start the preview window if the user has
		///  requested it (by setting PreviewWindow).
		/// </summary>
		private void startPreviewIfNeeded()
		{
			// Render preview 
			if (_wantPreviewRendered && _isPreviewRendered && !_isCaptureRendered)
			{
				// Run the graph (ignore errors)
				// We can run the entire graph becuase the capture
				// stream should not be rendered (and that is enforced
				// in the if statement above)
				_mediaControl.Run();
			}
		}

		/// <summary>
		///  Disconnect and remove all filters except the device
		///  and compressor filters. This is the opposite of
		///  renderGraph(). Soem properties such as FrameRate
		///  can only be set when the device output pins are not
		///  connected. 
		/// </summary>
		private void derenderGraph()
		{
			// Stop the graph if it is running (ignore errors)
			if (_mediaControl != null)
				_mediaControl.Stop();

			// Free the preview window (ignore errors)
			if (_videoWindow != null)
			{
				_videoWindow.put_Visible(OABool.False);
				_videoWindow.put_Owner(IntPtr.Zero);
				_videoWindow = null;
			}

			// Remove the Resize event handler
			if (PreviewWindow != null)
				_previewWindow.Resize -= new EventHandler(onPreviewWindowResize);

			if ((int)_graphState >= (int)GraphState.Rendered)
			{
				// Update the state
				_graphState = GraphState.Created;
				_isCaptureRendered = false;
				_isPreviewRendered = false;

				// Disconnect all filters downstream of the 
				// video and audio devices. If we have a compressor
				// then disconnect it, but don't remove it
				if (_videoDeviceFilter != null)
					removeDownstream(_videoDeviceFilter, (_videoCompressor == null));
				if (_audioDeviceFilter != null)
					removeDownstream(_audioDeviceFilter, (_audioCompressor == null));

				// These filters should have been removed by the
				// calls above. (Is there anyway to check?)
				_muxFilter = null;
				_fileWriterFilter = null;
			}
		}

		/// <summary>
		///  Removes all filters downstream from a filter from the graph.
		///  This is called only by derenderGraph() to remove everything
		///  from the graph except the devices and compressors. The parameter
		///  "removeFirstFilter" is used to keep a compressor (that should
		///  be immediately downstream of the device) if one is begin used.
		/// </summary>
		private void removeDownstream(IBaseFilter filter, bool removeFirstFilter)
		{
			// Get a pin enumerator off the filter
			IEnumPins pinEnum;
			int hr = filter.EnumPins(out pinEnum);
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
						// Get the pin it is connected to
						IPin pinTo = null;
						pins[0].ConnectedTo(out pinTo);
						if (pinTo != null)
						{
							// Is this an input pin?
							PinInfo info = new PinInfo();
							hr = pinTo.QueryPinInfo(out info);
							if ((hr == 0) && (info.dir == (PinDirection.Input)))
							{
								// Recurse down this branch
								removeDownstream(info.filter, true);

								// Disconnect 
								_graphBuilder.Disconnect(pinTo);
								_graphBuilder.Disconnect(pins[0]);

								// Remove this filter
								// but don't remove the video or audio compressors
								if ((info.filter != _videoCompressorFilter) &&
									 (info.filter != _audioCompressorFilter))
									_graphBuilder.RemoveFilter(info.filter);
							}
							Marshal.ReleaseComObject(info.filter);
							Marshal.ReleaseComObject(pinTo);
						}
						Marshal.ReleaseComObject(pins[0]);
					}
				}
				while (hr == 0);

				Marshal.ReleaseComObject(pinEnum); 
				pinEnum = null;
			}
		}

		/// <summary>
		///  Completely tear down a filter graph and 
		///  release all associated resources.
		/// </summary>
		private void destroyGraph()
		{
			// Derender the graph (This will stop the graph
			// and release preview window. It also destroys
			// half of the graph which is unnecessary but
			// harmless here.) (ignore errors)
			try 
			{ 
				derenderGraph(); 
			}
			catch { }

			// Update the state after derender because it
			// depends on correct status. But we also want to
			// update the state as early as possible in case
			// of error.
			_graphState = GraphState.Null;
			_isCaptureRendered = false;
			_isPreviewRendered = false;

			// Remove graph from the ROT
			//if (_rotCookie != 0)
			//{
			//    DsROT.RemoveGraphFromRot(ref _rotCookie);
			//    _rotCookie = 0;
			//}

			// Remove filters from the graph
			// This should be unnecessary but the Nvidia WDM
			// video driver cannot be used by this application 
			// again unless we remove it. Ideally, we should
			// simply enumerate all the filters in the graph
			// and remove them. (ignore errors)
			if (_muxFilter != null)
				_graphBuilder.RemoveFilter(_muxFilter);
			if (_videoCompressorFilter != null)
				_graphBuilder.RemoveFilter(_videoCompressorFilter);
			if (_audioCompressorFilter != null)
				_graphBuilder.RemoveFilter(_audioCompressorFilter);
			if (_videoDeviceFilter != null)
				_graphBuilder.RemoveFilter(_videoDeviceFilter);
			if (_audioDeviceFilter != null)
				_graphBuilder.RemoveFilter(_audioDeviceFilter);

			// Clean up properties
			if (_videoSources != null)
				_videoSources.Dispose(); 
			_videoSources = null;
			if (_audioSources != null)
				_audioSources.Dispose(); 
			_audioSources = null;
			if (_propertyPages != null)
				_propertyPages.Dispose(); 
			_propertyPages = null;
			if (Tuner != null)
				Tuner.Dispose(); 
			Tuner = null;

			// Cleanup
			if (_graphBuilder != null)
				Marshal.ReleaseComObject(_graphBuilder); 
			_graphBuilder = null;
			if (_captureGraphBuilder != null)
				Marshal.ReleaseComObject(_captureGraphBuilder); 
			_captureGraphBuilder = null;
			if (_muxFilter != null)
				Marshal.ReleaseComObject(_muxFilter); 
			_muxFilter = null;
			if (_fileWriterFilter != null)
				Marshal.ReleaseComObject(_fileWriterFilter); 
			_fileWriterFilter = null;
			if (_videoDeviceFilter != null)
				Marshal.ReleaseComObject(_videoDeviceFilter); 
			_videoDeviceFilter = null;
			if (_audioDeviceFilter != null)
				Marshal.ReleaseComObject(_audioDeviceFilter); 
			_audioDeviceFilter = null;
			if (_videoCompressorFilter != null)
				Marshal.ReleaseComObject(_videoCompressorFilter); 
			_videoCompressorFilter = null;
			if (_audioCompressorFilter != null)
				Marshal.ReleaseComObject(_audioCompressorFilter); 
			_audioCompressorFilter = null;

			// These are copies of graphBuilder
			_mediaControl = null;
			_videoWindow = null;

			// For unmanaged objects we haven't released explicitly
			GC.Collect();
		}

		/// <summary> 
		/// Resize the preview when the PreviewWindow is resized 
		/// </summary>
		private void onPreviewWindowResize(object sender, EventArgs e)
		{
			if (_videoWindow != null)
			{
				// Position video window in client rect of owner window
				Rectangle rc = _previewWindow.ClientRectangle;
				_videoWindow.SetWindowPosition(0, 0, rc.Right, rc.Bottom);
			}
		}

		/// <summary> 
		///  Get a valid temporary filename (with path). We aren't using 
		///  Path.GetTempFileName() because it creates a 0-byte file 
		/// </summary>
		private string getTempFilename()
		{
			string s;
			try
			{
				int count = 0;
				int i;
				Random r = new Random();
				string tempPath = System.IO.Path.GetTempPath();
				do
				{
					i = r.Next();
					s = System.IO.Path.Combine(tempPath, i.ToString("X") + ".avi");
					count++;
					if (count > 100) 
						throw new InvalidOperationException("Unable to find temporary file.");
				} while (System.IO.File.Exists(s));
			}
			catch 
			{ 
				s = "c:\temp.avi"; 
			}
			return s;
		}

		/// <summary>
		///  Retrieves the value of one member of the IAMStreamConfig format block.
		///  Helper function for several properties that expose
		///  video/audio settings from IAMStreamConfig.GetFormat().
		///  IAMStreamConfig.GetFormat() returns a AMMediaType struct.
		///  AMMediaType.formatPtr points to a format block structure.
		///  This format block structure may be one of several 
		///  types, the type being determined by AMMediaType.formatType.
		/// </summary>
		private object getStreamConfigSetting(IAMStreamConfig streamConfig, string fieldName)
		{
			if (streamConfig == null)
				throw new NotSupportedException();
			assertStopped();
			derenderGraph();

			object returnValue = null;
			//IntPtr pmt = IntPtr.Zero;
			AMMediaType mediaType = new AMMediaType();

			try
			{
				// Get the current format info
				Marshal.ThrowExceptionForHR(streamConfig.GetFormat(out mediaType));
				//Marshal.PtrToStructure(pmt, mediaType);

				// The formatPtr member points to different structures
				// dependingon the formatType
				object formatStruct;
				if (mediaType.formatType == FormatType.WaveEx)
					formatStruct = new WaveFormatEx();
				else if (mediaType.formatType == FormatType.VideoInfo)
					formatStruct = new VideoInfoHeader();
				else if (mediaType.formatType == FormatType.VideoInfo2)
					formatStruct = new VideoInfoHeader2();
				else
					throw new NotSupportedException("This device does not support a recognized format block.");

				// Retrieve the nested structure
				Marshal.PtrToStructure(mediaType.formatPtr, formatStruct);

				// Find the required field
				Type structType = formatStruct.GetType();
				System.Reflection.FieldInfo fieldInfo = structType.GetField(fieldName);
				if (fieldInfo == null)
					throw new NotSupportedException("Unable to find the member '" + fieldName + "' in the format block.");

				// Extract the field's current value
				returnValue = fieldInfo.GetValue(formatStruct);
			}
			finally
			{
				DsUtils.FreeAMMediaType(mediaType);
				//Marshal.FreeCoTaskMem(pmt);
			}
			renderGraph();
			startPreviewIfNeeded();

			return returnValue;
		}

		/// <summary>
		///  Set the value of one member of the IAMStreamConfig format block.
		///  Helper function for several properties that expose
		///  video/audio settings from IAMStreamConfig.GetFormat().
		///  IAMStreamConfig.GetFormat() returns a AMMediaType struct.
		///  AMMediaType.formatPtr points to a format block structure.
		///  This format block structure may be one of several 
		///  types, the type being determined by AMMediaType.formatType.
		/// </summary>
		private object setStreamConfigSetting(IAMStreamConfig streamConfig, string fieldName, object newValue)
		{
			if (streamConfig == null)
				throw new NotSupportedException();
			assertStopped();
			derenderGraph();

			object returnValue = null;
			//IntPtr pmt = IntPtr.Zero;
			AMMediaType mediaType = new AMMediaType();

			try
			{
				// Get the current format info
				Marshal.ThrowExceptionForHR(streamConfig.GetFormat(out mediaType));
				//Marshal.PtrToStructure(pmt, mediaType);

				// The formatPtr member points to different structures
				// dependingon the formatType
				object formatStruct;
				if (mediaType.formatType == FormatType.WaveEx)
					formatStruct = new WaveFormatEx();
				else if (mediaType.formatType == FormatType.VideoInfo)
					formatStruct = new VideoInfoHeader();
				else if (mediaType.formatType == FormatType.VideoInfo2)
					formatStruct = new VideoInfoHeader2();
				else
					throw new NotSupportedException("This device does not support a recognized format block.");

				// Retrieve the nested structure
				Marshal.PtrToStructure(mediaType.formatPtr, formatStruct);

				// Find the required field
				Type structType = formatStruct.GetType();
				System.Reflection.FieldInfo fieldInfo = structType.GetField(fieldName);
				if (fieldInfo == null)
					throw new NotSupportedException("Unable to find the member '" + fieldName + "' in the format block.");

				// Update the value of the field
				fieldInfo.SetValue(formatStruct, newValue);

				// PtrToStructure copies the data so we need to copy it back
				Marshal.StructureToPtr(formatStruct, mediaType.formatPtr, false);

				// Save the changes
				Marshal.ThrowExceptionForHR(streamConfig.SetFormat(mediaType));
			}
			finally
			{
				DsUtils.FreeAMMediaType(mediaType);
				//Marshal.FreeCoTaskMem(pmt);
			}
			renderGraph();
			startPreviewIfNeeded();

			return returnValue;
		}

		/// <summary>
		///  Assert that the class is in a Stopped state.
		/// </summary>
		private void assertStopped()
		{
			if (!Stopped)
				throw new InvalidOperationException("This operation not allowed while Capturing. Please Stop the current capture.");
		}
	}
}
