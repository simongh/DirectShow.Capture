using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Provides collections of devices and compression codecs
	///  installed on the system. 
	/// </summary>
	/// <example>
	///  Devices and compression codecs are implemented in DirectShow 
	///  as filters, see the <see cref="Filter"/> class for more 
	///  information. To list the available video devices:
	///  <code><div style="background-color:whitesmoke;">
	///   Filters filters = new Filters();
	///   foreach ( Filter f in filters.VideoInputDevices )
	///   {
	///		Debug.WriteLine( f.Name );
	///   }
	///  </div></code>
	///  <seealso cref="Filter"/>
	/// </example>
	public class Filters
	{
		/// <summary>Gets a collection of available video capture devices. </summary>
		public FilterCollection VideoInputDevices
		{
			get;
			private set;
		}

		/// <summary>Gets a collection of available audio capture devices. </summary>
		public FilterCollection AudioInputDevices
		{
			get;
			private set;
		}

		/// <summary>Gets a collection of available video compressors. </summary>
		public FilterCollection VideoCompressors
		{
			get;
			private set;
		}

		/// <summary>Gets a collection of available audio compressors. </summary>
		public FilterCollection AudioCompressors
		{
			get;
			private set;
		}

		public Filters()
		{
			VideoInputDevices = new FilterCollection(FilterCategory.VideoInputDevice);
			AudioInputDevices = new FilterCollection(FilterCategory.AudioInputDevice);
			VideoCompressors = new FilterCollection(FilterCategory.VideoCompressorCategory);
			AudioCompressors = new FilterCollection(FilterCategory.AudioCompressorCategory);
		}
	}
}
