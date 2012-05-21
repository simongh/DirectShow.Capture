using System;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///  Represents a DirectShow filter (e.g. video capture device, 
	///  compression codec).
	/// </summary>
	/// <remarks>
	///  To save a chosen filer for later recall
	///  save the MonikerString property on the filter: 
	///  <code><div style="background-color:whitesmoke;">
	///   string savedMonikerString = myFilter.MonikerString;
	///  </div></code>
	///  
	///  To recall the filter create a new Filter class and pass the 
	///  string to the constructor: 
	///  <code><div style="background-color:whitesmoke;">
	///   Filter mySelectedFilter = new Filter( savedMonikerString );
	///  </div></code>
	/// </remarks>
	public class Filter : IComparable
	{
		/// <summary>Gets the human-readable name of the filter </summary>
		public string Name
		{
			get;
			private set;
		}

		/// <summary>Gets the unique string referencing this filter. This string can be used to recreate this filter. </summary>
		public string MonikerString
		{
			get;
			private set;
		}

		/// <summary> Create a new filter from its moniker string. </summary>
		public Filter(string monikerString)
		{
			Name = getName(monikerString);
			MonikerString = monikerString;
		}

		/// <summary> Create a new filter from its moniker </summary>
		internal Filter(System.Runtime.InteropServices.ComTypes.IMoniker moniker)
		{
			Name = getName(moniker);
			MonikerString = getMonikerString(moniker);
		}

		/// <summary> Retrieve the a moniker's display name (i.e. it's unique string) </summary>
		protected string getMonikerString(System.Runtime.InteropServices.ComTypes.IMoniker moniker)
		{
			string s;
			moniker.GetDisplayName(null, null, out s);
			return (s);
		}

		/// <summary> Retrieve the human-readable name of the filter </summary>
		protected string getName(System.Runtime.InteropServices.ComTypes.IMoniker moniker)
		{
			object bagObj = null;
			IPropertyBag bag = null;
			try
			{
				Guid bagId = typeof(IPropertyBag).GUID;
				
				moniker.BindToStorage(null, null, ref bagId, out bagObj);
				bag = (IPropertyBag)bagObj;
				
				object val = "";
				int hr = bag.Read("FriendlyName", out val, null);
				//int hr = bag.Read("FriendlyName", ref val, IntPtr.Zero);
				Marshal.ThrowExceptionForHR(hr);

				string ret = val as string;
				if ((ret == null) || (ret.Length < 1))
					throw new NotImplementedException("Device FriendlyName");
				return (ret);
			}
			catch (Exception)
			{
				return ("");
			}
			finally
			{
				bag = null;
				if (bagObj != null)
					Marshal.ReleaseComObject(bagObj); bagObj = null;
			}
		}

		/// <summary> Get a moniker's human-readable name based on a moniker string. </summary>
		protected string getName(string monikerString)
		{
			System.Runtime.InteropServices.ComTypes.IMoniker parser = null;
			System.Runtime.InteropServices.ComTypes.IMoniker moniker = null;
			try
			{
				parser = getAnyMoniker();
				int eaten;
				parser.ParseDisplayName(null, null, monikerString, out eaten, out moniker);
				return (getName(parser));
			}
			finally
			{
				if (parser != null)
					Marshal.ReleaseComObject(parser); parser = null;
				if (moniker != null)
					Marshal.ReleaseComObject(moniker); moniker = null;
			}
		}

		/// <summary>
		///  This method gets a UCOMIMoniker object.
		/// 
		///  HACK: The only way to create a UCOMIMoniker from a moniker 
		///  string is to use UCOMIMoniker.ParseDisplayName(). So I 
		///  need ANY UCOMIMoniker object so that I can call 
		///  ParseDisplayName(). Does anyone have a better solution?
		/// 
		///  This assumes there is at least one video compressor filter
		///  installed on the system.
		/// </summary>
		protected System.Runtime.InteropServices.ComTypes.IMoniker getAnyMoniker()
		{
			Guid category = FilterCategory.VideoCompressorCategory;
			object comObj = null;
			System.Runtime.InteropServices.ComTypes.IEnumMoniker enumMon = null;

			try
			{
				// Get the system device enumerator
				//Type srvType = Type.GetTypeFromCLSID(Clsid.SystemDeviceEnum);
				//if (srvType == null)
				//	throw new NotImplementedException("System Device Enumerator");
				//comObj = Activator.CreateInstance(srvType);

				ICreateDevEnum enumDev = (ICreateDevEnum)new CreateDevEnum();

				// Create an enumerator to find filters in category
				int hr = enumDev.CreateClassEnumerator(category, out enumMon, 0);
				//int hr = enumDev.CreateClassEnumerator(ref category, out enumMon, 0);
				if (hr != 0)
					throw new NotSupportedException("No devices of the category");

				// Get first filter
				IntPtr f = IntPtr.Zero;
				var mon = new System.Runtime.InteropServices.ComTypes.IMoniker[1];
				hr = enumMon.Next(1, mon, f);
				if ((hr != 0))
					mon[0] = null;

				return (mon[0]);
			}
			finally
			{
				if (enumMon != null)
					Marshal.ReleaseComObject(enumMon); enumMon = null;
				if (comObj != null)
					Marshal.ReleaseComObject(comObj); comObj = null;
			}
		}

		/// <summary>
		///  Compares the current instance with another object of 
		///  the same type.
		/// </summary>
		public int CompareTo(object obj)
		{
			if (obj == null)
				return (1);

			Filter f = (Filter)obj;
			return this.Name.CompareTo(f.Name);
		}
	}
}
