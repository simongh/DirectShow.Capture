using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using DirectShowLib;

namespace DirectShow.Capture
{
	/// <summary>
	///	 A collection of Filter objects (DirectShow filters).
	///	 This is used by the <see cref="Capture"/> class to provide
	///	 lists of capture devices and compression filters.
	/// </summary>
	public class FilterCollection : System.Collections.ObjectModel.ReadOnlyCollection<Filter>
	{
		/// <summary> Populate the collection with a list of filters from a particular category. </summary>
		internal FilterCollection(Guid category)
			: base(new List<Filter>(FilterCollection.getFilters(category)))
		{ }

		/// <summary> Populate the InnerList with a list of filters from a particular category </summary>
		protected static IEnumerable<Filter> getFilters(Guid category)
		{
			object comObj = null;
			System.Runtime.InteropServices.ComTypes.IEnumMoniker enumMon = null;
			var mon = new System.Runtime.InteropServices.ComTypes.IMoniker[1];

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
				if (hr != 0)
					throw new NotSupportedException("No devices of the category");

				// Loop through the enumerator
				IntPtr f = IntPtr.Zero;
				do
				{
					// Next filter
					hr = enumMon.Next(1, mon, f);
					if ((hr != 0) || (mon[0] == null))
						break;

					// Add the filter
					Filter filter = new Filter(mon[0]);
					//InnerList.Add(filter);
					yield return filter;

					// Release resources
					Marshal.ReleaseComObject(mon[0]);
					mon[0] = null;
				}
				while (true);

				// Sort
				//InnerList.Sort();
			}
			finally
			{
				if (mon[0] != null)
					Marshal.ReleaseComObject(mon[0]); mon[0] = null;
				if (enumMon != null)
					Marshal.ReleaseComObject(enumMon); enumMon = null;
				if (comObj != null)
					Marshal.ReleaseComObject(comObj); comObj = null;
			}
		}
	}
}
