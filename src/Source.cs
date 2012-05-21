using System;

namespace DirectShow.Capture
{
	/// <summary>
	///  Represents a physical connector or source on an audio/video device.
	/// </summary>
	public class Source : IDisposable
	{
		private bool _IsDisposed;

		/// <summary>Gets the name of the source. </summary>
		public string Name
		{ 
			get;
			protected set;
		}

		/// <summary>Gets whether this source is enabled. </summary>
		public virtual bool Enabled
		{
			get { throw new NotSupportedException("This method should be overriden in derrived classes."); }
			set { throw new NotSupportedException("This method should be overriden in derrived classes."); }
		}

		/// <summary> Release unmanaged resources. </summary>
		~Source()
		{
			if (!_IsDisposed)
			{
				Dispose(false);
				_IsDisposed = true;
			}
		}

		/// <summary> Obtains the String representation of this instance. </summary>
		public override string ToString()
		{
			return (Name);
		}

		/// <summary> Release unmanaged resources. </summary>
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
		}
	}
}
