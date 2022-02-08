using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace NPOI.Util
{
	public static class RecyclableMemory
	{
		private static Microsoft.IO.RecyclableMemoryStreamManager _memoryManager;
		private static bool _dataInitialized = false;
		private static object _dataLock = new object();

		private static Microsoft.IO.RecyclableMemoryStreamManager MemoryManager
		{
			get
			{
				return LazyInitializer.EnsureInitialized(ref _memoryManager, ref _dataInitialized, ref _dataLock);
			}
		}

		public static void SetRecyclableMemoryStreamManager(Microsoft.IO.RecyclableMemoryStreamManager recyclableMemoryStreamManager)
		{
			_dataInitialized = recyclableMemoryStreamManager is object;
			_memoryManager = recyclableMemoryStreamManager;
		}

		public static MemoryStream GetStream()
		{
			return MemoryManager.GetStream();
		}
		public static MemoryStream GetStream(byte[] array)
		{
			return MemoryManager.GetStream(array);
		}

		public static MemoryStream GetStream(int capacity)
		{
			return MemoryManager.GetStream(null, capacity);
		}
	}
}