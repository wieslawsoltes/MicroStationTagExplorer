using System;
using System.Runtime.InteropServices;

namespace MicroStationTagExplorer
{
    internal static class ComUtilities
    {
        public static T CreateObject<T>(string progID)
        {
            return (T)Activator.CreateInstance(Type.GetTypeFromProgID(progID));
        }

        public static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
