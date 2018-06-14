using System;

namespace MicroStationTagExplorer
{
    public static class Utilities
    {
        public static T CreateObject<T>(string progID)
        {
            return (T)Activator.CreateInstance(Type.GetTypeFromProgID(progID));
        }
    }
}
