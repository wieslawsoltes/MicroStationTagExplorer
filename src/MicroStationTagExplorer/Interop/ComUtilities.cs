using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace MicroStationTagExplorer
{
    public static class ComUtilities
    {
        public static Type GetProgIDType(string progID)
        {
            return Type.GetTypeFromProgID(progID);
        }

        public static T CreateObject<T>(string progID)
        {
            Type type = GetProgIDType(progID);
            if (type != null)
            {
                return (T)Activator.CreateInstance(type);
            }
            return default(T);
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

        public static List<string> GetProgIDs()
        {
            var regClis = Registry.ClassesRoot.OpenSubKey("CLSID");
            var progs = new List<string>();

            foreach (var clsid in regClis.GetSubKeyNames())
            {
                var regClsidKey = regClis.OpenSubKey(clsid);
                var ProgID = regClsidKey.OpenSubKey("ProgID");
                var regPath = regClsidKey.OpenSubKey("InprocServer32");

                if (regPath == null)
                    regPath = regClsidKey.OpenSubKey("LocalServer32");

                if (regPath != null && ProgID != null)
                {
                    var pid = ProgID.GetValue("");
                    var filePath = regPath.GetValue("");
                    progs.Add(pid + " -> " + filePath);
                    regPath.Close();
                }

                regClsidKey.Close();
            }

            return progs;
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        public static object GetRunningCOMObjectByName(string objectDisplayName)
        {
            IRunningObjectTable runningObjectTable = null;
            IEnumMoniker monikerList = null;

            try
            {
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    IBindCtx bindInfo;
                    string displayName;

                    CreateBindCtx(0, out bindInfo);
                    monikerContainer[0].GetDisplayName(bindInfo, null, out displayName);
                    Marshal.ReleaseComObject(bindInfo);

                    if (displayName.IndexOf(objectDisplayName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        object comInstance;
                        runningObjectTable.GetObject(monikerContainer[0], out comInstance);
                        return comInstance;
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);

                if (monikerList != null)
                    Marshal.ReleaseComObject(monikerList);
            }
            return null;
        }

        public static List<string> GetGUIDs(string[] progIDs)
        {
            List<string> guids = new List<string>();

            foreach (string progID in progIDs)
            {
                Type type = GetProgIDType(progID);
                if (type != null)
                {
                    guids.Add(type.GUID.ToString());
                }
            }

            return guids;
        }

        public static List<object> GetRunningCOMObjects(string[] progIDs)
        {
            IRunningObjectTable runningObjectTable = null;
            IEnumMoniker enumMoniker = null;

            try
            {
                List<object> results = new List<object>();
                List<string> guids = GetGUIDs(progIDs);

                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                runningObjectTable.EnumRunning(out enumMoniker);
                enumMoniker.Reset();

                IMoniker[] moniker = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                while (enumMoniker.Next(1, moniker, pointerFetchedMonikers) == 0)
                {
                    IBindCtx bindInfo;
                    string displayName;

                    CreateBindCtx(0, out bindInfo);
                    moniker[0].GetDisplayName(bindInfo, null, out displayName);
                    Marshal.ReleaseComObject(bindInfo);

                    foreach (string guid in guids)
                    {
                        if (displayName.IndexOf(guid, StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            object runningObject;
                            runningObjectTable.GetObject(moniker[0], out runningObject);
                            if (runningObject != null)
                            {
                                results.Add(runningObject);
                                break;
                            }
                        }
                    }
                }
                return results;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);

                if (enumMoniker != null)
                    Marshal.ReleaseComObject(enumMoniker);
            }
        }
    }
}
