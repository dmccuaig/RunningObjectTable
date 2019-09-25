using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace RunningObjectTable
{
    public static class RotUtil
    {
        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(uint reserved, out IBindCtx ctx);

        public static IBindCtx CreateBindCtx()
        {
            if (CreateBindCtx(0, out IBindCtx ctx) != 0 || ctx == null)
            {
                throw new ApplicationException($"{nameof(CreateBindCtx)} failed");
            }
            return ctx;
        }

        public static IRunningObjectTable GetRunningObjectTable()
        {
            var ctx = CreateBindCtx();
            ctx.GetRunningObjectTable(out IRunningObjectTable rot);
            if (rot == null)
            {
                throw new ApplicationException($"Could not get {nameof(IRunningObjectTable)}");
            }
            return rot;
        }

        public static IEnumMoniker GetEnumMoniker()
        {
            var rot = GetRunningObjectTable();
            rot.EnumRunning(out IEnumMoniker enumMoniker);
            if (enumMoniker == null)
            {
                throw new ApplicationException($"Could not get {nameof(IEnumMoniker)}");
            }

            enumMoniker.Reset();
            return enumMoniker;
        }

        public static bool IsRunning(string progID)
        {
            string clsId = Type.GetTypeFromProgID(progID).GUID.ToString().ToUpper();

            var enumMoniker = GetEnumMoniker();

            IMoniker[] monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                var ctx = CreateBindCtx();
                var moniker = monikers[0];

                moniker.GetDisplayName(ctx, null, out string name);
                if (name.ToUpper().IndexOf(clsId, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

    }
}