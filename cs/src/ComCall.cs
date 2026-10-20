// 直接调 IDispatch::Invoke —— 只为了那几个"有 ByRef 出参"的 VBA 方法。
//
// ★为什么非写这一层不可：
//   `CodePane.GetSelection(StartLine, StartColumn, EndLine, EndColumn)` 的四个
//   参数全是 `ByRef ... As Long`。C# 的 dynamic 后期绑定走的是
//   ComBinder -> IDispatch，它【不支持】带出参的调用（实测报
//   "Error while invoking GetSelection."）。而 `CodePane.Selection` 这个本该
//   返回 TextSelection 的属性，在 IDispatch 上根本查不到（GetIDsOfNames 说不存在）。
//   ⇒ 想让"光标在哪"这件事有解，就只能自己拼 DISPPARAMS，把四个
//      VT_BYREF|VT_I4 的 VARIANT 递进去。
//
// 只支持本项目实际用到的形状：全 byref Long 出参、无入参。不试图做成通用 COM 框架。

using System;
using System.Runtime.InteropServices;

namespace VbeComplete
{
    internal static class ComCall
    {
        private const ushort VT_BYREF = 0x4000;
        private const ushort VT_I4 = 0x0003;
        private const ushort DISPATCH_METHOD = 0x01;

        // x64 上 VARIANT 是 24 字节（vt 2 + 保留 6 + union 16）。
        // 本项目只发布 win-x64；x86 下这里会是错的，所以进程位数不对就整体放弃。
        [StructLayout(LayoutKind.Explicit, Size = 24)]
        private struct VARIANT
        {
            [FieldOffset(0)] public ushort vt;
            [FieldOffset(8)] public IntPtr byref;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPPARAMS
        {
            public IntPtr rgvarg;
            public IntPtr rgdispidNamedArgs;
            public int cArgs;
            public int cNamedArgs;
        }

        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatchRaw
        {
            [PreserveSig] int GetTypeInfoCount(out uint pctinfo);
            [PreserveSig] int GetTypeInfo(uint iTInfo, int lcid, out IntPtr ppTInfo);
            [PreserveSig] int GetIDsOfNames(ref Guid riid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
                uint cNames, int lcid,
                [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig] int Invoke(int dispIdMember, ref Guid riid, int lcid, ushort wFlags,
                ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
        }

        private static readonly bool _ok = Environment.Is64BitProcess;

        /// <summary>
        /// 调一个"四个 ByRef Long 出参"的方法（目前只有 GetSelection 是这个形状）。
        /// 顺序按源码参数顺序给出；内部按 COM 规矩逆序摆放。
        /// </summary>
        public static bool CallFourOut(object target, string method,
            out int a, out int b, out int c, out int d)
        {
            a = b = c = d = 0;
            if (!_ok || target == null || string.IsNullOrEmpty(method)) return false;

            IntPtr pDisp = IntPtr.Zero;
            IntPtr args = IntPtr.Zero;
            IntPtr ints = IntPtr.Zero;
            try
            {
                pDisp = Marshal.GetIDispatchForObject(target);
                var disp = (IDispatchRaw)Marshal.GetObjectForIUnknown(pDisp);
                if (disp == null) return false;

                int[] ids = new int[1];
                Guid iid = Guid.Empty;
                int hr = disp.GetIDsOfNames(ref iid, new[] { method }, 1, 0, ids);
                if (hr != 0 || ids[0] < 0) return false;

                const int n = 4;
                int varSize = Marshal.SizeOf(typeof(VARIANT));
                args = Marshal.AllocHGlobal(varSize * n);
                ints = Marshal.AllocHGlobal(sizeof(int) * n);
                for (int i = 0; i < n; i++) Marshal.WriteInt32(ints, i * sizeof(int), 0);
                for (int i = 0; i < varSize * n; i++) Marshal.WriteByte(args, i, 0);

                // rgvarg 是【逆序】的：最后一个参数在最前面
                for (int i = 0; i < n; i++)
                {
                    int paramIndex = n - 1 - i;
                    IntPtr slot = IntPtr.Add(args, i * varSize);
                    Marshal.WriteInt16(slot, 0, unchecked((short)(VT_BYREF | VT_I4)));
                    Marshal.WriteIntPtr(slot, 8, IntPtr.Add(ints, paramIndex * sizeof(int)));
                }

                DISPPARAMS dp = new DISPPARAMS
                {
                    rgvarg = args,
                    rgdispidNamedArgs = IntPtr.Zero,
                    cArgs = n,
                    cNamedArgs = 0,
                };

                hr = disp.Invoke(ids[0], ref iid, 0, DISPATCH_METHOD, ref dp,
                    IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                if (hr != 0) return false;

                a = Marshal.ReadInt32(ints, 0);
                b = Marshal.ReadInt32(ints, 1 * sizeof(int));
                c = Marshal.ReadInt32(ints, 2 * sizeof(int));
                d = Marshal.ReadInt32(ints, 3 * sizeof(int));
                return true;
            }
            catch (Exception e)
            {
                Log.Write("IDispatch 调用 {0} 失败：{1}", method, e.Message);
                return false;
            }
            finally
            {
                if (ints != IntPtr.Zero) Marshal.FreeHGlobal(ints);
                if (args != IntPtr.Zero) Marshal.FreeHGlobal(args);
                if (pDisp != IntPtr.Zero) Marshal.Release(pDisp);
            }
        }
    }
}
