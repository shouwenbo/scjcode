using System;
using System.Runtime.InteropServices;
using System.Text;

namespace BatchTranslateApp
{

    public static class ClipboardHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalFree(IntPtr hMem);

        private const uint CF_UNICODETEXT = 13;
        private const uint GMEM_MOVEABLE = 0x0002;

        public static void SetText(string text)
        {
            if (!OpenClipboard(IntPtr.Zero))
            {
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }

            try
            {
                if (!EmptyClipboard())
                {
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }

                IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)(Encoding.Unicode.GetByteCount(text) + 2));
                if (hGlobal == IntPtr.Zero)
                {
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }

                IntPtr target = GlobalLock(hGlobal);
                if (target == IntPtr.Zero)
                {
                    GlobalFree(hGlobal);
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }

                try
                {
                    Marshal.Copy(Encoding.Unicode.GetBytes(text), 0, target, Encoding.Unicode.GetByteCount(text));
                    Marshal.WriteInt16(target, Encoding.Unicode.GetByteCount(text), 0);
                }
                finally
                {
                    GlobalUnlock(hGlobal);
                }

                if (SetClipboardData(CF_UNICODETEXT, hGlobal) == IntPtr.Zero)
                {
                    GlobalFree(hGlobal);
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }
            }
            finally
            {
                CloseClipboard();
            }
        }
    }

}
