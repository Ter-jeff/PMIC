#define WPF


using System;
using System.Security.Cryptography;
using System.Text;
#if WINFORM || WPF
using unvell.Common.Win32Lib;
#endif // WINFORM || WPF

namespace unvell.Common
{
    /// <summary>
    ///     Common Toolkit
    /// </summary>
    public static class Toolkit
    {
        private static MD5 md5;

        /// <summary>
        ///     Default font size list.
        /// </summary>
        public static readonly float[] FontSizeList =
        {
            5f, 6f, 7f, 8f, 9f, 10f, 10.5f, 11f, 11.5f, 12f, 12.5f, 14f, 16f, 18f,
            20f, 22f, 24f, 26f, 28f, 30f, 32f, 34f, 38f, 46f, 58f, 64f, 78f, 92f
        };
#if WINFORM || WPF
        /// <summary>
        ///     Check whether or not the specified key is pressed.
        /// </summary>
        /// <param name="vkey">Windows virtual key.</param>
        /// <returns>true if pressed, otherwise false if not pressed.</returns>
        public static bool IsKeyDown(Win32.VKey vkey)
        {
            return ((Win32.GetKeyState(vkey) >> 15) & 1) == 1;
        }
#endif // WINFORM || WPF

        internal static byte[] GetMD5Hash(string str)
        {
            if (md5 == null) md5 = MD5.Create();

            return md5.ComputeHash(Encoding.Default.GetBytes(str));
        }

        internal static string GetHexString(byte[] data)
        {
            return Convert.ToBase64String(data);
        }

        internal static string GetMD5HashedString(string str)
        {
            return GetHexString(GetMD5Hash(str));
        }

        internal static double Ceiling(double val, double scale)
        {
            var m = val % scale;
            if (m == 0) return val;

            return val - m + scale;
        }
    }
}