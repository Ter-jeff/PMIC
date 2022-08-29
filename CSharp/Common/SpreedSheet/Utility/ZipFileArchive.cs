#define WPF

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using Ionic.Zip;

namespace unvell.ReoGrid.Utility
{
    internal interface IZipArchive
    {
        IZipEntry GetFile(string path);
        IZipEntry AddFile(string path, Stream stream = null);
        bool IsFileExist(string path);
        void Flush();
        void Close();
    }

    internal interface IZipEntry
    {
        Stream GetStream();
        Stream CreateStream();
    }

    internal class NET35ZipArchive : IZipArchive, IDisposable
    {
        public enum CompressionMethodEnum
        {
            Stored,
            Deflated
        }

        public enum DeflateOptionEnum
        {
            Normal,
            Maximum,
            Fast,
            SuperFast
        }

        internal static CachedTypeWrapper zipArchiveWrapper;
        internal static Type msCompressionMethodEnumType;
        internal static Type msDeflateOptionEnumType;
        internal object external;

        public IEnumerable<NET35ZipFileEntry> Files
        {
            get
            {
                var coll = zipArchiveWrapper.Invoke(external, "GetFiles") as IEnumerable; //ZipFileInfoCollection

                if (coll != null)
                    foreach (var p in coll)
                        yield return new NET35ZipFileEntry(p);
            }
        }

        public IEnumerable<string> FileNames
        {
            get { return Files.Select(p => p.Name).OrderBy(p => p); }
        }

        public void Dispose()
        {
            ((IDisposable)external).Dispose();
        }

        public IZipEntry AddFile(string path, Stream stream = null)
        {
            return AddFile(path, CompressionMethodEnum.Deflated);
        }

        public bool IsFileExist(string path)
        {
            return (bool)zipArchiveWrapper.Invoke(external, "FileExists", path);
        }

        public IZipEntry GetFile(string path)
        {
            if (path.StartsWith("/")) path = path.Substring(1);
            return new NET35ZipFileEntry(zipArchiveWrapper.Invoke(external, "GetFile", path));
        }

        public void Flush()
        {
            zipArchiveWrapper.Invoke(external, "Flush");
        }

        public void Close()
        {
            zipArchiveWrapper.Invoke(external, "Close");
        }

        public NET35ZipFileEntry AddFile(string path,
            CompressionMethodEnum compmeth = CompressionMethodEnum.Deflated,
            DeflateOptionEnum option = DeflateOptionEnum.Normal)
        {
            var comp = msCompressionMethodEnumType.GetField(compmeth.ToString()).GetValue(null);
            var opti = msDeflateOptionEnumType.GetField(option.ToString()).GetValue(null);

            return new NET35ZipFileEntry(zipArchiveWrapper.Invoke(external, "AddFile", path, comp, opti));
        }

        public void DeleteFile(string name)
        {
            zipArchiveWrapper.StaticInvoke("DeleteFile", external, name);
        }
    }

#if WINFORM || WPF
    /// <summary>
    ///     Original Document: http://www.codeproject.com/Articles/209731/Csharp-use-Zip-archives-without-external-libraries
    /// </summary>
    internal class NET35ZipArchiveFactory
    {
        static NET35ZipArchiveFactory()
        {
            var msZipArchiveType = typeof(Package).Assembly.GetType("MS.Internal.IO.Zip.ZipArchive");
            NET35ZipArchive.zipArchiveWrapper = new CachedTypeWrapper(msZipArchiveType);
            NET35ZipArchive.msCompressionMethodEnumType =
                msZipArchiveType.Assembly.GetType("MS.Internal.IO.Zip.CompressionMethodEnum");
            NET35ZipArchive.msDeflateOptionEnumType =
                msZipArchiveType.Assembly.GetType("MS.Internal.IO.Zip.DeflateOptionEnum");
        }

        // prevent construct from outside
        private NET35ZipArchiveFactory()
        {
        }

        public static IZipArchive OpenOnFile(string path, FileMode mode = FileMode.Open,
            FileAccess access = FileAccess.Read,
            FileShare share = FileShare.Read, bool streaming = false)
        {
            return new NET35ZipArchive
            {
                external = NET35ZipArchive.zipArchiveWrapper.StaticInvoke("OpenOnFile", path, mode, access, share,
                    streaming)
            };
        }

        public static IZipArchive OpenOnStream(Stream stream, FileMode mode = FileMode.Open,
            FileAccess access = FileAccess.Read, bool streaming = false)
        {
            return new NET35ZipArchive
            {
                external = NET35ZipArchive.zipArchiveWrapper.StaticInvoke("OpenOnStream", stream, mode, access,
                    streaming)
            };
        }
    }
#endif // WINFORM || WPF

    internal class MZipArchiveFactory
    {
        public static IZipArchive OpenOnStream(Stream stream)
        {
            return MZipArchive.OpenOnStream(stream);
        }

        public static IZipArchive CreateOnStream(Stream stream)
        {
            return MZipArchive.CreateOnStream(stream);
        }
    }

    internal class MZipArchive : IZipArchive
    {
        private Stream stream;
        private ZipFile zip;

        private MZipArchive()
        {
        }

        public IZipEntry GetFile(string path)
        {
            var entry = zip.SingleOrDefault(e => e.FileName == path);
            if (entry == null) return null;

            return new MZipEntry(entry);
        }

        public IZipEntry AddFile(string path, Stream stream)
        {
            return new MZipEntry(zip.AddEntry(path, stream));
        }

        public bool IsFileExist(string path)
        {
            return zip.Any(entry => entry.FileName == path);
        }

        public void Flush()
        {
            zip.Save(stream);
        }

        public void Close()
        {
            zip.Dispose();
        }

        internal static MZipArchive OpenOnStream(Stream stream)
        {
            return new MZipArchive
            {
                zip = ZipFile.Read(stream),
                stream = stream
            };
        }

        internal static MZipArchive CreateOnStream(Stream stream)
        {
            var mzip = new MZipArchive
            {
                zip = new ZipFile(),
                stream = stream
            };

            return mzip;
        }
    }

    internal class MZipEntry : IZipEntry
    {
        private readonly ZipEntry entry;

        internal MZipEntry(ZipEntry entry)
        {
            this.entry = entry;
        }

        public Stream CreateStream()
        {
            return null;
        }

        public Stream GetStream()
        {
            var ms = new MemoryStream();
            entry.Extract(ms);
            ms.Position = 0;
            return ms;
        }
    }

    internal class NET35ZipFileEntry : IZipEntry
    {
        private static CachedTypeWrapper zipEntry;

        private readonly object external;

        internal NET35ZipFileEntry(object external)
        {
            if (zipEntry == null) zipEntry = new CachedTypeWrapper(external.GetType());

            this.external = external;
        }

        public string Name
        {
            get { return (string)GetProperty("Name"); }
        }

        public DateTime LastModFileDateTime
        {
            get { return (DateTime)GetProperty("LastModFileDateTime"); }
        }

        public bool FolderFlag
        {
            get { return (bool)GetProperty("FolderFlag"); }
        }

        public bool VolumeLabelFlag
        {
            get { return (bool)GetProperty("VolumeLabelFlag"); }
        }

        public object CompressionMethod
        {
            get { return GetProperty("CompressionMethod"); }
            set { SetProperty("CompressionMethod", value); }
        }

        public object DeflateOption
        {
            get { return GetProperty("DeflateOption"); }
        }

        public Stream GetStream()
        {
            return GetStream(FileMode.Open);
        }

        public Stream CreateStream()
        {
            return GetStream(FileMode.Create, FileAccess.Write);
        }

        private object GetProperty(string name)
        {
            return zipEntry.GetProperty(external, name);
        }

        private void SetProperty(string name, object value)
        {
            zipEntry.SetProperty(external, name, value);
        }

        public override string ToString()
        {
            return Name; // base.ToString();
        }

        public Stream GetStream(FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read)
        {
            return zipEntry.Invoke(external, "GetStream", mode, access) as Stream;
        }
    }

    internal class CachedTypeWrapper
    {
        private readonly Dictionary<string, MethodInfo> cachedMethods = new Dictionary<string, MethodInfo>();
        private readonly Dictionary<string, PropertyInfo> cachedProperties = new Dictionary<string, PropertyInfo>();

        public CachedTypeWrapper(Type type)
        {
            Type = type;
        }

        public Type Type { get; }

        public object StaticInvoke(string name, params object[] args)
        {
            return Invoke(null, name, args);
        }

        public object Invoke(object instance, string name, params object[] args)
        {
            MethodInfo mi = null;

            if (!cachedMethods.TryGetValue(name, out mi))
                mi = Type.GetMethod(name,
                    (instance == null ? BindingFlags.Static : BindingFlags.Instance) | BindingFlags.NonPublic);

            return mi == null ? null : mi.Invoke(instance, args);
        }

        public object GetProperty(object instance, string name)
        {
            PropertyInfo pi = null;

            if (!cachedProperties.TryGetValue(name, out pi))
            {
                pi = Type.GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic);

                cachedProperties[name] = pi;
            }

            return pi == null ? null : pi.GetValue(instance, null);
        }

        public void SetProperty(object instance, string name, object value)
        {
            PropertyInfo pi = null;

            if (!cachedProperties.TryGetValue(name, out pi))
            {
                pi = Type.GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic);

                cachedProperties[name] = pi;
            }

            if (pi != null) pi.SetValue(instance, name, null);
        }
    }

    /// <summary>
    ///     Helper class for compress and decompress zip stream
    /// </summary>
    /// <remarks>Provided by Rick Meyer</remarks>
    public class ZipStreamHelper
    {
        /// <summary>
        ///     Decompress a byte array
        /// </summary>
        /// <param name="zippedData">Compressed Byte Array</param>
        /// <returns>Decompressed Byte Array</returns>
        public static byte[] Decompress(byte[] zippedData)
        {
            using (var outputStream = new MemoryStream())
            {
                using (var inputStream = new MemoryStream(zippedData))
                {
                    using (var zip = new GZipStream(inputStream, CompressionMode.Decompress))
                    {
                        //zip.CopyTo(outputStream);    // cannot be used in .NET 3.5
                        var readBytes = 0;
                        var buf = new byte[4096];
                        while ((readBytes = zip.Read(buf, 0, buf.Length)) > 0) outputStream.Write(buf, 0, readBytes);

                        return outputStream.ToArray();
                    }
                }
            }
        }

        /// <summary>
        ///     Compress a byte Array using Gzip
        /// </summary>
        /// <param name="plainData">The byte array to compress</param>
        /// <returns>Returns a compressed byte array</returns>
        public static byte[] Compress(byte[] plainData)
        {
            if (plainData == null)
                throw new ArgumentNullException("Tried to compress null byte array - can't get smaller than zero!");
            byte[] compressesData = null;
            using (var outputStream = new MemoryStream())
            {
                using (var zip = new GZipStream(outputStream, CompressionMode.Compress))
                {
                    zip.Write(plainData, 0, plainData.Length);
                }

                //Dont get the MemoryStream data before the GZipStream is closed 
                //since it doesn’t yet contain complete compressed data.
                //GZipStream writes additional data including footer information when its been disposed

                compressesData = outputStream.ToArray();
            }

            return compressesData;
        }
    }
}