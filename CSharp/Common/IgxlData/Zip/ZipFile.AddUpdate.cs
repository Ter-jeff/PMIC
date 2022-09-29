using Ionic.Zip;
using System;
using System.Text;

namespace IgxlData.Zip
{
    public static class ZipFileExtention
    {
        public static ZipEntry AddUpdateEntry(this ZipFile zipFile, string entryName, string content)
        {
            if (zipFile.ContainsEntry(entryName))
                return UpdateEntry(zipFile, entryName, content, Encoding.Default);
            return zipFile.AddEntry(entryName, content, Encoding.Default);
        }

        private static ZipEntry UpdateEntry(ZipFile zipFile, string entryName, string content, Encoding encoding)
        {
            RemoveEntryForUpdate(zipFile, entryName);
            return zipFile.AddEntry(entryName, content, encoding);
        }

        private static void RemoveEntryForUpdate(ZipFile zipFile, string entryName)
        {
            if (String.IsNullOrEmpty(entryName))
                throw new ArgumentNullException("entryName");

            //string directoryPathInArchive = null;
            //if (entryName.IndexOf('\\') != -1)
            //{
            //    directoryPathInArchive = Path.GetDirectoryName(entryName);
            //    entryName = Path.GetFileName(entryName);
            //}
            //var key = ZipEntry.NameInArchive(entryName, directoryPathInArchive);
            //if (zipFile[key] != null)
            zipFile.RemoveEntry(entryName);
        }
    }
}