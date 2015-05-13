using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace DocumentComparison
{
    public class FileSystemItem
    {
        public FileSystemItem(FileInfo file)
        {
            this.Name = file.Name;
            this.FullName = file.FullName;
            this.Size = file.Length;
            this.CreationTime = file.CreationTime;
            this.LastAccessTime = file.LastAccessTime;
            this.LastWriteTime = file.LastWriteTime;
            this.IsFolder = false;
        }

        public FileSystemItem(DirectoryInfo folder)
        {
            this.Name = folder.Name;
            this.FullName = folder.FullName;
            this.Size = null;
            this.CreationTime = folder.CreationTime;
            this.LastAccessTime = folder.LastAccessTime;
            this.LastWriteTime = folder.LastWriteTime;
            this.IsFolder = true;
        }

        public string Name { get; set; }
        public string FullName { get; set; }
        public long? Size { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime LastAccessTime { get; set; }
        public DateTime LastWriteTime { get; set; }
        public bool IsFolder { get; set; }

        public string FileSystemType
        {
            get
            {
                if (this.IsFolder)
                    return "File folder";
                else
                {
                    var extension = Path.GetExtension(this.Name);

                    if (IsMatch(extension, ".txt"))
                        return "Text file";
                    else if (IsMatch(extension, ".pdf"))
                        return "PDF file";
                    else if (IsMatch(extension, ".doc", ".docx"))
                        return "Microsoft Word document";
                    else if (IsMatch(extension, ".xls", ".xlsx"))
                        return "Microsoft Excel document";
                    else if (IsMatch(extension, ".jpg", ".jpeg"))
                        return "JPEG image file";
                    else if (IsMatch(extension, ".gif"))
                        return "GIF image file";
                    else if (IsMatch(extension, ".png"))
                        return "PNG image file";


                    // If we reach here, return the name of the extension
                    if (string.IsNullOrEmpty(extension))
                        return "Unknown file type";
                    else
                        return extension.Substring(1).ToUpper() + " file";
                }
            }
        }

        private bool IsMatch(string extension, params string[] extensionsToCheck)
        {
            foreach (var str in extensionsToCheck)
                if (string.CompareOrdinal(extension, str) == 0)
                    return true;

            // If we reach here, no match
            return false;
        }
    }
}