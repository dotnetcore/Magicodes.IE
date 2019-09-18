using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    public class TemplateFileInfo
    {
        public string FileName
        {
            get;
            set;
        }

        public string FileType
        {
            get;
            set;
        }

        public TemplateFileInfo()
        {

        }

        public TemplateFileInfo(string fileName, string fileType)
        {
            this.FileName = fileName;
            this.FileType = fileType;
        }
    }
}
