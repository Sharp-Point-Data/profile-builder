using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ProfileBuilder.System.Data
{
    class Environment
    {
        public static DirectoryInfo ExcelPath
        {
            get
            {
                return new DirectoryInfo(Properties.Settings.Default.ExcelPath);
            }
        }

        public static DirectoryInfo Font
        {
            get
            {
                return new DirectoryInfo(Properties.Settings.Default.Font);
            }
        }

        public static DirectoryInfo InputFolder
        {
            get
            {
                return new DirectoryInfo(Properties.Settings.Default.InputFolder);
            }
        }

        public static DirectoryInfo OutputFolder
        {
            get
            {
                return new DirectoryInfo(Properties.Settings.Default.OutputFolder);
            }
        }
    }
}
