using ExcelReport.Renderers;
using System;
using System.IO;

namespace ExcelReport
{
    /// <summary>
    /// 
    /// </summary>
    /// <remarks>NPOI暂不建议使用2.5.0+版本,该版本导出Excel插入行会导致后面的合并行丢失,目前已验证2.4.1正常</remarks>
    public static class ExportHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <remarks>NPOI暂不建议使用2.5.0+版本,该版本导出Excel插入行会导致后面的合并行丢失,目前已验证2.4.1正常</remarks>
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="sheetRenderers"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        public static void ExportToLocal(string templateFile, string targetFile, params SheetRenderer[] sheetRenderers)
        {
            if (string.IsNullOrWhiteSpace(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }
            var dir = Path.GetDirectoryName(targetFile);
            if (!Directory.Exists(dir))
            { Directory.CreateDirectory(dir); }
            var buffer = ExportToBytes(templateFile, sheetRenderers);
            using var fs = File.OpenWrite(targetFile);
            fs.Write(buffer, 0, buffer.Length);
            fs.Flush();
        }
        public static byte[] ExportToBytes(string templateFile, params SheetRenderer[] sheetRenderers)
        {
            if (string.IsNullOrWhiteSpace(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException("template file not exist");
            }
            return Export.ExportToBuffer(templateFile, sheetRenderers);
        }
    }
}
