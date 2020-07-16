using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;

namespace PSWikiTable
{
    [Cmdlet(verbName: "ConvertTo", nounName: "WikiTable", ConfirmImpact = ConfirmImpact.None)]
    public partial class ConvertToWikiTableCmdlet : Cmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Alias("FullName")]
        public string Path { get; set; }

        [Parameter(Mandatory = false)]
        public string Worksheet { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter NoFormatting { get; set; }

        [Parameter(Mandatory = false)]
        public Uri WikiBaseUri { get; set; }

        [Parameter(Mandatory = false)]
        public Hashtable Templates { get; set; }

        protected override void ProcessRecord()
        {
            if (!File.Exists(Path))
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new FileNotFoundException($"Cannot find path {Path} because it does not exist."),
                    errorId: "PathNotFound",
                    ErrorCategory.ObjectNotFound,
                    targetObject: Path
                ));
            }
            string extension = System.IO.Path.GetExtension(Path);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new ArgumentException($"Wrong file type. Accepted file types are .xlsx and .xlsm."),
                    errorId: "WrongFileType",
                    ErrorCategory.InvalidArgument,
                    targetObject: Path
                ));
            }
            Dictionary<string, string> templateDictionary = null;
            if (Templates != null && Templates.Count > 0)
            {
                templateDictionary = Templates.Cast<DictionaryEntry>()
                    .ToDictionary(de => (string)de.Key, de => (string)de.Value, StringComparer.CurrentCultureIgnoreCase);
            }

            FileInfo file = new FileInfo(Path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet sheet;
                if (string.IsNullOrEmpty(Worksheet))
                {
                    sheet = package.Workbook.Worksheets[0];
                }
                else
                {
                    sheet = package.Workbook.Worksheets[Worksheet];
                }
                int columnCount = GetTableWidth(sheet);
                int rowCount = GetTableHeight(sheet, columnCount);
                TableSettings settings = new TableSettings()
                {
                    Width = columnCount,
                    Height = rowCount,
                    NoFormatting = NoFormatting,
                    WikiBaseUri = WikiBaseUri,
                    Templates = templateDictionary
                };
                WriteObject(BuildWikiTable(sheet, settings));
            }
        }
    }
}
