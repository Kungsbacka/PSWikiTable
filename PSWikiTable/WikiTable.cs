using OfficeOpenXml;
using System;
using System.Management.Automation;
using System.Text;
using System.Web;

namespace PSWikiTable
{
    public partial class ConvertToWikiTableCmdlet
    {
        private int GetTableWidth(ExcelWorksheet sheet)
        {
            // Detect table width
            int columnCount = 0;
            for (; sheet.Cells[1, columnCount + 1].Value != null && columnCount < 1000; columnCount++)
            {
                ;
            }

            if (columnCount == 0)
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new InvalidOperationException($"Could not determine table width. Width is too small (0 columns)."),
                    errorId: "ExcelTableError",
                    ErrorCategory.InvalidData,
                    targetObject: sheet
                ));
            }
            if (columnCount >= 1000)
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new InvalidOperationException($"Could not determine table width. Width is too big (>1000 columns)."),
                    errorId: "ExcelTableError",
                    ErrorCategory.InvalidData,
                    targetObject: sheet
                ));
            }
            return columnCount;
        }

        private int GetTableHeight(ExcelWorksheet sheet, int columnCount)
        {
            // Detect table height
            int rowCount = 0;
            bool haveValues;
            do
            {
                haveValues = false;
                for (int i = 0; i < columnCount; i++)
                {
                    if (sheet.Cells[rowCount + 1, i + 1].Value != null)
                    {
                        haveValues = true;
                        rowCount++;
                        break;
                    }
                }
            }
            while (haveValues && rowCount < 10000);

            if (rowCount == 0)
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new InvalidOperationException($"Could not determine table height. Height is too small (0 rows)."),
                    errorId: "ExcelTableError",
                    ErrorCategory.InvalidData,
                    targetObject: Path
                ));
            }
            if (haveValues)
            {
                ThrowTerminatingError(new ErrorRecord(
                    exception: new InvalidOperationException($"Could not determine table height. Height is too big (>10000 rows)."),
                    errorId: "ExcelTableError",
                    ErrorCategory.InvalidData,
                    targetObject: Path
                ));
            }
            return rowCount;
        }

        private bool IsWikiLink(Uri link, Uri baseUri)
        {
            return link.AbsolutePath.IEquals(baseUri.AbsolutePath)
                && link.Host.IEquals(baseUri.Host);

        }

        private string GetWikiTitle(Uri link)
        {
            if (string.IsNullOrEmpty(link.Query) && link.Segments.Length > 0)
            {
                return link.Segments[link.Segments.Length - 1];
            }
            return HttpUtility.ParseQueryString(link.Query).Get("title");
        }

        private string BuildWikiTable(ExcelWorksheet sheet, TableSettings settings)
        {
            StringBuilder wikiTable = new StringBuilder();
            wikiTable.AppendLine("{|class=\"wikitable sortable\"");

            // Heading
            for (int i = 0; i < settings.Width; i++)
            {
                wikiTable.Append("!");
                AppendWikiTableColumn(sheet.Cells[1, i + 1], wikiTable, settings, isHeading: true);
                wikiTable.AppendLine();
            }

            // Data
            for (int i = 1; i < settings.Height; i++)
            {
                wikiTable.AppendLine("|-");
                for (int j = 0; j < settings.Width; j++)
                {
                    wikiTable.Append("|");
                    AppendWikiTableColumn(sheet.Cells[i + 1, j + 1], wikiTable, settings);
                    wikiTable.AppendLine();
                }
            }
            wikiTable.AppendLine("|}");
            return wikiTable.ToString();
        }

        private void AppendWikiTableColumn(ExcelRange cell, StringBuilder wikiTable, TableSettings settings, bool isHeading = false)
        {
            StringBuilder data = new StringBuilder();
            StringBuilder style = new StringBuilder();
            bool isTemplated = false;
            if (cell.Value == null)
            {
                data.Append("&nbsp;");
            }
            else
            {
                // Content
                string cellValue = (string)cell.Value;
                if (cell.Hyperlink != null)
                {
                    string urlEncodedValue = HttpUtility.UrlEncode(cellValue);
                    // Internal Wiki link?
                    if (settings.WikiBaseUri != null && IsWikiLink(cell.Hyperlink, settings.WikiBaseUri))
                    {
                        string title = GetWikiTitle(cell.Hyperlink);
                        if (title.IEquals(cellValue) || title.IEquals(urlEncodedValue))
                        {
                            data.Append("[[");
                            data.Append(title);
                            data.Append("]]");
                        }
                        else
                        {
                            data.Append("[[");
                            data.Append(title);
                            data.Append("|");
                            data.Append(cellValue);
                            data.Append("]]");
                        }
                    }
                    else // External link
                    {
                        if (cell.Hyperlink.OriginalString.IEquals(cellValue) || cell.Hyperlink.OriginalString.IEquals(urlEncodedValue))
                        {
                            data.Append("[");
                            data.Append(cell.Hyperlink.ToString());
                            data.Append("]");
                        }
                        else
                        {
                            data.Append("[");
                            data.Append(cell.Hyperlink.ToString());
                            data.Append(" ");
                            data.Append(cellValue);
                            data.Append("]");
                        }
                    }
                }
                else
                {
                    if (settings.Templates != null && settings.Templates.TryGetValue(cellValue, out string template))
                    {
                        data.Append("{{");
                        data.Append(template);
                        data.Append("}}");
                        isTemplated = true;
                    }
                    else
                    {
                        data.Append(cellValue);
                    }
                }

                // Line breaks
                data.Replace("\n", "<br />");

                // Font style
                if (cell.Style.Font.Italic)
                {
                    data.Insert(0, "''");
                    data.Append("''");
                }
                if (cell.Style.Font.Bold && !isHeading)
                {
                    data.Insert(0, "'''");
                    data.Append("'''");
                }
                if (cell.Style.Font.Strike)
                {
                    data.Insert(0, "<s>");
                    data.Append("</s>");
                }
                if (!isTemplated && !settings.NoFormatting)
                {
                    if (!isHeading && cell.Style.HorizontalAlignment == OfficeOpenXml.Style.ExcelHorizontalAlignment.Center)
                    {
                        style.Append("text-align:center;");
                    }
                    else if (cell.Style.HorizontalAlignment == OfficeOpenXml.Style.ExcelHorizontalAlignment.Right)
                    {
                        style.Append("text-align:right;");
                    }
                    else if (isHeading && cell.Style.HorizontalAlignment == OfficeOpenXml.Style.ExcelHorizontalAlignment.Left)
                    {
                        style.Append("text-align:left;");
                    }
                    if (cell.Style.Font.Size != 11)
                    {
                        style.Append("font-size:");
                        style.Append((int)cell.Style.Font.Size);
                        style.Append("px;");
                    }
                }
            }
            if (!isTemplated && !settings.NoFormatting)
            {
                if (!cell.Style.Fill.BackgroundColor.Auto)
                {
                    string color = null;
                    if (!string.IsNullOrEmpty(cell.Style.Fill.BackgroundColor.Rgb))
                    {
                        color = cell.Style.Fill.BackgroundColor.Rgb;
                    }
                    else if (cell.Style.Fill.BackgroundColor.Theme != null)
                    {
                        color = ColorTheme.GetThemeRgb((OfficeOpenXml.Drawing.eThemeSchemeColor)cell.Style.Fill.BackgroundColor.Theme, cell.Style.Fill.BackgroundColor.Tint);
                    }
                    if (!string.IsNullOrEmpty(color) && !color.IEquals("#FFFFFF")) // White is wiki default
                    {
                        style.Append("background-color:");
                        style.Append(color);
                        style.Append(";");
                    }
                }
                if (!cell.Style.Font.Color.Auto)
                {
                    string color = null;
                    if (!string.IsNullOrEmpty(cell.Style.Font.Color.Rgb))
                    {
                        color = cell.Style.Font.Color.Rgb;
                    }
                    else if (cell.Style.Font.Color.Theme != null)
                    {
                        color = ColorTheme.GetThemeRgb((OfficeOpenXml.Drawing.eThemeSchemeColor)cell.Style.Font.Color.Theme, cell.Style.Font.Color.Tint);
                    }
                    if (!string.IsNullOrEmpty(color) && !color.IEquals("#000000")) // Black is wiki default
                    {
                        style.Append("color:");
                        style.Append(color);
                        style.Append(";");
                    }
                }
            }
            if (style.Length > 0)
            {
                wikiTable.Append("style=\"");
                wikiTable.Append(style);
                wikiTable.Append("\" | ");
            }
            wikiTable.Append(data);
        }
    }
}
