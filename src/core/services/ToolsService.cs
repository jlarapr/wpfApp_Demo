namespace wpfApp_Demo.src.core.services {
    using OfficeOpenXml.Style;
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Threading.Tasks;
    using wpfApp_Demo.src.core.interfaces.services;

    internal class ToolsService : IToolsService {

        public void TotalesToExcel(string excelFileName, string titel, string applicationName, string myDate,
           DataTable table) {
            try {
                if (File.Exists(excelFileName))
                    File.Delete(excelFileName);

                string worksheetsName = applicationName + " " + myDate;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName));
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                    worksheet.Cells["A:XFD"].Style.Font.Bold = true;

                    ExcelRange rng = worksheet.Cells[8, 2, 8, 2];
                    {

                        worksheet.Column(1).Width = 19.00;
                        worksheet.Column(2).Width = 19.00;
                        worksheet.Column(3).Width = 19.00;
                        worksheet.Column(4).Width = 19.00;

                        worksheet.Cells[8, 2, 8, 4].Merge = true;
                        worksheet.Cells[8, 2, 8, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        rng.Value = titel.ToUpper();
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;

                        rng = worksheet.Cells[9, 2, 9, 2];
                        worksheet.Cells[9, 2, 9, 4].Merge = true;
                        worksheet.Cells[9, 2, 9, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        rng.Value = applicationName.ToUpper();
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;

                        rng = worksheet.Cells[10, 2, 10, 2];
                        worksheet.Cells[10, 2, 10, 4].Merge = true;
                        worksheet.Cells[10, 2, 10, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                        myDate = DateTime.TryParseExact(myDate, "MMddyy", CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out DateTime dBht)
                            ? dBht.ToString("MM/dd/yyyy")
                            : DateTime.Now.ToString("MM/dd/yyyy");

                        rng.Value = "FECHA RECLAMACION : " + myDate.ToUpper();
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;

                        rng = worksheet.Cells[12, 2, 12, 2];
                        worksheet.Cells[12, 2, 12, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        rng.Value = "Año";
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                        rng = worksheet.Cells[12, 3, 12, 3];
                        worksheet.Cells[12, 3, 12, 3].Style.HorizontalAlignment =
                            ExcelHorizontalAlignment.Center;
                        rng.Value = "Total Imagenes";
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                        rng = worksheet.Cells[12, 4, 12, 4];
                        worksheet.Cells[12, 4, 12, 4].Style.HorizontalAlignment =
                            ExcelHorizontalAlignment.Center;
                        rng.Value = "Total Reclamaciones";
                        rng.Style.Font.Size = 11;
                        rng.Style.Font.Bold = true;
                        rng.Style.Font.Italic = false;
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                        int r = 13;
                        int c = 2;

                        foreach (DataRow dataRow in table.Rows) {
                            for (int i = 0; i < 3; i++) {
                                rng = worksheet.Cells[r, c, r, c];
                                rng.Style.Font.Bold = false;

                                rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                if (i > 0)
                                    rng.Style.Numberformat.Format = "#,###,##0";

                                worksheet.Cells[r, c, r, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                rng.Value = dataRow[i];
                                c++;
                            }

                            r++;
                            c = 2;
                        }

                        for (int ic = 2; ic <= 4; ic++) {
                            rng = worksheet.Cells[r + 1, ic];
                            rng.Style.Font.Bold = true;

                            if (ic == 2)
                                rng.Value = "Totales:";

                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        }

                        worksheet.Cells[r + 1, 3].Style.Numberformat.Format = "#,###,##0";
                        worksheet.Cells[r + 1, 3].Formula = "=SUM(" + worksheet.Cells[13, 3].Address + ":" +
                                                            worksheet.Cells[r - 1, 3].Address + ")";
                        worksheet.Cells[r + 1, 4].Style.Numberformat.Format = "#,###,##0";
                        worksheet.Cells[r + 1, 4].Formula = "=SUM(" + worksheet.Cells[13, 4].Address + ":" +
                                                            worksheet.Cells[r - 1, 4].Address + ")";
                    }
                    worksheet.Protection.IsProtected = false;
                    worksheet.Protection.AllowSelectLockedCells = false;
                    package.Save();
                }
                GC.Collect();
            } catch (Exception ex) {
                MethodBase? site = ex.TargetSite;
                string?[] msg = { ex.Message, " ", site?.Name };
                throw new Exception(string.Concat(msg));
            }
        }

    }
}
