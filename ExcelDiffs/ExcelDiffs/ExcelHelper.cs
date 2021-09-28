using ExcelDiffs;
using ExcelDiffs.Models;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ExcelDiffs
{
    public static class ExcelHelper
    {
        public static MemoryStream GenerateItemsExcel<T>(List<T> items, string worksheetName,
            Tuple<int, int> dataStartCell)
        {
            ExcelPackage package = new ExcelPackage();
            package.GenerateItemsWorksheet(items, worksheetName, dataStartCell);
            //package.SaveAs(new FileInfo("test.xlsx"));
            MemoryStream result = new MemoryStream(package.GetAsByteArray());
            package.Dispose();
            return result;
        }

        public static void SetError(this IExcelDataValidation dataValidation, bool allowBlank = true)
        {
            dataValidation.ShowInputMessage = true;
            dataValidation.Prompt = "Error Message";
            dataValidation.ShowErrorMessage = true;
            dataValidation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            dataValidation.ErrorTitle = "Error Message";
            dataValidation.Error = "Error Message";
            dataValidation.AllowBlank = allowBlank;
        }

        //public static ExcelByteArrayResponse ToByteArrayResponse(this MemoryStream stream)
        //{
        //    return new ExcelByteArrayResponse
        //    {
        //        ByteArray = Base64.ToBase64String(stream.ToArray()),
        //        MimeType = MimeTypes.GetMimeTypeFromExtension(".xlsx"),
        //    };
        //}

        // ReSharper disable once UnusedMember.Local
        private static List<int> GetBooleanTypesIndexes<T>()
        {
            List<int> result = new List<int>();
            PropertyInfo[] properties = typeof(T).GetProperties();

            for (int i = 0; i < properties.Count(); i++)
                if (properties[i].PropertyType == typeof(bool) || properties[i].PropertyType == typeof(bool?))
                    result.Add(i);

            return result;
        }

        public static ExcelWorksheet GenerateItemsWorksheet<T>(this ExcelPackage package,
            List<T> items, string name,
            Tuple<int, int> dataStartCell,
            List<string> hiddenColumns = null)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(name);
            worksheet.PopulateWorksheet(items, dataStartCell, hiddenColumns);
            return worksheet;
        }

        public static void PopulateWorksheet<T>(this ExcelWorksheet worksheet, List<T> items,
            Tuple<int, int> dataStartCell, List<string> hiddenColumns = null)
        {
            var (rowStart, columnStart) = dataStartCell;

            MemberInfo[] membersToInclude = typeof(T)
                .GetProperties()
                .OrderBy(p => p.MetadataToken)
                .Where(p => !hiddenColumns?.Contains(p.Name) ?? true)
                .ToArray();

            ExcelRangeBase headerRange = worksheet.Cells[rowStart, columnStart].LoadFromArrays(new List<string[]>(new[]
            {
                membersToInclude.Select(property => Regex.Replace(property.Name, "(\\B[A-Z])", " $1"))
                    .ToArray()
            }));

            headerRange.Style.Font.Bold = true;

            if (items != null && items.Count > 0)
                worksheet.Cells[rowStart + 1, columnStart]
                    .LoadFromCollection(items, false,
                        TableStyles.None,
                        BindingFlags.Public | BindingFlags.Instance,
                        membersToInclude);

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        public static List<string> GetColumnDataFromExcelFile(IFormFile file, int worksheetIndex,
            Tuple<int, int> columnDataStartCell)
        {
            List<string> result = new List<string>();

            Stream stream = file.OpenReadStream();
            using ExcelPackage package = new ExcelPackage(stream);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

            if (columnDataStartCell.Item1 == worksheet.Dimension.End.Row)
            {
                string singleValue = worksheet.Cells[columnDataStartCell.Item1, columnDataStartCell.Item2,
                    worksheet.Dimension.End.Row, columnDataStartCell.Item2].Value.ToString();
                result.Add(singleValue);
            }
            else
            {
                Array data = worksheet.Cells[columnDataStartCell.Item1, columnDataStartCell.Item2,
                    worksheet.Dimension.End.Row, columnDataStartCell.Item2].Value as Array;

                if (data == null) return result.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();

                for (int i = 0; i < data.Length; i++)
                    result.Add(data.GetValue(i, columnDataStartCell.Item2 - 1)?.ToString());
            }

            return result.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        }

        public static List<T> GetDataFromExcelFile<T>(IFormFile file, int worksheetIndex, Tuple<int, int> dataStartCell,
            Tuple<int, int> dataEndCell = null, List<int> hiddenColumns = null)
            where T : class, new()
        {
            var result = new List<T>();
            Stream stream = file.OpenReadStream();
            using ExcelPackage package = new ExcelPackage(stream);

            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];
            var (lastRow, lastCol) = new Tuple<int, int>(
                dataEndCell?.Item1 ?? worksheet.Dimension.End.Row,
                dataEndCell?.Item2 ?? worksheet.Dimension.End.Column);

            var properties = typeof(T).GetProperties().OrderBy(p => p.MetadataToken).ToArray();

            if (hiddenColumns != null)
            {
                Array headersArray =
                    worksheet.Cells[dataStartCell.Item1 - 1, dataStartCell.Item2, dataStartCell.Item1 - 1, lastCol]
                        .Value as Array;

                var headersList = new List<string>();

                for (int j = 0; j < lastCol - dataStartCell.Item2 + 1; j++)
                    headersList.Add(headersArray?.GetValue(0, j)?.ToString());

                for (int i = 0; i < properties.Length; i++)
                    if (hiddenColumns.Contains(i) &&
                        headersList.Contains(Regex.Replace(properties[i].Name, "(\\B[A-Z])", " $1")) ||
                        !hiddenColumns.Contains(i) &&
                        !headersList.Contains(Regex.Replace(properties[i].Name, "(\\B[A-Z])", " $1")))
                        return null;
            }

            Array data = worksheet.Cells[dataStartCell.Item1, dataStartCell.Item2, lastRow, lastCol].Value as Array;

            for (int i = 0; i < lastRow - dataStartCell.Item1 + 1; i++)
            {
                T obj = new T();

                for (int j = 0; j < lastCol - dataStartCell.Item2 + 1; j++)
                {
                    if (j >= properties.Length)
                        continue;

                    try
                    {
                        if (Nullable.GetUnderlyingType(properties[j].PropertyType) != null)
                            properties[j].SetValue(obj,
                                Convert.ChangeType(data?.GetValue(i, j),
                                    Nullable.GetUnderlyingType(properties[j].PropertyType)!));
                        else
                            properties[j].SetValue(obj,
                                Convert.ChangeType(data?.GetValue(i, j),
                                    properties[j].PropertyType));
                    }
                    catch (Exception)
                    {
                        // ignored
                    }
                }

                result.Add(obj);
            }

            return result;
        }

        public static List<List<string>> GetAllExcelData(IFormFile file, int worksheetIndex)
        {
            using ExcelPackage package = new ExcelPackage(file.OpenReadStream());
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

            ExcelCellAddress startCell = worksheet.Dimension.Start;
            ExcelCellAddress endCell = worksheet.Dimension.End;

            Array data = worksheet.Cells[startCell.Row, startCell.Column, endCell.Row, endCell.Column].Value as Array;
            var result = new List<List<string>>();

            for (int i = startCell.Row - 1; i < endCell.Row; i++)
            {
                result.Add(new List<string>());
                result[i] = new List<string>();

                for (int j = startCell.Column - 1; j < endCell.Column; j++)
                    result[i].Add(data?.GetValue(i, j)?.ToString() ?? "");
            }

            return result;
        }

        public static ExcelPackage UpdateExcelColumn(IFormFile file,
            List<string> items, int worksheetIndex, Tuple<int, int> column)
        {
            Stream stream = file.OpenReadStream();
            ExcelPackage package = new ExcelPackage(stream);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

            var itemsFormatted = items
                .Select(p => p.Split('.'))
                .Select(p => string.Join(".\n", p))
                .ToList();

            ExcelRange cells =
                worksheet.Cells[column.Item1, column.Item2, itemsFormatted.Count + column.Item1 - 1, column.Item2];
            cells.LoadFromCollection(itemsFormatted);
            cells.Style.Font.Color.SetColor(Color.Red);
            cells.AutoFitColumns();

            for (var i = column.Item1; i < itemsFormatted.Count + column.Item1; i++)
                worksheet.Row(i).Height =
                    worksheet.DefaultRowHeight * itemsFormatted[i - column.Item1].Split('.').Length;

            cells.Style.WrapText = true;
            return package;
        }

        public static void ColorCells(this ExcelWorksheet worksheet, Color color, string cells)
        //usage excel.colorCells(0, Color.Ceva, "A1:A5,A8,A10,H8")
        {
            worksheet.Cells[cells].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[cells].Style.Fill.BackgroundColor.SetColor(color);
        }

        public static void FreezeCells(this ExcelWorksheet worksheet, int rows, int columns)
        {
            worksheet.View.FreezePanes(rows, columns);
        }

        public static void AddFullBorder(this ExcelWorksheet worksheet, Color borderColor, string cells)
        {
            worksheet.Cells[cells].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[cells].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[cells].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[cells].Style.Border.Right.Style = ExcelBorderStyle.Thick;

            worksheet.Cells[cells].Style.Border.Top.Color.SetColor(borderColor);
            worksheet.Cells[cells].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[cells].Style.Border.Left.Color.SetColor(borderColor);
            worksheet.Cells[cells].Style.Border.Right.Color.SetColor(borderColor);
        }

        public static void BuildBorders(this ExcelWorksheet worksheet,
            string leftTopCell, string rightTopCell, string rightBottomCell, string leftBottomCell,
            Color borderColor)
        {
            //top
            var top = leftTopCell + ":" + rightTopCell;
            worksheet.Cells[top].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[top].Style.Border.Top.Color.SetColor(borderColor);
            Console.WriteLine("top ok");

            //right
            var right = rightTopCell + ":" + rightBottomCell;
            worksheet.Cells[right].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[right].Style.Border.Right.Color.SetColor(borderColor);
            Console.WriteLine("right ok");
            //bottom
            var bottom = leftBottomCell + ":" + rightBottomCell;
            worksheet.Cells[bottom].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[bottom].Style.Border.Bottom.Color.SetColor(borderColor);
            Console.WriteLine("bottom ok");
            //left
            var left = leftTopCell + ":" + leftBottomCell;
            worksheet.Cells[left].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[left].Style.Border.Left.Color.SetColor(borderColor);
            Console.WriteLine("left ok");
        }

        public static void BoldCells(this ExcelWorksheet worksheet, string cells)
        {
            cells = cells.Replace(" ", "");
            worksheet.Cells[cells].Style.Font.Bold = true;
        }

        public static void ItalicCells(this ExcelWorksheet worksheet, string cells)
        {
            var excelCells = cells.Replace(" ", "")
                .Split(",").ToList();
            foreach (var cell in excelCells)
                worksheet.Cells[cell].Style.Font.Italic = true;

        }

        public static void SetFontInCells(this ExcelWorksheet worksheet, string cells,
            string fontName, float fontSize = 11)
        {
            worksheet.Cells[cells].Style.Font.Name = fontName;
            worksheet.Cells[cells].Style.Font.Size = fontSize;
        }

        public static void SetNumberFormatInCalls(this ExcelWorksheet worksheet, List<(string, string)> values)
        {
            foreach (var (cell, value) in values)
                worksheet.Cells[cell].Style.Numberformat.Format = value;
        }

        //validation
        public static void ValidateStringLengthForCells(this ExcelWorksheet worksheet,
            List<(string, int)> values, bool allowBlank = true)
        {

            foreach (var (cells, value) in values)
            {
                var cell = cells.Replace(",", " ");
                var validation = worksheet.DataValidations.AddTextLengthValidation(cell);
                validation.SetCustomError(Errors.InvalidLength);
                validation.AllowBlank = allowBlank;
                validation.Formula.Value = value;
                validation.Formula2.Value = value;
            }
            worksheet.Calculate();
        }


        public static void CheckIntervalForCells(this ExcelWorksheet worksheet, string cells,
            int? left = null, int? right = null, bool allowBlank = true)
        {
            cells = cells.Replace(",", " ");
            var validation = worksheet.DataValidations.AddDecimalValidation(cells);
            validation.SetCustomError(Errors.NotInInterval);
            validation.Operator = ExcelDataValidationOperator.between;
            validation.AllowBlank = allowBlank;
            validation.Formula.Value = left ?? int.MinValue;
            validation.Formula2.Value = right ?? int.MaxValue;
        }

        public static void CheckForOnlyCertainValuesInCells(this ExcelWorksheet worksheet,
            string cells, string values, bool allowBlank = true)
        {
            var stringValues = values.Replace(" ", "").Split(",")
                .ToList();
            cells = cells.Replace(",", " ");
            var validation = worksheet.DataValidations.AddListValidation(cells);
            validation.SetCustomError(Errors.ValueNotAllowed);
            validation.AllowBlank = allowBlank;

            stringValues.ForEach(value =>
                validation.Formula.Values.Insert(validation.Formula.Values.Count, value));

        }

        public static void AddSortingAndFiltering(this ExcelWorksheet worksheet, int worksheetIndex,
            string cells)
        {
            worksheet.Cells[cells].AutoFilter = true;
            worksheet.Cells[cells].AutoFitColumns();
        }

        public static void SetCustomError(this IExcelDataValidation dataValidation, CustomError error,
            bool allowBlank = true)
        {
            dataValidation.ShowInputMessage = true;
            dataValidation.Prompt = error.Title;
            dataValidation.ShowErrorMessage = true;
            dataValidation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            dataValidation.ErrorTitle = error.Title;
            dataValidation.Error = error.Body;
            dataValidation.AllowBlank = allowBlank;
        }

        public static void CheckCellsForOnlyNumbers(this ExcelWorksheet worksheet,
            string cells, bool allowBlank = true)
        {
            var validation = worksheet.DataValidations.AddCustomValidation(cells);
            validation.SetCustomError(Errors.OnlyNumber);
            validation.AllowBlank = allowBlank;
            validation.Formula.ExcelFormula = $"ISNUMBER({cells})";
            validation.Operator = ExcelDataValidationOperator.equal;
        }

        public static void CollapseCells(this ExcelWorksheet worksheet, string cells)
        {
            worksheet.Cells[cells].Merge = true;
        }

        public static void WriteInCells(this ExcelWorksheet worksheet, List<(string, string)> values)
        {
            foreach (var (cell, value) in values)
                worksheet.Cells[cell].Value = value;
        }

        public static void AddPagination(this ExcelWorksheet worksheet, int breakAfter = 1000)
        {
            worksheet.Row(breakAfter).PageBreak = true;
        }

        public static void SetDecimalType(this ExcelWorksheet worksheet,
            string cells)
        {
            cells = cells.Replace(" ", "");
            var currentCells = worksheet.Cells[cells];
            foreach (var cell in currentCells)
            {
                if (cell.Value == null)
                    continue;
                if (cell.Value.ToString() != "")
                    cell.Value = Convert.ToDecimal(cell.Value);
            }
        }

        public static void StyleHeaders(this ExcelWorksheet worksheet, int row, int height)
        {
            worksheet.Row(row).Height = height;
            worksheet.Row(row).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(row).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }

}