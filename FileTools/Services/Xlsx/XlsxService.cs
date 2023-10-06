using System.Data;
using System.Drawing;
using System.Reflection;
using FileTools.Extensions;
using FileTools.Models.Xlsx;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FileTools.Services.Xlsx
{
    public class XlsxService
    {
        private List<ExcelWorksheet> _worksheets;
        private readonly string _fullpath;
        public XlsxService(string fullpath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _worksheets = new List<ExcelWorksheet>();
            _fullpath = fullpath;
        }
        #region Write
        public void AddWorksheet<T>(string nameWorksheet, T obj, bool printHeaders = true)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica un nombre para la hoja de trabajo a agregar en el archivo '{_fullpath}'.");

            if (WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' ya existe en el archivo '{_fullpath}'.");

            if (obj == null)
                throw new ArgumentNullException(nameof(obj), $"El objeto tipo '{typeof(T)}' es null y el archivo '{_fullpath}' no puede ser procesado.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            try
            {
                ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add(nameWorksheet);
                ExcelRangeBase range = worksheet.Cells[1, 1].LoadFromCollection(new List<T> { obj }, printHeaders);

                worksheet.Cells.AutoFitColumns();

                _worksheets.Add(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void AddWorksheet<T>(string nameWorksheet, List<CustomHeader> customHeaders, T obj)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica un nombre para la hoja de trabajo a agregar en el archivo '{_fullpath}'.");

            if (customHeaders.IsNullOrEmpty<CustomHeader>())
                throw new ArgumentNullException(nameof(customHeaders), $"La lista tipo '{typeof(List<CustomHeader>)}' es null o está vacía y no puede ser procesado el archivo '{_fullpath}'.");

            if (customHeaders.HasNullItems<CustomHeader>())
                throw new Exception($"Un objeto de tipo '{typeof(CustomHeader)}' es null en la lista de tipo '{typeof(List<CustomHeader>)}' y el archivo '{_fullpath}' no puede ser procesado .");

            if (obj == null)
                throw new ArgumentNullException(nameof(obj), $"El objeto tipo '{typeof(T)}' es null y el archivo '{_fullpath}' no puede ser procesado.");

            if (WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' ya existe en el archivo '{_fullpath}'.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            try
            {
                ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add(nameWorksheet);

                for (int position = 0; position < customHeaders.Count; position++)
                    worksheet.Cells[1, position + 1].Value = customHeaders[position].HeaderName;

                for (int column = 0; column < customHeaders.Count; column++)
                {
                    CustomHeader customHeader = customHeaders[column];
                    var propertyValue = obj?.GetType()?.GetProperty(customHeader.PropertyName)?.GetValue(obj);

                    if (customHeader.CustomStyle != null)
                        worksheet = SetCustomStyle(worksheet, customHeader.CustomStyle, 2, column + 1);

                    worksheet.Cells[2, column + 1].Value = propertyValue;
                }

                worksheet.Cells.AutoFitColumns();

                _worksheets.Add(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void AddWorksheet(string nameWorksheet, DataTable data)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica un nombre para la hoja de trabajo a agregar en el archivo '{_fullpath}'.");

            if (data.IsNullOrEmpty())
                throw new ArgumentNullException(nameof(data), $"El objeto tipo {typeof(DataTable)} '{nameof(data)}' es null o está vacío y no puede ser procesado en '{_fullpath}'.");

            if (WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' ya existe en el archivo '{_fullpath}'.");

            try
            {
                ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add(nameWorksheet);
                ExcelRangeBase range = worksheet.Cells[1, 1].LoadFromDataTable(data, true);

                worksheet.Cells.AutoFitColumns();

                _worksheets.Add(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void AddWorksheet<T>(string nameWorksheet, List<T> data, bool printHeaders = true)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica un nombre para la hoja de trabajo a agregar en el archivo '{_fullpath}'.");

            if (data.IsNullOrEmpty<T>())
                throw new ArgumentNullException(nameof(data), $"La lista tipo '{typeof(List<T>)}' es null o está vacía y no puede ser procesada en '{_fullpath}'.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            if (WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' ya existe en el archivo '{_fullpath}'.");

            try
            {
                ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add(nameWorksheet);
                ExcelRangeBase range = worksheet.Cells[1, 1].LoadFromCollection(data, printHeaders);

                worksheet.Cells.AutoFitColumns();

                _worksheets.Add(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void AddWorksheet<T>(string nameWorksheet, List<CustomHeader> customHeaders, List<T> data)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica un nombre para la hoja de trabajo a agregar en el archivo '{_fullpath}'.");

            if (customHeaders.IsNullOrEmpty<CustomHeader>())
                throw new ArgumentNullException(nameof(customHeaders), $"La lista tipo '{typeof(List<CustomHeader>)}' es null o está vacía y no puede ser procesado el archivo '{_fullpath}'.");

            if (customHeaders.HasNullItems<CustomHeader>())
                throw new Exception($"Un objeto de tipo '{typeof(CustomHeader)}' es null en la lista de tipo '{typeof(List<CustomHeader>)}' y el archivo '{_fullpath}' no puede ser procesado .");

            if (data.IsNullOrEmpty<T>())
                throw new ArgumentNullException(nameof(data), $"La lista tipo '{typeof(List<T>)}' es null o está vacía y no puede ser procesada en '{_fullpath}'.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            if (WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' ya existe en el archivo '{_fullpath}'.");

            try
            {
                ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add(nameWorksheet);

                for (int position = 0; position < customHeaders.Count; position++)
                    worksheet.Cells[1, position + 1].Value = customHeaders[position].HeaderName;

                for (int row = 0; row < data.Count; row++)
                {
                    T dataRow = data[row];
                    for (int column = 0; column < customHeaders.Count; column++)
                    {
                        CustomHeader customHeader = customHeaders[column];
                        var propertyValue = dataRow?.GetType()?.GetProperty(customHeader.PropertyName)?.GetValue(dataRow);

                        if (customHeader.CustomStyle != null)
                            worksheet = SetCustomStyle(worksheet, customHeader.CustomStyle, row + 2, column + 1);

                        worksheet.Cells[row + 2, column + 1].Value = propertyValue;
                    }
                }

                worksheet.Cells.AutoFitColumns();

                _worksheets.Add(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void InsertRow<T>(string nameWorksheet, T obj)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica el nombre de la hoja de trabajo a la que le agregarás un nuevo renglón en el archivo '{_fullpath}'.");

            if (obj == null)
                throw new ArgumentNullException(nameof(obj), $"El objeto tipo '{typeof(T)}' es null y el archivo '{_fullpath}' no puede ser procesado.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            if (_worksheets.IsNullOrEmpty<ExcelWorksheet>())
                throw new ArgumentNullException(nameof(_worksheets), $"La lista tipo '{typeof(ExcelWorksheet)}' es null o está vacía. Asegúrate de primero agregar una hoja de trabajo al archivo '{_fullpath}'.");

            if (!WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' no existe en el archivo '{_fullpath}'.");

            try
            {
                ExcelWorksheet worksheet = _worksheets.FirstOrDefault(worksheet => worksheet.Name == nameWorksheet);

                int lastRow = worksheet.Dimension?.End.Row ?? 0;
                int newRow = lastRow + 1;

                ExcelRangeBase range = worksheet.Cells[newRow, 1].LoadFromCollection(new List<T> { obj }, false);

                worksheet.Cells.AutoFitColumns();
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void InsertRow<T>(string nameWorksheet, List<CustomHeader> customHeaders, T obj)
        {
            if (string.IsNullOrEmpty(nameWorksheet))
                throw new ArgumentException($"Especifica el nombre de la hoja de trabajo a la que le agregarás un nuevo renglón en el archivo '{_fullpath}'.");

            if (customHeaders.IsNullOrEmpty<CustomHeader>())
                throw new ArgumentNullException(nameof(customHeaders), $"La lista tipo '{typeof(List<CustomHeader>)}' es null o está vacía y no puede ser procesado el archivo '{_fullpath}'.");

            if (customHeaders.HasNullItems<CustomHeader>())
                throw new Exception($"Un objeto de tipo '{typeof(CustomHeader)}' es null en la lista de tipo '{typeof(List<CustomHeader>)}' y el archivo '{_fullpath}' no puede ser procesado .");

            if (obj == null)
                throw new ArgumentNullException(nameof(obj), $"El objeto tipo '{typeof(T)}' es null y el archivo '{_fullpath}' no puede ser procesado.");

            if (!IsSingleTypeObject(typeof(T)))
                throw new Exception($"Sólo se permite insertar objetos de tipo único y no se admiten tipos de colección para la hoja de trabajo '{nameWorksheet}' y el archivo '{_fullpath}' no puede ser procesado.");

            if (_worksheets.IsNullOrEmpty<ExcelWorksheet>())
                throw new ArgumentNullException(nameof(_worksheets), $"La lista tipo '{typeof(ExcelWorksheet)}' es null o está vacía. Asegúrate de primero agregar una hoja de trabajo al archivo '{_fullpath}'.");

            if (!WorkSheetExist(_worksheets, nameWorksheet))
                throw new Exception($"La hoja de trabajo '{nameWorksheet}' no existe en el archivo '{_fullpath}'.");

            try
            {
                ExcelWorksheet worksheet = _worksheets.FirstOrDefault(worksheet => worksheet.Name == nameWorksheet);

                int lastRow = worksheet.Dimension?.End.Row ?? 0;
                int newRow = lastRow + 1;

                for (int column = 0; column < customHeaders.Count; column++)
                {
                    CustomHeader customHeader = customHeaders[column];
                    var propertyValue = typeof(T)?.GetProperty(customHeaders[column].PropertyName)?.GetValue(obj);
                    if (customHeader.CustomStyle != null)
                        worksheet = SetCustomStyle(worksheet, customHeader.CustomStyle, newRow, column + 1);

                    worksheet.Cells[newRow, column + 1].Value = propertyValue;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void WriteXlsx()
        {
            if (_worksheets.IsNullOrEmpty<ExcelWorksheet>())
                throw new ArgumentNullException(nameof(_worksheets), $"La lista tipo '{typeof(ExcelWorksheet)}' es null o está vacía. Asegúrate de primero agregar una hoja de trabajo al archivo '{_fullpath}'.");

            try
            {
                if (_worksheets.HasNullItems<ExcelWorksheet>())
                    _worksheets.RemoveAll(worksheet => worksheet == null);

                using (ExcelPackage package = new ExcelPackage(new FileInfo(_fullpath)))
                {
                    foreach (ExcelWorksheet worksheet in _worksheets)
                        package.Workbook.Worksheets.Add(worksheet.Name, worksheet);

                    package.Save();
                }

                if (File.Exists(_fullpath))
                    Console.WriteLine($"El archivo '{_fullpath}' se ha escrito correctamente.");
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion
        #region Read
        public List<T> ReadXlsx<T>(string nameWorksheet, bool hasHeaders) where T : new()
        {
            if (!File.Exists(_fullpath))
                throw new FileNotFoundException($"El archivo '{_fullpath}' no se encontró en la ubicación especificada.");

            try
            {
                List<T> objs = new List<T>();

                using (ExcelPackage package = new ExcelPackage(new FileInfo(_fullpath)))
                {
                    List<ExcelWorksheet> worksheets = package.Workbook.Worksheets.ToList();

                    if (!WorkSheetExist(worksheets, nameWorksheet))
                        throw new Exception($"La hoja de trabajo '{nameWorksheet}' no existe en el archivo '{_fullpath}'.");

                    ExcelWorksheet worksheet = (from workSheets in worksheets where workSheets.Name.Equals(nameWorksheet) select workSheets).First();

                    PropertyInfo[] properties = typeof(T).GetProperties();

                    for (int row = hasHeaders ? 2 : 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        T obj = new T();

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            Type typeProperty = properties[col - 1].PropertyType;
                            var value = Convert.ChangeType(worksheet.Cells[row, col].Text, typeProperty);
                            properties[col - 1].SetValue(obj, value);
                        }

                        objs.Add(obj);
                    }
                }

                return objs;
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion
        private ExcelWorksheet SetCustomStyle(ExcelWorksheet worksheet, CustomStyle customStyle, int row, int column)
        {
            try
            {
                ExcelWorksheet excelWorksheet = worksheet;

                // Format
                if (customStyle.CellFormat != null)
                {
                    if (string.IsNullOrEmpty(customStyle.CellFormat?.CustomFormat))
                    {
                        switch (customStyle.CellFormat?.Format)
                        {
                            case Format.Date:
                                customStyle.CellFormat.CustomFormat = "dd/MM/yyyy";
                                break;
                            case Format.Number:
                                customStyle.CellFormat.CustomFormat = "0";
                                break;
                            case Format.Decimal:
                                customStyle.CellFormat.CustomFormat = "0.00";
                                break;
                            default:
                                break;
                        }
                    }

                    worksheet.Cells[row, column].Style.Numberformat.Format = customStyle.CellFormat.CustomFormat;
                }

                // Font Size
                worksheet.Cells[row, column].Style.Font.Size = customStyle.FontSize == 0 ? 11 : customStyle.FontSize;
                // Font Bold
                worksheet.Cells[row, column].Style.Font.Bold = customStyle.FontBold;
                // Font Italic
                worksheet.Cells[row, column].Style.Font.Italic = customStyle.FontItalic;
                // Font Color
                if (customStyle.FontColor != null)
                    worksheet.Cells[row, column].Style.Font.Color.SetColor(customStyle.FontColor.Value);
                // Font Name
                if (customStyle.FontName != null)
                    worksheet.Cells[row, column].Style.Font.Name = customStyle.FontName;
                // Pattern Type
                if (customStyle.PatternType != null)
                    worksheet.Cells[row, column].Style.Fill.PatternType = customStyle.PatternType.Value;
                // Background Color
                if (customStyle.BackgroundColor != null)
                    if (customStyle.PatternType != null && customStyle.PatternType != ExcelFillStyle.None)
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(customStyle.BackgroundColor.Value);
                    else
                        throw new ArgumentNullException(nameof(customStyle.PatternType), $"Para definir un '{typeof(Color)}' en la propiedad '{nameof(customStyle.BackgroundColor)}' de tu objeto tipo '{typeof(CustomStyle)}', especifica antes un valor tipo '{typeof(ExcelFillStyle)}' diferente de '{typeof(ExcelFillStyle)}.None' o 'null' en la propiedad '{nameof(customStyle.PatternType)}' de tu objeto tipo '{typeof(CustomStyle)}'.");
                // Border Style
                if (customStyle.BorderStyle != null)
                {
                    worksheet.Cells[row, column].Style.Border.Top.Style = customStyle.BorderStyle.Value;
                    worksheet.Cells[row, column].Style.Border.Right.Style = customStyle.BorderStyle.Value;
                    worksheet.Cells[row, column].Style.Border.Bottom.Style = customStyle.BorderStyle.Value;
                    worksheet.Cells[row, column].Style.Border.Left.Style = customStyle.BorderStyle.Value;
                }
                // Border Color
                if (customStyle.BorderColor != null)
                {
                    if (customStyle.BorderStyle != null && customStyle.BorderStyle != ExcelBorderStyle.None)
                    {
                        worksheet.Cells[row, column].Style.Border.Top.Color.SetColor(customStyle.BorderColor.Value);
                        worksheet.Cells[row, column].Style.Border.Right.Color.SetColor(customStyle.BorderColor.Value);
                        worksheet.Cells[row, column].Style.Border.Bottom.Color.SetColor(customStyle.BorderColor.Value);
                        worksheet.Cells[row, column].Style.Border.Left.Color.SetColor(customStyle.BorderColor.Value);
                    }
                    else
                    {
                        throw new ArgumentNullException(nameof(customStyle.BorderStyle), $"Para definir un '{typeof(Color)}' en la propiedad '{nameof(customStyle.BorderColor)}' de tu objeto tipo '{typeof(CustomStyle)}', especifica antes un valor tipo '{typeof(ExcelBorderStyle)}' diferente de '{typeof(ExcelBorderStyle)}.None' o 'null' en la propiedad '{nameof(customStyle.BorderStyle)}' de tu objeto tipo '{typeof(CustomStyle)}'.");
                    }
                }
                // Horizontal Alignment
                if (customStyle.HorizontalAlignment != null)
                    worksheet.Cells[row, column].Style.HorizontalAlignment = customStyle.HorizontalAlignment.Value;
                // Vertical Alignment
                if (customStyle.VerticalAlignment != null)
                    worksheet.Cells[row, column].Style.VerticalAlignment = customStyle.VerticalAlignment.Value;

                return excelWorksheet;
            }
            catch (Exception)
            {
                throw;
            }
        }
        private bool WorkSheetExist(List<ExcelWorksheet> worksheets, string nameWorksheet) => worksheets.Any(workSheet => workSheet.Name.Equals(nameWorksheet));
        private bool IsSingleTypeObject(Type type) => type.IsClass && !type.IsArray && !(type == typeof(string)) && !type.IsGenericType;
    }
}