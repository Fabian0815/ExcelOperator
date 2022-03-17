using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelOperator
{
    public class ExcelOperator
    {
        private bool output = false;
        private string message;
        private IWorkbook workbook;
        private string file;
        public static void Main()
        {
            //以下为测试代码
            string filePath = @"F:\ExcelOperator\TestData\Data.xlsx";//文件需要修改成本地路径
            ExcelOperator excelOperator = new ExcelOperator();
            bool bRes = excelOperator.ReadExcel(filePath);
            if (!bRes)
            {
                Console.WriteLine($"解析文件失败，失败原因为{excelOperator.Message()}");
            }
            else
            {
                //读取单元格内容测试
                string strResult;
                bRes = excelOperator.GetCellValue(0, 0, 0, out strResult);
                if (bRes)
                {
                    Console.WriteLine($"读取字符串数据成功：{strResult}");
                }
                else
                {

                    Console.WriteLine("读取字符串数据失败");
                }

                int intResult;
                bRes = excelOperator.GetCellValue(0, 1, 0, out intResult);
                if (bRes)
                {
                    Console.WriteLine($"读取整型数据成功：{intResult}");
                }
                else
                {

                    Console.WriteLine("读取整型数据失败");
                }

                double doubleResult;
                bRes = excelOperator.GetCellValue(0, 2, 0, out doubleResult);
                if (bRes)
                {
                    Console.WriteLine($"读取浮点型数据成功：{doubleResult}");
                }
                else
                {

                    Console.WriteLine("读取浮点型数据失败");
                }

                bool boolResult;
                bRes = excelOperator.GetCellValue(0, 3, 0, out boolResult);
                if (bRes)
                {
                    Console.WriteLine($"读取布尔型数据成功：{boolResult}");
                }
                else
                {

                    Console.WriteLine("读取布尔型数据失败");
                }

                //写入单元格内容测试
                bRes = excelOperator.SetCellValue(0, 3, 3, "文本内容");
                if (bRes)
                {
                    Console.WriteLine("写入字符串数据成功");
                }
                else
                {
                    Console.WriteLine("写入字符串数据失败");
                }
                bRes = excelOperator.SetCellValue(0, 3, 4, 5);
                if (bRes)
                {
                    Console.WriteLine("写入整型数据成功");
                }
                else
                {
                    Console.WriteLine("写入整型数据失败");
                }
                bRes = excelOperator.SetCellValue(0, 3, 5, 5.5);
                if (bRes)
                {
                    Console.WriteLine("写入浮点型数据成功");
                }
                else
                {
                    Console.WriteLine("写入浮点型数据失败");
                }
                bRes = excelOperator.SetCellValue(0, 3, 6, true);
                if (bRes)
                {
                    Console.WriteLine("写入布尔型数据成功");
                }
                else
                {
                    Console.WriteLine("写入布尔型数据失败");
                }
                bRes = excelOperator.WriteExcel();
                if (bRes)
                {
                    Console.WriteLine("保存文件成功");
                }
                else
                {
                    Console.WriteLine($"保存文件失败，失败原因为：{excelOperator.Message()}");
                }
            }
        }
        public void HelloWorld()
        {
            Console.WriteLine("Hello World!");
        }


        public string Message()
        {
            return message;
        }
        public void SetMessage(string message)
        {
            this.message = message;
        }
        public bool ReadExcel(string filePath)
        {
            try
            {
                workbook = null;
                FileStream file = File.OpenRead(filePath);
                string extension = Path.GetExtension(filePath);
                if (extension.ToLower().Equals(".xls"))
                {
                    workbook = new HSSFWorkbook(file);
                }
                else if (extension.ToLower().Equals(".xlsx"))
                {
                    workbook = new XSSFWorkbook(file);
                }
                else
                {
                    file.Close();
                    message = "文件后缀错误";
                    return false;
                }
                file.Close();
                if (workbook != null)
                {
                    this.file = filePath;
                    message = "解析成功";
                    return true;
                }
                else
                {
                    message = "解析失败";
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool WriteExcel()
        {
            try
            {
                if (workbook != null)
                {
                    FileStream file = File.Create(this.file);
                    workbook.Write(file);
                    file.Close();
                    return true;
                }
                else
                {
                    message = "保存失败";
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public int SheetCount()
        {
            if (workbook != null)
            {
                return workbook.NumberOfSheets;
            }
            return -1;
        }
        public int RowCount(int sheetIndex)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum + 1;
                    return rowCount;
                }
                return -1;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return -1;
        }
        public int ColumnCount(int sheetIndex)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                int columnCount = -1;
                if (sheet != null)
                {
                    for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null && row.LastCellNum > columnCount)
                        {
                            columnCount = row.LastCellNum;
                        }
                    }
                }
                return columnCount;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return -1;
        }

        //读取单元格类型
        public bool GetCellType(int sheetIndex, int row, int column, ref int result)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell != null)
                        {
                            result = (int)tCell.CellType;
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        //读取单元格内容
        public bool GetCellValue(int sheetIndex, int row, int column, out string result)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell != null)
                        {
                            result = tCell.StringCellValue;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            result = "";
            return false;
        }
        public bool GetCellValue(int sheetIndex, int row, int column, out int result)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell != null)
                        {
                            double value = tCell.NumericCellValue;
                            result = (int)value;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            result = 0;
            return false;
        }
        public bool GetCellValue(int sheetIndex, int row, int column, out double result)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell != null)
                        {
                            result = tCell.NumericCellValue;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            result = 0;
            return false;
        }
        public bool GetCellValue(int sheetIndex, int row, int column, out bool result)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell != null)
                        {
                            result = tCell.BooleanCellValue;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            result = false;
            return false;
        }


        //写入单元格内容
        public bool SetCellValue(int sheetIndex, int row, int column, string value)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, int value)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, double value)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, bool value)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, string value, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)horizontalAlign);
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)verticalAlign);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, int value, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)horizontalAlign);
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)verticalAlign);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, double value, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)horizontalAlign);
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)verticalAlign);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellValue(int sheetIndex, int row, int column, bool value, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.SetCellValue(value);
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)horizontalAlign);
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)verticalAlign);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        //设置样式
        public bool SetCellBorderStyle(int sheetIndex, int row, int column, int style)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_LEFT, (BorderStyle)style);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_TOP, (BorderStyle)style);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_RIGHT, (BorderStyle)style);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_BOTTOM, (BorderStyle)style);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellBorderStyle(int sheetIndex, int row, int column, int borderLeft, int borderTop, int borderRight, int borderBottom)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_LEFT, (BorderStyle)borderLeft);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_TOP, (BorderStyle)borderTop);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_RIGHT, (BorderStyle)borderRight);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BORDER_BOTTOM, (BorderStyle)borderBottom);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        //设置边框颜色之前需要先设置边框样式
        public bool SetCellBorderColor(int sheetIndex, int row, int column, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.LEFT_BORDER_COLOR, colorIndex);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.TOP_BORDER_COLOR, colorIndex);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.RIGHT_BORDER_COLOR, colorIndex);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.BOTTOM_BORDER_COLOR, colorIndex);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellBackgroundColor(int sheetIndex, int row, int column, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.FILL_FOREGROUND_COLOR, colorIndex);
                            CellUtil.SetCellStyleProperty(tCell, CellUtil.FILL_PATTERN, FillPattern.SolidForeground);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        public bool CreateCellStyle(ref short styleIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle style = workbook.CreateCellStyle();
                styleIndex = style.Index;
                return true;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool GetCellStyle(int sheetIndex, int row, int column, ref short styleIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            styleIndex = tCell.CellStyle.Index;
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyle(int sheetIndex, int row, int column, short styleIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            tCell.CellStyle = workbook.GetCellStyleAt(styleIndex);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        public bool SetCellFontColor(int sheetIndex, int row, int column, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            //以下方式会出现问题：重复调用本接口，花费时间越来越多，猜测是因为IFont越来越多导致的
                            //ICellStyle cellStyle = tCell.CellStyle;
                            //IFont fontOrigin = cellStyle.GetFont(workbook);
                            //IFont fontNew = workbook.CreateFont();
                            //fontNew.Charset = fontOrigin.Charset;
                            //fontNew.FontHeight = fontOrigin.FontHeight;
                            //fontNew.FontHeightInPoints = fontOrigin.FontHeightInPoints;
                            //fontNew.FontName = fontOrigin.FontName;
                            //fontNew.IsBold = fontOrigin.IsBold;
                            //fontNew.IsItalic = fontOrigin.IsItalic;
                            //fontNew.IsStrikeout = fontOrigin.IsStrikeout;
                            //fontNew.TypeOffset = fontOrigin.TypeOffset;
                            //fontNew.Underline = fontOrigin.Underline;
                            //fontNew.Color = colorIndex;
                            //CellUtil.SetFont(tCell, fontNew);

                            //这种方式会同时改动到其他单元格的字体
                            ICellStyle cellStyle = tCell.CellStyle;
                            cellStyle.GetFont(workbook).Color = colorIndex;
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        //设置文字字号
        public bool SetCellFontSize(int sheetIndex, int row, int column, double fontSize)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            //以下方式会出现问题：重复调用本接口，花费时间越来越多，猜测是因为IFont越来越多导致的
                            //ICellStyle cellStyle = tCell.CellStyle;
                            //IFont fontOrigin = cellStyle.GetFont(workbook);
                            //IFont fontNew = workbook.CreateFont();
                            //fontNew.Charset = fontOrigin.Charset;
                            //fontNew.FontHeightInPoints = fontSize;
                            //fontNew.FontName = fontOrigin.FontName;
                            //fontNew.IsBold = fontOrigin.IsBold;
                            //fontNew.IsItalic = fontOrigin.IsItalic;
                            //fontNew.IsStrikeout = fontOrigin.IsStrikeout;
                            //fontNew.TypeOffset = fontOrigin.TypeOffset;
                            //fontNew.Underline = fontOrigin.Underline;
                            //fontNew.Color = fontOrigin.Color;
                            //CellUtil.SetFont(tCell, fontNew);

                            //这种方式会同时改动到其他单元格的字体
                            ICellStyle cellStyle = tCell.CellStyle;
                            cellStyle.GetFont(workbook).FontHeightInPoints = fontSize;
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellHorizontalAlign(int sheetIndex, int row, int column, short align)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)align);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellVerticalAlign(int sheetIndex, int row, int column, short align)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)align);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellAlign(int sheetIndex, int row, int column, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    IRow tRow = sheet.GetRow(row);
                    if (tRow == null)
                    {
                        tRow = sheet.CreateRow(row);
                    }
                    if (tRow != null)
                    {
                        ICell tCell = tRow.GetCell(column);
                        if (tCell == null)
                        {
                            tCell = tRow.CreateCell(column);
                        }
                        if (tCell != null)
                        {
                            ICellStyle cellStyle = tCell.CellStyle;
                            CellUtil.SetAlignment(tCell, (HorizontalAlignment)horizontalAlign);
                            CellUtil.SetVerticalAlignment(tCell, (VerticalAlignment)verticalAlign);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }


        //下面是设置样式接口
        public bool SetCellStyleBorderStyle(short styleIndex, int borderLeft, int borderTop, int borderRight, int borderBottom)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.BorderLeft = (BorderStyle)borderLeft;
                    cellStyle.BorderTop = (BorderStyle)borderTop;
                    cellStyle.BorderRight = (BorderStyle)borderRight;
                    cellStyle.BorderBottom = (BorderStyle)borderBottom;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleBorderStyle(short styleIndex, int style)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.BorderLeft = (BorderStyle)style;
                    cellStyle.BorderTop = (BorderStyle)style;
                    cellStyle.BorderRight = (BorderStyle)style;
                    cellStyle.BorderBottom = (BorderStyle)style;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        //设置边框颜色之前需要先设置边框样式
        public bool SetCellStyleBorderColor(short styleIndex, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.LeftBorderColor = colorIndex;
                    cellStyle.TopBorderColor = colorIndex;
                    cellStyle.RightBorderColor = colorIndex;
                    cellStyle.BottomBorderColor = colorIndex;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleBackgroundColor(short styleIndex, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.FillBackgroundColor = colorIndex;
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleFontColor(short styleIndex, short colorIndex)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.GetFont(workbook).Color = colorIndex;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }

        //设置文字字号
        public bool SetCellStyleFontSize(short styleIndex, double fontSize)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.GetFont(workbook).FontHeightInPoints = fontSize;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleHorizontalAlign(short styleIndex, short align)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.Alignment = (HorizontalAlignment)align;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleVerticalAlign(short styleIndex, short align)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.VerticalAlignment = (VerticalAlignment)align;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
        public bool SetCellStyleAlign(short styleIndex, short horizontalAlign, short verticalAlign)
        {
            //捕获操作异常，避免程序崩溃
            try
            {
                ICellStyle cellStyle = workbook.GetCellStyleAt(styleIndex);
                if (cellStyle != null)
                {
                    cellStyle.Alignment = (HorizontalAlignment)horizontalAlign;
                    cellStyle.VerticalAlignment = (VerticalAlignment)verticalAlign;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                if (output)
                {
                    Console.WriteLine(ex.Message);
                }
                message = ex.Message;
                //throw ex;
            }
            return false;
        }
    }
}
