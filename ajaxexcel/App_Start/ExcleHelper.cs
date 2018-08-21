using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ajaxexcel
{
    public class ExcleHelper
    {
        //高版本
        public static MemoryStream BuildWorkbook(DataTable dt)
        {
            HSSFWorkbook book = new HSSFWorkbook();
            ISheet sheet = book.CreateSheet("Sheet1");
            //Data Rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow drow = sheet.CreateRow(i);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = drow.CreateCell(j, CellType.String);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }
            //自动列宽
            for (int i = 0; i <= dt.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i, true);
            }

            MemoryStream file = new MemoryStream();
            book.Write(file);
            file.Seek(0, SeekOrigin.Begin);

            return file;
        }

        public static DataTable ToDataTable<T>(IList<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }

        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }

        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }
    }

    // 网上列子忽略
    public static class ExcelHelperForCs
    {

        /// <summary>
        ///  组装workbook.
        /// </summary>
        /// <param name="dt">dataTable资源</param>
        /// <param name="columnHeader">表头</param>
        /// <returns></returns>
        public static HSSFWorkbook BuildWorkbook1(DataTable dt, string columnHeader = "")
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(string.IsNullOrWhiteSpace(dt.TableName) ? "Sheet1" : dt.TableName);

            var dateStyle = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            var arrColWidth = new int[dt.Columns.Count];
            foreach (DataColumn item in dt.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dt.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            foreach (DataRow row in dt.Rows)
            {
                #region 表头 列头
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式
                    {
                        var headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(columnHeader);
                        //CellStyle
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;// 左右居中    
                        headStyle.VerticalAlignment = VerticalAlignment.Center;// 上下居中 
                                                                               // 设置单元格的背景颜色（单元格的样式会覆盖列或行的样式）    
                        headStyle.FillForegroundColor = (short)11;
                        //定义font
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, dt.Columns.Count - 1));
                    }
                    #endregion


                    #region 列头及样式
                    {
                        var headerRow = sheet.CreateRow(1);
                        //CellStyle
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;// 左右居中    
                        headStyle.VerticalAlignment = VerticalAlignment.Center;// 上下居中 
                                                                               //定义font
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dt.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        }
                    }
                    #endregion
                    if (columnHeader != "")
                    {
                        //header row
                        IRow row0 = sheet.CreateRow(0);
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            ICell cell = row0.CreateCell(i, CellType.String);
                            cell.SetCellValue(dt.Columns[i].ColumnName);
                        }
                    }

                    rowIndex = 2;
                }
                #endregion


                #region 内容
                var dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dt.Columns)
                {
                    var newCell = dataRow.CreateCell(column.Ordinal);

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型
                            newCell.SetCellValue(drValue);
                            break;
                        case "System.DateTime"://日期类型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle;//格式化显示
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }
                #endregion

                rowIndex++;
            }
            //自动列宽
            for (int i = 0; i <= dt.Columns.Count; i++)
                sheet.AutoSizeColumn(i, true);

            return workbook;
        }
        public static DataTable ToDataTable<T>(IList<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }
        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }

        /// <summary>
        /// DataTable导出Excel2007（.xlsx）
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="file">文件路径（.xlsx）</param>
        /// <param name="sheetname">Excel工作表名</param>
        public static void TableToExcelForXLSX2007(DataTable dt, string file, string sheetname)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();//建立Excel2007对象
            ISheet sheet = xssfworkbook.CreateSheet(sheetname);//新建一个名称为sheetname的工作簿

            //设置基本样式
            ICellStyle style = xssfworkbook.CreateCellStyle();
            style.WrapText = true;
            IFont font = xssfworkbook.CreateFont();
            font.FontHeightInPoints = 9;
            font.FontName = "Arial";
            style.SetFont(font);

            //设置统计样式
            ICellStyle style1 = xssfworkbook.CreateCellStyle();
            style1.WrapText = true;
            IFont font1 = xssfworkbook.CreateFont();
            font1.FontHeightInPoints = 9;
            font1.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            font1.FontName = "Arial";
            style1.SetFont(font1);

            //设置大类样式
            ICellStyle style2 = xssfworkbook.CreateCellStyle();
            style2.WrapText = true;
            //style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Gold.Index;
            //style2.FillPattern = FillPattern.SolidForeground;
            IFont font2 = xssfworkbook.CreateFont();
            font2.FontHeightInPoints = 9;
            font2.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            font2.FontName = "Arial";
            style2.SetFont(font2);


            //设置列名
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                var rowName = dt.Columns[i].ColumnName;
                string rowRealName = "";
                switch (rowName)
                {
                    case "IncomeType":
                        rowRealName = "交易类型";
                        break;
                    case "CreateDate":
                        rowRealName = "发生日期";
                        break;
                    case "ChangeAmount":
                        rowRealName = "合计金额";
                        break;
                    case "SubsectionName":
                        rowRealName = "分段名称";
                        break;
                    case "CorporateName":
                        rowRealName = "公司名称";
                        break;
                    case "Province":
                        rowRealName = "省份";
                        break;
                    case "ShuntName":
                        rowRealName = "项目";
                        break;
                    case "CountAmount":
                        rowRealName = "本年累计金额";
                        break;
                    default:
                        rowRealName = "";
                        break;
                }
                cell.SetCellValue(rowRealName);
                cell.CellStyle = style;
            }
            int paymentRowIndex = 1;
            //单元格赋值
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);

                    if (dt.Rows[i][j].ToString().Contains("小计") || dt.Rows[i][j].ToString().Contains("流量净额"))
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style2;
                    }
                    else if (dt.Rows[i][j].ToString().Contains("一") || dt.Rows[i][j].ToString().Contains("二") || dt.Rows[i][j].ToString().Contains("三"))
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style;
                    }

                }
                paymentRowIndex++;
            }

            //列宽自适应，只对英文和数字有效
            for (int i = 0; i <= dt.Rows.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            //获取当前列的宽度，然后对比本列的长度，取最大值
            for (int columnNum = 0; columnNum <= dt.Rows.Count; columnNum++)
            {
                int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (sheet.GetRow(rowNum) == null)
                    {
                        currentRow = sheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = sheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                sheet.SetColumnWidth(columnNum, columnWidth * 256);
            }

            using (System.IO.Stream stream = System.IO.File.OpenWrite(file))
            {
                //写入文件
                xssfworkbook.Write(stream);
                stream.Close();
            }
        }


        /// <summary>
        /// DataTable导出Excel2003（.xls）
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="file">文件路径（.xls）</param>
        /// <param name="sheetname">Excel工作表名</param>
        public static void TableToExcelForXLSX2003(DataTable dt, string file, string sheetname)
        {
            HSSFWorkbook xssfworkbook = new HSSFWorkbook();//建立Excel2003对象
            HSSFSheet sheet = (HSSFSheet)xssfworkbook.CreateSheet(sheetname);//新建一个名称为sheetname的工作簿


            //设置基本样式
            ICellStyle style = xssfworkbook.CreateCellStyle();
            style.WrapText = true;
            IFont font = xssfworkbook.CreateFont();
            font.FontHeightInPoints = 9;
            font.FontName = "Arial";
            style.SetFont(font);

            //设置统计样式
            ICellStyle style1 = xssfworkbook.CreateCellStyle();
            style1.WrapText = true;
            IFont font1 = xssfworkbook.CreateFont();
            font1.FontHeightInPoints = 9;
            font1.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            font1.FontName = "Arial";
            style1.SetFont(font1);

            //设置大类样式
            ICellStyle style2 = xssfworkbook.CreateCellStyle();
            style2.WrapText = true;
            //style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Gold.Index;
            //style2.FillPattern = FillPattern.SolidForeground;
            IFont font2 = xssfworkbook.CreateFont();
            font2.FontHeightInPoints = 9;
            font2.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            font2.FontName = "Arial";
            style2.SetFont(font2);

            //设置列名
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = (ICell)row.CreateCell(i);
                var rowName = dt.Columns[i].ColumnName;
                //cell.SetCellValue(dt.Columns[i].ColumnName);
                string rowRealName = "";
                switch (rowName)
                {
                    case "IncomeType":
                        rowRealName = "交易类型";
                        break;
                    case "CreateDate":
                        rowRealName = "发生日期";
                        break;
                    case "ChangeAmount":
                        rowRealName = "合计金额";
                        break;
                    case "SubsectionName":
                        rowRealName = "分段名称";
                        break;
                    case "CorporateName":
                        rowRealName = "公司名称";
                        break;
                    case "Province":
                        rowRealName = "省份";
                        break;
                    case "ShuntName":
                        rowRealName = "项目";
                        break;
                    case "CountAmount":
                        rowRealName = "本年累计金额";
                        break;
                    default:
                        rowRealName = "";
                        break;
                }
                cell.SetCellValue(rowRealName);
                cell.CellStyle = style;
            }
            int paymentRowIndex = 1;
            //单元格赋值
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);

                    if (dt.Rows[i][j].ToString().Contains("小计") || dt.Rows[i][j].ToString().Contains("流量净额"))
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style2;
                    }
                    else if (dt.Rows[i][j].ToString().Contains("一") || dt.Rows[i][j].ToString().Contains("二") || dt.Rows[i][j].ToString().Contains("三"))
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = style;

                    }

                }
                paymentRowIndex++;
            }
            //列宽自适应，只对英文和数字有效
            for (int i = 0; i <= dt.Rows.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            //获取当前列的宽度，然后对比本列的长度，取最大值
            for (int columnNum = 0; columnNum <= dt.Rows.Count; columnNum++)
            {
                int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (sheet.GetRow(rowNum) == null)
                    {
                        currentRow = sheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = sheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                sheet.SetColumnWidth(columnNum, columnWidth * 256);
            }
            using (System.IO.Stream stream = System.IO.File.OpenWrite(file))
            {
                xssfworkbook.Write(stream);
                stream.Close();
            }

        }
    }

}