using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoleLaboratoryExcel
{

    public class DataTableUtil
    {


        //static void Main()
        //{
        //    // 读取模板Excel到DataTable
        //    DataTable templateTable = ReadExcelToTable("模板.xlsx");

        //    // 行转列处理
        //    DataTable transformedTable = TransposeDataTable(templateTable);

        //    // 写入新的DataTable到Excel
        //    WriteTableToExcel(transformedTable, "整理后格式.xlsx");
        //}


        // 读取所有Excel文件到DataTable集合中
        //    List<DataTable> dataTables = new List<DataTable>();
        //    foreach (var file in excelFiles)
        //    {
        //        dataTables.Add(ReadExcelToDataTable(file));
        //    }

        //// 合并所有DataTable
        //combinedDataTable = MergeDataTables(dataTables);

        public static DataTable RemoveEmptyOrNullRowsEfficiently(DataTable table, string columnName)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            if (string.IsNullOrWhiteSpace(columnName))
                throw new ArgumentException("Column name cannot be null or empty.", nameof(columnName));

            if (!table.Columns.Contains(columnName))
                throw new ArgumentException($"Column '{columnName}' does not exist in the DataTable.", nameof(columnName));

            // 直接遍历DataTable.Rows并删除符合条件的行
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = table.Rows[i];
                if (row[columnName] == DBNull.Value || row[columnName].ToString() == "")
                {
                    row.Delete();
                }
            }

            // 提交删除操作
            table.AcceptChanges();

            // 返回DataTable对象（实际上是对传入对象的引用）
            return table;
        }

        public static DataTable PivotDataTable(DataTable sourceTable, string idColumnName, string sampleNameColumn, string pivotColumnName, string valueColumnName)
        {
            DataTable pivotedTable = new DataTable();

            // 添加ID果表
            pivotedTable.Columns.Add(idColumnName, typeof(string));

            // 添加Sample Name列到结果表
            pivotedTable.Columns.Add(sampleNameColumn, typeof(string));

            // 获取所有唯一的pivot值，并将它们作为列添加到结果表中
            var pivotValues = new HashSet<string>(from DataRow row in sourceTable.AsEnumerable()
                                                  select row.Field<string>(pivotColumnName));
            foreach (var pivotValue in pivotValues)
            {
                pivotedTable.Columns.Add(pivotValue, typeof(string)); // 使用double?来允许null值
            }

            // 填充数据到结果表
            var groups = from row in sourceTable.AsEnumerable()
                         group row by new { Id = row.Field<string>(idColumnName), SampleName = row.Field<string>(sampleNameColumn) } into g
                         select new
                         {
                             ID = g.Key.Id,
                             SampleName = g.Key.SampleName,
                             Values = g.ToDictionary(r => r.Field<string>(pivotColumnName), r => r.Field<string>(valueColumnName))
                         };

            foreach (var group in groups)
            {
                DataRow newRow = pivotedTable.NewRow();
                newRow[idColumnName] = group.ID;
                newRow[sampleNameColumn] = group.SampleName;

                foreach (var pivotValue in pivotValues)
                {
                    if (group.Values.ContainsKey(pivotValue))
                    {
                        newRow[pivotValue] = group.Values[pivotValue];
                    }
                    else
                    {
                        newRow[pivotValue] = DBNull.Value; // 使用DBNull表示缺失值
                    }
                }

                pivotedTable.Rows.Add(newRow);
            }

            return pivotedTable;
        }
        //DataTable中只保留特定的几列，并移除其他所有列
        public static DataTable RetainColumns(DataTable sourceTable, params string[] columnsToRetain)
        {
            // 创建一个新的DataTable，只包含需要的列
            DataTable resultTable = new DataTable();

            // 添加需要的列到新的DataTable中
            foreach (var columnName in columnsToRetain)
            {
                if (sourceTable.Columns.Contains(columnName))
                {
                    DataColumn column = new DataColumn(columnName, sourceTable.Columns[columnName].DataType);
                    resultTable.Columns.Add(column);
                }
                else
                {
                    throw new ArgumentException($"Column '{columnName}' does not exist in the source table.");
                }
            }

            // 遍历原始DataTable的每一行，复制需要的列的值到新DataTable中
            foreach (DataRow sourceRow in sourceTable.Rows)
            {
                DataRow newRow = resultTable.NewRow();
                foreach (var columnName in columnsToRetain)
                {
                    newRow[columnName] = sourceRow[columnName];
                }
                resultTable.Rows.Add(newRow);
            }

            return resultTable;
        }
        // 读取模板Excel到DataTable
        public static DataTable MergeDataTables(List<DataTable> dataTables)
        {

            DataTable mergedTable = dataTables[0].Clone();

            foreach (var table in dataTables)
            {
                foreach (DataRow row in table.Rows)
                {
                    mergedTable.ImportRow(row);
                }
            }

            return mergedTable;
        }
        public static DataTable ExcelToDataTable(string filePath, int headerRowIndex = 7)
        {
            DataTable dataTable = new DataTable();
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 根据文件扩展名确定使用哪个工厂类
                string fileExt = Path.GetExtension(filePath).ToLower();
                if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else
                {
                    throw new NotSupportedException("Unsupported file format.");
                }

                ISheet sheet = workbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(headerRowIndex);

                if (headerRow != null)
                {
                    // 用于跟踪列名出现的次数
                    Dictionary<string, int> columnNameCount = new Dictionary<string, int>();

                    // 处理列标题
                    for (int i = 0; i < headerRow.LastCellNum; i++)
                    {
                        string columnName = headerRow.GetCell(i)?.ToString() ?? $"Column_{i + 1}";

                        // 如果列名为空，使用列索引作为列名
                        if (string.IsNullOrWhiteSpace(columnName))
                        {
                            columnName = $"Column_{i + 1}";
                        }

                        // 检查列名是否已存在
                        if (columnNameCount.ContainsKey(columnName))
                        {
                            // 如果已存在，在列名后添加序号
                            columnNameCount[columnName]++;
                            // 跳过重复的列，不添加到DataTable中
                            continue;
                        }
                        else
                        {
                            columnNameCount[columnName] = 1;
                            dataTable.Columns.Add(columnName);
                        }
                    }

                    // 读取数据行
                    for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            int dataColumnIndex = 0; // 用于跟踪DataTable的列索引

                            for (int cellIndex = 0; cellIndex < row.LastCellNum && dataColumnIndex < dataTable.Columns.Count; cellIndex++)
                            {
                                string columnName = headerRow.GetCell(cellIndex)?.ToString() ?? $"Column_{cellIndex + 1}";

                                // 如果是第一次出现的列名，则处理数据
                                if (columnNameCount[columnName] == 1)
                                {
                                    ICell cell = row.GetCell(cellIndex);
                                    if (cell != null)
                                    {
                                        switch (cell.CellType)
                                        {
                                            case CellType.String:
                                                dataRow[dataColumnIndex] = cell.StringCellValue;
                                                break;
                                            case CellType.Numeric:
                                                if (DateUtil.IsCellDateFormatted(cell))
                                                {
                                                    dataRow[dataColumnIndex] = cell.DateCellValue;
                                                }
                                                else
                                                {
                                                    dataRow[dataColumnIndex] = cell.NumericCellValue;
                                                }
                                                break;
                                            case CellType.Boolean:
                                                dataRow[dataColumnIndex] = cell.BooleanCellValue;
                                                break;
                                            case CellType.Formula:
                                                dataRow[dataColumnIndex] = cell.CellFormula;
                                                break;
                                            case CellType.Blank:
                                                dataRow[dataColumnIndex] = DBNull.Value;
                                                break;
                                            case CellType.Error:
                                                dataRow[dataColumnIndex] = "Error";
                                                break;
                                            default:
                                                dataRow[dataColumnIndex] = cell.ToString();
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        dataRow[dataColumnIndex] = DBNull.Value;
                                    }
                                    dataColumnIndex++;
                                }
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }
            }
            return dataTable;
        }
        public static DataTable ReadExcelToTable(string filePath)
        {
            DataTable table = new DataTable();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // 假设Excel文件只有一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                // 添加列
                for (int col = 1; col <= colCount; col++)
                {
                    table.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // 添加行
                for (int row = 8; row <= rowCount; row++)
                {
                    DataRow newRow = table.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    table.Rows.Add(newRow);
                }
            }
            return table;
        }

        // 行转列处理
        public static DataTable TransposeDataTable(DataTable sourceTable)
        {
            DataTable transposedTable = new DataTable();

            // 创建列
            foreach (DataRow row in sourceTable.Rows)
            {
                if (!transposedTable.Columns.Contains(row["Sample Name"].ToString()))
                {
                    transposedTable.Columns.Add(row["Sample Name"].ToString(), typeof(string));
                }
            }

            // 创建行并填充数据
            foreach (DataColumn col in sourceTable.Columns)
            {
                if (col.ColumnName != "Sample Name" && col.ColumnName != "Target Name" && col.ColumnName != "Task" && col.ColumnName != "Reporter" && col.ColumnName != "Quencher")
                {
                    DataRow newRow = transposedTable.NewRow();
                    foreach (DataRow row in sourceTable.Rows)
                    {
                        if (row["Target Name"].ToString() == col.ColumnName)
                        {
                            newRow[row["Sample Name"].ToString()] = row[col];
                        }
                    }
                    transposedTable.Rows.Add(newRow);
                }
            }

            return transposedTable;
        }

        public static DataTable PivotDataTable(DataTable sourceTable, string pivotColumn, string valueColumn, bool keepNonPivotColumns = true)
        {
            // 获取所有唯一的pivot列值
            var uniquePivotValues = sourceTable.AsEnumerable()
                .Select(row => row.Field<string>(pivotColumn))
                .Distinct()
                .ToList();

            // 创建新的DataTable
            DataTable pivotedTable = new DataTable();

            // 添加ID列（如果存在）
            if (sourceTable.Columns.Contains("Sample Name"))
            {
                pivotedTable.Columns.Add("Sample Name", typeof(string));
            }

            // 添加pivot列作为新表的列
            foreach (var pivotValue in uniquePivotValues)
            {
                pivotedTable.Columns.Add(pivotValue, typeof(object)); // 使用object类型以支持不同的数据类型
            }

            // 如果需要保留pivot列，则将它们添加到新表中
            if (keepNonPivotColumns)
            {
                foreach (DataColumn column in sourceTable.Columns)
                {
                    if (column.ColumnName != pivotColumn && column.ColumnName != valueColumn && column.ColumnName != "Sample Name")
                    {
                        pivotedTable.Columns.Add(column.ColumnName, column.DataType);
                    }
                }
            }

            // 填充新表的数据
            foreach (DataRow sourceRow in sourceTable.Rows)
            {
                DataRow newRow = pivotedTable.NewRow();

                // 设置ID列（如果存在）
                if (pivotedTable.Columns.Contains("Sample Name"))
                {
                    newRow["Sample Name"] = sourceRow["Sample Name"];
                }

                // 设置pivot列对应的值
                var pivotValue = sourceRow[pivotColumn].ToString();
                newRow[pivotValue] = sourceRow[valueColumn];

                // 如果需要保留非pivot列，则将它们复制到新行中
                if (keepNonPivotColumns)
                {
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        if (column.ColumnName != pivotColumn && column.ColumnName != valueColumn && column.ColumnName != "Sample Name")
                        {
                            newRow[column.ColumnName] = sourceRow[column.ColumnName];
                        }
                    }
                }

                pivotedTable.Rows.Add(newRow);
            }

            return pivotedTable;
        }


        // 读取模板Excel到DataTable
        public static void WriteTableToExcel(DataTable table, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromDataTable(table, true);
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }

        // 新增2列
        public static void AddWellColumns(DataTable dt,string colName)
        {
            // 添加两个新列
            dt.Columns.Add("Number", typeof(string)).SetOrdinal(0);    // 设置为第一列
            dt.Columns.Add("Group", typeof(string)).SetOrdinal(1);     // 设置为第二列

            foreach (DataRow row in dt.Rows)
            {
                string wellValue = row[colName].ToString();  // 例如 "A1", "B2" 等
                if (wellValue.Length >= 2)
                {
                    // 提取字母组别
                    string group = wellValue.Substring(0, 1);

                    // 提取数字
                    string number = wellValue.Substring(1);

                    // 设置新列的值
                    row["Group"] = group;
                    row["Number"] = number;
                }
            }
        }
        public static DataTable TransformRowsToColumns(
            DataTable sourceTable,
            string[] groupByColumns,                  // 分组列名数组，如 ["Well", "Sample Name"]
            string categoryColumn,                    // 用于区分不同类别的列名(如Target Name)
            string[] valueColumns,                    // 需要转换的值列名数组(如 Cт, Quantity)
            string[] suffixes = null)                // 对应值列的后缀数组(如 -Cт, -Quantity)
        {
            // 首先过滤掉分类列为空的记录
            var filteredRows = sourceTable.AsEnumerable()
                .Where(row => !string.IsNullOrWhiteSpace(row.Field<string>(categoryColumn)))
                .CopyToDataTable();

            // 创建结果表
            var resultTable = new DataTable();

            // 添加分组列
            foreach (var groupCol in groupByColumns)
            {
                resultTable.Columns.Add(groupCol, typeof(string));
            }

            // 获取所有唯一的类别值
            var categories = filteredRows.AsEnumerable()
                .Select(row => row.Field<string>(categoryColumn))
                .Distinct()
                .ToList();

            // 添加数据列
            for (int j = 0; j < valueColumns.Length; j++)
            {
                foreach (var category in categories)
                {
                    string suffix = (suffixes != null && j < suffixes.Length) ? suffixes[j] : "";
                    string columnName = $"{category}{suffix}";
                    resultTable.Columns.Add(columnName, typeof(string));
                }
            }

            // 按多列分组处理数据
            var groupedData = filteredRows.AsEnumerable()
                .GroupBy(row => new
                {
                    GroupValues = string.Join("_", groupByColumns.Select(col => row.Field<string>(col)))
                });

            // 处理每个分组的数据
            foreach (var group in groupedData)
            {
                var newRow = resultTable.NewRow();

                // 设置分组列的值
                var firstRow = group.First();
                foreach (var groupCol in groupByColumns)
                {
                    newRow[groupCol] = firstRow.Field<string>(groupCol);
                }

                foreach (DataRow sourceRow in group)
                {
                    string category = sourceRow.Field<string>(categoryColumn);

                    // 填充每个列的数据
                    for (int i = 0; i < valueColumns.Length; i++)
                    {
                        string suffix = (suffixes != null && i < suffixes.Length) ? suffixes[i] : "";
                        string columnName = $"{category}{suffix}";

                        object value = sourceRow[valueColumns[i]];
                        newRow[columnName] = value == DBNull.Value ? "" : value.ToString();
                    }
                }

                resultTable.Rows.Add(newRow);
            }

            return resultTable;
        }

        public static void WriteTableToExcelSheets(Dictionary<string, DataTable> tableDict, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                foreach (var kvp in tableDict)
                {
                    // 获取不含扩展名的文件名作为sheet名
                    string sheetName = Path.GetFileNameWithoutExtension(kvp.Key);

                    // 创建新的工作表
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // 将数据写入工作表
                    worksheet.Cells["A1"].LoadFromDataTable(kvp.Value, true);

                    // 自动调整列宽
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                // 保存Excel文件
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }

        // QPCR数据转换的具体实现示例
        public static DataTable TransformQPCRData(DataTable sourceTable)
        {
            return TransformRowsToColumns(
                sourceTable,
                groupByColumns: new[] { "Well", "Sample Name" },
                categoryColumn: "Target Name",
                valueColumns: new[] { "Cт", "Quantity" },
                suffixes: new[] { "-Cт", "-Quantity" }
            );
        }
    }
}
