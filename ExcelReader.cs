using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Tools
{
    /// <summary>
    /// Install LinqToExcel NuGet
    /// 
    /// Excel extension .xlsx
    ///
    /// T must have empty constructor and public properties
    /// with the same name as the title of the column in the excel sheet
    ///
    /// You can use UrlConnection and make your own querys
    /// </summary>
    public class ExcelReader
    {
        private string PathExcelFile { get; set; }
        public ExcelQueryFactory UrlConnection { get; }

        public ExcelReader(string pathExcelFile)
        {
            this.PathExcelFile = pathExcelFile;
            UrlConnection = new ExcelQueryFactory(PathExcelFile)
            {
                ReadOnly = true,
                UsePersistentConnection = true
            };
        }

        public IEnumerable<T> GetObjectsFromSheet1<T>()
        {
            return GetObjectsFromSheet<T>("Sheet1");
        }

        public IEnumerable<T> GetObjectsFromSheet<T>(string sheetName)
        {
            var query = from a in UrlConnection.Worksheet<T>(sheetName) select a;

            return query;
        }

        public IEnumerable<T> GetObjectsFromSheet<T>(int sheetIndex)
        {
            var query = from a in UrlConnection.Worksheet<T>(sheetIndex) select a;

            return query;
        }

        public IEnumerable<T> GetObjectsFromSheetAndSpecificColumn<T>(string sheetName, int column)
        {
            var query = from a in UrlConnection.WorksheetNoHeader(sheetName) select a[column];

            return (IEnumerable<T>)query;
        }

        public IEnumerable<T> GetObjectsFromSheetAndSpecificColumn<T>(int sheetIndex, int column)
        {
            var query = from a in UrlConnection.WorksheetNoHeader(sheetIndex) select a[column];

            return (IEnumerable<T>)query;
        }

        ~ExcelReader()
        {
            try
            {
                UrlConnection.Dispose();
            }
            catch (Exception){}
        }
    }
}
