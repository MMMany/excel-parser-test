using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Collections.Specialized;
using System.Text.RegularExpressions;

namespace ExcelParser
{
    internal class ExcelWrapper : IDisposable
    {
        public ExcelWrapper(string fileName)
        {
            OpenFile(fileName);
        }

        public string ReadCell(string sheetName, string cellName)
        {
            if (_workbook == null) return null;

            Worksheet sheet = _workbook.Sheets[sheetName];
            Range cell = sheet.Range[cellName];

            var value = cell.Value2.ToString();

            ReleaseComObjects(cell, sheet);

            return value;
        }

        public async Task<string> AsyncReadCell(string sheetName, string cellAddress, CancellationToken token = default)
        {
            return await Task.Run(() => ReadCell(sheetName, cellAddress), token);
        }

        public string ReadCell(string sheetName, int row, int col)
        {
            if (_workbook == null) return null;

            Worksheet sheet = _workbook.Sheets[sheetName];
            Range cell = sheet.Cells[row, col];

            var value = cell.Value2.ToString();

            ReleaseComObjects(cell, sheet);

            return value;
        }

        public async Task<string> AsyncReadCell(string sheetName, int row, int col, CancellationToken token = default)
        {
            return await Task.Run(() => ReadCell(sheetName, row, col), token);
        }

        public string[,] ReadCells(string sheetName, string startCell, string endCell)
        {
            if (_workbook == null) return null;

            Worksheet sheet = _workbook.Sheets[sheetName];
            Range range = sheet.Range[$"{startCell}:{endCell}"];

            object[,] values = range.Value2;
            string[,] result = null;

            if (values != null)
            {
                result = new string[values.GetLength(0), values.GetLength(1)];
                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int col = 1; col <= values.GetLength(1); col++)
                    {
                        result[row - 1, col - 1] = values[row, col]?.ToString();
                    }
                }
            }

            ReleaseComObjects(range, sheet);

            return result;
        }

        public async Task<string[,]> AsyncReadCells(string sheetName, string start, string end, CancellationToken token = default)
        {
            return await Task.Run(() => ReadCells(sheetName, start, end), token);
        }

        public string VLookUp(string sheetName, string lookupValue, string tableRange, int colNum)
        {
            if (sheetName == null) return null;

            Worksheet sheet = _workbook.Sheets[sheetName];
            Range lookupRange = sheet.Range[tableRange];

            string value = null;

            try
            {
                value = _app.WorksheetFunction.VLookup(lookupValue, lookupRange, colNum, false);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            

            ReleaseComObjects(lookupRange, sheet);

            return value;
        }

        public async Task<string> AsyncVLookUp(string sheetName, string lookupValue, string tableRange, int colNum, CancellationToken token = default)
        {
            return await Task.Run(() => VLookUp(sheetName, lookupValue, tableRange, colNum), token);
        }

        #region Private
        private Application _app = new Application
        {
            Visible = false,
        };
        private Workbook _workbook;

        private void OpenFile(string fileName)
        {
            var finfo = new FileInfo(fileName);
            if (!finfo.Exists)
            {
                throw new FileNotFoundException($"Cannot found \"{finfo.FullName}\".");
            }

            _workbook = _app.Workbooks.Open(finfo.FullName);
        }

        private void ReleaseComObjects(params object[] comObjects)
        {
            foreach (var obj in  comObjects)
            {
                Marshal.FinalReleaseComObject(obj);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region IDisposable Support
        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: 관리형 상태(관리형 개체)를 삭제합니다.
                }

                // TODO: 비관리형 리소스(비관리형 개체)를 해제하고 종료자를 재정의합니다.
                // TODO: 큰 필드를 null로 설정합니다.
                _workbook.Close();
                _app.Quit();

                Marshal.FinalReleaseComObject(_workbook);
                Marshal.FinalReleaseComObject(_app);

                _workbook = null;
                _app = null;
                disposedValue = true;
            }
        }

        // // TODO: 비관리형 리소스를 해제하는 코드가 'Dispose(bool disposing)'에 포함된 경우에만 종료자를 재정의합니다.
        // ~ExcelWrapper()
        // {
        //     // 이 코드를 변경하지 마세요. 'Dispose(bool disposing)' 메서드에 정리 코드를 입력합니다.
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // 이 코드를 변경하지 마세요. 'Dispose(bool disposing)' 메서드에 정리 코드를 입력합니다.
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
