using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelParser
{
    internal class Parser : IDisposable
    {
        public Parser(string excelPath)
        {
            this.ExcelFile = new FileInfo(excelPath);
            if (!this.ExcelFile.Exists)
            {
                throw new FileNotFoundException(excelPath);
            }
            _excelWrapper = new ExcelWrapper(excelPath);
        }

        public async Task<FileInfo[]> ParseTo(string outputDirPath)
        {
            if (Directory.Exists(outputDirPath) == false)
            {
                throw new DirectoryNotFoundException(outputDirPath);
            }
            var outputDir = new DirectoryInfo(outputDirPath);

            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            var xDocList = new List<XDocument>();
            //foreach (var type in new[] { "TMDS", "FRL" })
            //{
            //    foreach (var mode in new[] { Mode.SIGNAL_OFF, Mode.SIGNAL_ON, Mode.QMS, Mode.GAMING })
            //    {
            //        token.ThrowIfCancellationRequested();

            //        var xDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "false"));
            //        var root = new XElement("DATAOBJ");
            //        xDoc.Add(root);

            //        root.Add(new XElement("HEADER",
            //            new XAttribute("TYPE", "HDMI2_SINK_CDF"),
            //            new XAttribute("VERSION", "1.0")));

            //        await ParseCDF(mode, root);
            //    }
            //}
            var xDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "false"));
            var root = new XElement("DATAOBJ");
            xDoc.Add(root);
            root.Add(new XElement("HEADER",
                new XAttribute("TYPE", "HDMI2_SINK_CDF"),
                new XAttribute("VERSION", "1.0")));

            await ParseCDF(Mode.SIGNAL_ON, root);
            await ParseDTD(root);

            Console.WriteLine(xDoc.Declaration);
            Console.WriteLine(xDoc.ToString());

            var timestamp = DateTime.Now.ToString("yyMMdd");

            return null;
        }

        public void Abort()
        {
            _cts.Cancel();
        }

        #region Private
        private enum Mode
        {
            SIGNAL_OFF = 0,
            SIGNAL_ON,
            QMS,
            GAMING
        }

        private CancellationTokenSource _cts;

        private ExcelWrapper _excelWrapper;
        private FileInfo ExcelFile;

        private async Task ParseCDF(Mode mode, XElement parent)
        {
            var token = _cts.Token;

            string col;
            switch (mode)
            {
                case Mode.SIGNAL_OFF:
                    col = "B";
                    break;
                case Mode.SIGNAL_ON:
                    col = "C";
                    break;
                case Mode.QMS:
                    col = "D";
                    break;
                case Mode.GAMING:
                    col = "E";
                    break;
                default:
                    throw new InvalidDataException($"Unknown mode : {mode}");
            }

            var data = new XElement("DATA");
            var sheetName = "Sink";
            foreach (var item in MapCDF)
            {
                token.ThrowIfCancellationRequested();

                var elem = new XElement(item.Key);
                data.Add(elem);

                if (item.Value is int row)
                {
                    var value = await _excelWrapper.AsyncReadCell(sheetName, col + row, token);
                    if (BooleanCheck.Contains(value.ToLower()))
                    {
                        elem.Value = value.ToLower() == "y" ? "YES" : "NO";
                    }
                    else
                    {
                        elem.Value = value;
                    }
                }
                else if (item.Value is Dictionary<string, int> map)
                {
                    var fmts = new List<string>();
                    foreach (var item2 in map)
                    {
                        token.ThrowIfCancellationRequested();

                        var value = await _excelWrapper.AsyncReadCell(sheetName, col + item2.Value, token);
                        if (value.ToLower() == "y")
                        {
                            fmts.Add(item2.Key);
                        }
                    }
                    if (fmts.Count() > 0)
                    {
                        elem.Value = string.Join(",", fmts);
                    }
                }
            }

            parent.Add(data);
        }

        private async Task ParseDTD(XElement parent)
        {
            var token = _cts.Token;

            var data = new XElement("Sink_DTD");
            var sheetName = "Sink DTD";

            // Check DTD profiles
            var profileNames = await _excelWrapper.AsyncReadCells(sheetName, "C4", "M4", token);
            var profiles = new Dictionary<string, string>();
            for (int col = 0; col < profileNames.Length; col++)
            {
                token.ThrowIfCancellationRequested();

                var p = profileNames[0, col];
                if (p == null) break;

                profiles.Add(p, Convert.ToChar('C' + col).ToString());
            }

            foreach (var p in profiles)
            {
                token.ThrowIfCancellationRequested();

                var elemProf = new XElement(p.Key);
                data.Add(elemProf);

                var col = p.Value;
                foreach (var item in MapDTD)
                {
                    token.ThrowIfCancellationRequested();

                    var elem = new XElement(item.Key);
                    elemProf.Add(elem);

                    if (item.Value is int row)
                    {
                        var value = await _excelWrapper.AsyncReadCell(sheetName, col + row, token);
                        if (BooleanCheck.Contains(value.ToLower()))
                        {
                            elem.Value = value.ToLower() == "y" ? "YES" : "NO";
                        }
                        else
                        {
                            elem.Value = value;
                        }
                    }
                    else if (item.Value is Dictionary<string, int> map)
                    {
                        var fmts = new List<string>();
                        foreach (var item2 in map)
                        {
                            token.ThrowIfCancellationRequested();

                            var value = await _excelWrapper.AsyncReadCell(sheetName, col + item2.Value, token);
                            if (value.ToLower() == "y")
                            {
                                fmts.Add(item2.Key);
                            }
                        }
                        if (fmts.Count() > 0)
                        {
                            elem.Value = string.Join(",", fmts);
                        }
                    }
                }
            }

            parent.Add(data);
        }

        private static readonly string[] BooleanCheck = new[] { "y", "n" };

        private static readonly Dictionary<string, object> MapCDF = new Dictionary<string, object>
        {
            { "CAT_Sink_Basic", 7 },
            { "CAT_Sink_Extension", 8 },
            { "CAT_Supported_Format", 10 },
            { "CAT_New_Feature", 11 },
            { "CAT_Addon", 12 },
            { "CAT_Formats", new Dictionary<string, int>
            {
                { "11", 14 },
                { "12", 15 },
                { "13", 16 }
            } },
        };

        private static readonly Dictionary<string, object> MapDTD = new Dictionary<string, object>
        {
            { "DTD_V1", 5 },
            { "DTD_V2", 6 },
            { "DTD_V3", 7 },
            { "DTD_V4", new Dictionary<string, int>
            {
                { "51", 9 },
                { "52", 10 },
                { "53", 11 }
            } },
        };
        #endregion

        #region IDisposable
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
                _excelWrapper.Dispose();
                _excelWrapper = null;
                disposedValue = true;
            }
        }

        // // TODO: 비관리형 리소스를 해제하는 코드가 'Dispose(bool disposing)'에 포함된 경우에만 종료자를 재정의합니다.
        // ~Parser()
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
