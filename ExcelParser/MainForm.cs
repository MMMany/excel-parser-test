using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelParser
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls",
                RestoreDirectory = true,
            })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.progressBar1.Style = ProgressBarStyle.Marquee;

                    var finfo = new FileInfo(dialog.FileName);

                    Task.Run(async () =>
                    {
                        #region Study, How to use
                        if (false)
                        {
                            using (var wrapper = new ExcelWrapper(finfo.FullName))
                            {
                                var res1 = await wrapper.AsyncReadCell("Sink", "A8");
                                Console.WriteLine($"A8 value : {res1}");

                                await Task.Delay(TimeSpan.FromSeconds(2));

                                var res2 = await wrapper.AsyncReadCells("Sink DTD", "C4", "M4");
                                Console.Write("C4:M4 value : [");
                                var values = new List<string>();
                                for (int row = 0; row < res2.GetLength(0); row++)
                                {
                                    for (int col = 0; col < res2.GetLength(1); col++)
                                    {
                                        if (res2[row, col] != null)
                                        {
                                            values.Add(res2[row, col]);
                                        }
                                    }
                                }
                                for (int i = 0; i < values.Count; i++)
                                {
                                    if (i == values.Count - 1)
                                    {
                                        Console.WriteLine($"{values[i]}]");
                                    }
                                    else
                                    {
                                        Console.Write($"{values[i]}, ");
                                    }
                                }

                                await Task.Delay(TimeSpan.FromSeconds(2));

                                //var res3 = await wrapper.AsyncVLookUp("Sink", "CAT_Supported_Format", "A6:D12", 3);
                                var res3 = await wrapper.AsyncVLookUp("Sink", "How to Setup", "A1:D12", 2);
                                Console.WriteLine($"VLOOKUP value : {res3}");
                            }
                        }
                        #endregion

                        var parser = new Parser(finfo.FullName);
                        await parser.ParseTo("./");

                        this.progressBar1.BeginInvoke(new Action(() => this.progressBar1.Style = ProgressBarStyle.Blocks));
                    });
                }
            }
        }
    }
}
