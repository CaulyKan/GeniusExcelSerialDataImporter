using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public struct Frame
    {
        public Int16 Head;
        public Int16 StartTime;
        public Int16 EndTime;
        public Data[] Data;
        public Int16 Tail;
    }

    public struct Data
    {
        public double Flow;
        public double Speed;
        public double Voltage;
    }

    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            var fd = new OpenFileDialog();
            fd.Filter = "dat文件|*.dat";
            fd.Multiselect = true;
            fd.ShowDialog();
            if (fd.FileNames.Length > 0)
            {
                foreach (var file in fd.FileNames)
                {
                    var frames = ReadFramesFromFile(file);
                    var line = 2;

                    foreach (var frame in frames) { 

                        sheet.Cells[line, 1].Value = frame.StartTime;
                        sheet.Cells[line, 2].Value = frame.EndTime;
                        for (int i = 0; i < 30; i++)
                        {
                            sheet.Cells[line, 3 + i * 3].Value = frame.Data[i].Flow;
                            sheet.Cells[line, 4 + i * 3].Value = frame.Data[i].Speed;
                            sheet.Cells[line, 5 + i * 3].Value = frame.Data[i].Voltage;
                        }

                        line++;
                    }

                }
            }
        }

        private List<Frame> ReadFramesFromFile(string file)
        {
            var frameLength = 188;
            var result = new List<Frame>();
            var count = new FileInfo(file).Length / frameLength;
            var reader = new BinaryReader(new FileStream(file, FileMode.Open));

            for (int i = 0; i < count; i++)
            {
                var frame = new Frame();

                frame.Head = reader.ReadInt16();
                frame.StartTime = reader.ReadInt16();
                frame.EndTime = reader.ReadInt16();
                frame.Data = new Data[30];
                for (int j = 0; j < 30; j++)
                {
                    frame.Data[j].Flow = reader.ReadDouble();
                    frame.Data[j].Speed = reader.ReadDouble();
                    frame.Data[j].Voltage = reader.ReadDouble();
                }
                frame.Tail = reader.ReadInt16();

                result.Add(frame);
            }

            reader.Close();
            return result;
        }
    }
}
