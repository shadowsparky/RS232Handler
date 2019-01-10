using ClosedXML.Excel;
using System;
using System.IO.Ports;

namespace SerialPortExample
{
    class SerialPortProgram
    { 
        private SerialPort port = new SerialPort("COM2", 9600, Parity.None, 8, StopBits.One);
        private int count = 1;
        private XLWorkbook workbook = new XLWorkbook();
        private IXLWorksheet worksheet = null;

        static void Main(string[] args)
        {
            new SerialPortProgram();
        }

        private SerialPortProgram()
        {
            worksheet = workbook.Worksheets.Add("Sample");
            port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
            port.Open();
            Console.ReadKey();
            workbook.SaveAs("C:\\Users\\User\\Desktop\\docs\\test.xlsx");
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var buffer = new byte[port.ReadBufferSize];
            port.Read(buffer, 0, buffer.Length);
            Console.WriteLine(System.Text.Encoding.Default.GetString(buffer));
            worksheet.Cell("A" + count).Value = System.Text.Encoding.Default.GetString(buffer);
            count++;
        }
    }
}