using ClosedXML.Excel;
using System;
using System.IO;
using System.IO.Ports;

namespace Program
{
    class Program
    { 
        private SerialPort port = new SerialPort("COM2", 9600, Parity.None, 8, StopBits.One);
        private int count = 1;
        private XLWorkbook workbook = new XLWorkbook();
        private IXLWorksheet worksheet = null;

        static void Main(string[] args)
        {
            new Program();
        }

        private Program()
        {
            worksheet = workbook.Worksheets.Add("Sample");
            port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
            try
            {
                port.Open();
            }
            catch (UnauthorizedAccessException)
            {
                print("Запрещен доступ к порту или выбранный вами порт уже занят другим процессом");
            } 
            catch (ArgumentOutOfRangeException)
            {
                print("Один или несколько свойств для данного экземпляра являются недопустимыми. Например " +
                    "System.IO.Ports.SerialPort.Parity, System.IO.Ports.SerialPort.DataBits, или System.IO.Ports.SerialPort.Handshake " +
                    "свойства не являются допустимыми; System.IO.Ports.SerialPort.BaudRate меньше " +
                    "или равно нулю; System.IO.Ports.SerialPort.ReadTimeout или System.IO.Ports.SerialPort.WriteTimeout " +
                    "свойство меньше нуля и не System.IO.Ports.SerialPort.InfiniteTimeout.");
            }
            catch(ArgumentException)
            {
                print("Имя порта не начинается с «COM» или тип файла порта не поддерживается.");
            }
            catch(IOException)
            {
                print("Не удалось задать порт. Возможно, вы в строку для порта записываете ОБЪЕКТ");
            }
            catch(InvalidOperationException)
            {
                print("Указанный порт на текущий экземпляр порта уже открыт.");
            }
            Console.ReadKey();
            workbook.SaveAs("C:\\Users\\User\\Desktop\\docs\\test.xlsx");
        }

        private void print(string message)
        {
            Console.WriteLine(message);
        }

        private string bufferToString(byte[] buffer)
        {
            return System.Text.Encoding.Default.GetString(buffer);
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var buffer = new byte[port.ReadBufferSize];
            port.Read(buffer, 0, buffer.Length);
            print("new data: " + bufferToString(buffer));
            worksheet.Cell("A" + count).Value = System.Text.Encoding.Default.GetString(buffer);
            count++;
        }
    }
}