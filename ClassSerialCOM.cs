using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace MyProject
{
    internal class SerialCOM
    {
        static SerialPort serialPort = new SerialPort();

        public void OpenCommunication()
        {
            try
            {
                serialPort.PortName = AutodetectArduinoPort();
                serialPort.BaudRate = 9600;
                serialPort.DataBits = 8;
                serialPort.Parity = Parity.None;
                serialPort.Handshake = Handshake.None;
                serialPort.ReadTimeout = 500;
                serialPort.WriteTimeout = 500;

                serialPort.Open();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public SerialCOM()
        {
            OpenCommunication();
        }

        public string Read()
        {
            string readData = "";

            if (!serialPort.IsOpen)
                OpenCommunication();

            try
            {
                serialPort.DiscardInBuffer();
                readData = serialPort.ReadLine();
                if (readData == null)
                {
                    readData = "readData is null";
                    // Console.WriteLine("readData is null");
                }
                readData = readData.Replace("\r", string.Empty);
                //Console.WriteLine(readData);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }

            return readData;
        }

        public void Write(string data)
        {
            if (!serialPort.IsOpen)
                OpenCommunication();

            try
            {
                serialPort.WriteLine(data);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }

        }

        public string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("USB Serial Device"))
                    {
                        return deviceId;
                    }
                }
            }
            catch
            {
                /* Do Nothing */
            }

            return null;
        }

    }
}
