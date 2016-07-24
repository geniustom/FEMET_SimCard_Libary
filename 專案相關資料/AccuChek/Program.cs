using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;

namespace AccuChek
{
    class Program
    {
        static private byte[] sendToAccuChek( byte[] command)
        {
            byte[] sendByte = command;
            List<byte> recvByte = new List<byte>();

            using (SerialPort sp = new SerialPort(COM, 9600, Parity.None, 8, StopBits.One))
            {
                sp.DtrEnable = true;
                sp.RtsEnable = false;
                sp.Open();
                int Cnt = 0;
                foreach (byte b in sendByte)
                {
                    Thread.Sleep(2); // 送值需等待 2ms
                    sp.Write(sendByte, Cnt++, 1);
                    Console.Write("Send:{0:X} ", b);
                    Thread.Sleep(10); //接收值需等待 10ms

                    while (true)
                    {
                        if (sp.BytesToRead > 0)
                        {
                            Console.Write("Recv:");
                            while (sp.BytesToRead > 0)
                            {
                                int i = sp.ReadByte();
                                recvByte.Add((byte)i);
                                Console.Write("{0:X}  ", i);
                                Thread.Sleep(1);
                            }
                            Console.WriteLine();
                            break;
                        }

                        Thread.Sleep(10);//接收值需等待 10ms

                    }
                    //}
                }
                Console.WriteLine();
                sp.Close();
            }
            return recvByte.ToArray();
        }
        /// <summary>
        /// 清空 機器的 Error message
        /// </summary>
        /// <returns></returns>
        static private void clearErrorStatus()
        {
           
            byte[] data = sendToAccuChek( CLEAR_STATUS);
            Thread.Sleep(sleepMicroSecond);
            try
            {
                while (data[6] != 48 || data[7] != 48 || data[8] != 48 || data[9] != 48)
                {
                    data = sendToAccuChek( CLEAR_STATUS);
                    break;
                }
            }
            catch
            {
                //若機器對清值指令傳回NAK時，在重新呼叫清除指令
                clearErrorStatus();
            }         
        }
        static private void readMemory(int startIndex,int endIndex)
        {
            READ_MEMORY = new byte[] { 0x61, 0x09, 0x00,0x09,0x00, 0x0D, 0x06 };
            

            for (int i = startIndex; i <= endIndex; i++)
            {
                READ_MEMORY[2] =(byte) (48 + i); //+48是因為要轉成ascii的 數字
                READ_MEMORY[4] =(byte) (48 + i); //+48是因為要轉成ascii的 數字

               
                clearErrorStatus();
                Thread.Sleep(sleepMicroSecond);
                Console.WriteLine("讀取 第 {0} 筆資料",i);
                sendToAccuChek(READ_MEMORY);
                Thread.Sleep(sleepMicroSecond);
                Console.WriteLine("///////////////////////////////////");
            }


        }

        static private int readMemoryLength()
        {
            int Cnt =0;
           
            clearErrorStatus();

            Console.WriteLine("讀取memory 筆數");
            byte[] recvData = sendToAccuChek(READ_MEMORY_LEN);
            if(recvData.Length == 14) // 正常執行完成會送回14個byte           
            {
                // 6 7 8 為長度傳回的位址 位元傳回可參考文件
                Cnt = (recvData[6] - 48) * 100 + (recvData[7] - 48) * 10 + (recvData[8] - 48);                
            }
            Console.WriteLine("memory 共有 {0} 筆資料", Cnt);
            Thread.Sleep(sleepMicroSecond);
            Console.WriteLine("///////////////////////////////////");
            return Cnt;
        }

        static string COM = "COM5";
        static int sleepMicroSecond = 1000; //可自行調整等時間

        static byte[] CLEAR_STATUS = new byte[] { 0x0B, 0x0D, 0x06 };
        static byte[] READ_DATE = new byte[] { 0x53, 0x09, 0x31, 0x0D, 0x06 };
        static byte[] READ_TIME = new byte[] { 0x53, 0x09, 0x32, 0x0D, 0x06 };
        static byte[] READ_UNITS = new byte[] { 0x53, 0x09, 0x33, 0x0D, 0x06 };
        static byte[] READ_DATA = new byte[] { 0x53, 0x09, 0x78,0x09,0x31,0x0D, 0x06 };
        static byte[] READ_SN = new byte[] { 0x43, 0x09, 0x33, 0x0D, 0x06 };
        static byte[] READ_MODEL = new byte[] { 0x43, 0x09, 0x34, 0x0D, 0x06 };
        static byte[] READ_SW_VER = new byte[] { 0x43, 0x09, 0x31, 0x0D, 0x06 };
        /// <summary>
        /// Data sheet page 47 
        /// READ_MEMORY[2] = 從第 ? 筆開始
        /// READ_MEMORY[4] = 到第 ? 等結束
        /// </summary>
        static byte[] READ_MEMORY = new byte[] { 0x61, 0x09, 0x00,0x09,0x00, 0x0D, 0x06 };
        /// <summary>
        /// Data sheet page 47 
        /// </summary>
        static byte[] READ_MEMORY_LEN = new byte[] { 0x60, 0x0D, 0x06 };
        //Read Language
        static void Main(string[] args)
        {
            
            if (args.Length == 1)
            {
                COM = args[0];
            }
            int length = readMemoryLength();
            if(length >0)
            readMemory(1, length);


           
          
            Console.ReadLine();
        }
        
    }
}
