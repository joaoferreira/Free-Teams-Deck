using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FreeDeck
{
    class FreeDeckWindows
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow(); 
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [return: MarshalAs(UnmanagedType.Bool)]
        

        SerialPort sp;
        Boolean init = false;

        static void Main(string[] args)
        {
            Console.Title = "Free Teams Deck";
            //Open the Program function
            new FreeDeckWindows();

        }


        private FreeDeckWindows()
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;

            Console.WriteLine("                                             ");
            Console.WriteLine("                 &&&&&&&&&&                  ");
            Console.WriteLine("                &&&      &&&    &&&&&&       Microsoft Teams Deck");
            Console.WriteLine("                &&&      &&&   &&    &&      ");
            Console.WriteLine("  &&&&&&&&&&&&&&&&&&&&&&&&&     &&&&&&       Control Microsoft Teams from your Free Touch Deck");
            Console.WriteLine(" &&&&           &&&&&                        #FreeTeamsDeck");
            Console.WriteLine(" &&&&           &&&&&&&&&&&&&&&&&&&&&&&&     ");
            Console.WriteLine(" &&&&&&&&   &&&&&&&&&        &&        &&    ");
            Console.WriteLine(" &&&&&&&&   &&&&&&&&&        &&        &&    ");
            Console.WriteLine(" &&&&&&&&   &&&&&&&&&        &&        &&    ");
            Console.WriteLine(" &&&&&&&&   &&&&&&&&&        &&        &&    Application developed by:");
            Console.WriteLine("  &&&&&&&&&&&&&&&&&&         &&       &&     João Ferreira");
            Console.WriteLine("            &&&               &&&&&&&&&      https://teams.handsnontek.net/teamsdeck");
            Console.WriteLine("             &&&            &&&              ");
            Console.WriteLine("               &&&&&&&&&&&&&&                ");
            Console.WriteLine("                                             ");


            Console.ResetColor();

            //Serial communication parameters
            Console.Write(" COM Port Number: ");
            string comPort = Console.ReadLine();
            Console.Write(" Bound Rate: ");
            int boundRate = int.Parse(Console.ReadLine());
            Console.Write(" Status: ");
            


            //Set the datareceived event handler
            try
            {
                sp = new SerialPort("COM" + comPort, boundRate, Parity.None, 8, StopBits.One);
                sp.DataReceived += new SerialDataReceivedEventHandler(sp_DataReceived);
                //Open the serial port
                sp.Open();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Connected");
                Console.ResetColor();
                Console.WriteLine("");
                Console.Write(" To control Microsoft Teams when the application is not in focus, keep this window open");
            }
            catch(Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Unable to connect to Free Touch Deck");
                Console.ResetColor();
                Console.WriteLine("");
                Console.Write(" To learn how to connect to Free Touch Deck visit - https://teams.handsnontek.net/teamsdeck");
            }
            


            //Read from the console, to stop it from closing.
            Console.Read();
            
        }


        private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            IntPtr currentWindow = GetForegroundWindow();      

            try
            {
                //Variables
                string serialValue = sp.ReadLine();
                IntPtr teamsWindow = IntPtr.Zero;
                string menu = serialValue.Replace("\r", "");       


                //Get the main Teams process of Microsoft Teams
                Process[] processes = Process.GetProcessesByName("Teams");
                for (int i = 0; i < processes.Length; i++)
                {
                    if (processes[i].MainWindowHandle.ToString() != "0")
                    {
                        teamsWindow = processes[i].MainWindowHandle;  
                        if(currentWindow == teamsWindow)
                        {
                            //Teams is already the Window in focus
                            return;
                        }
                    }
                }

                

                //Loop through the menu options to make and send the keys
                switch (menu)
                {
                    case "Menu2":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        break;
                    case "Menu2-0":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        SendKeys.SendWait("^+(M)");
                        break;
                    case "Menu2-1":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        SendKeys.SendWait("^+(O)");
                        break;
                    case "Menu2-2":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        SendKeys.SendWait("^+(E)");
                        break;
                    case "Menu2-3":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        SendKeys.SendWait("^+(K)");
                        break;
                    case "Menu2-4":
                        ShowWindowAsync(teamsWindow, 1);
                        SetForegroundWindow(teamsWindow);
                        SendKeys.SendWait("^+(B)");
                        break;
                }       
  
            }
            catch (Exception ex)
            {
                //if there is no internet 
            }
        }
        
    }
}
