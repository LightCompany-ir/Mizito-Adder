using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace MizitoAdder
{
    public static class ShowMessage
    {
        public static void Clear()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("*** Light Company ***");
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void Welcome()
        {
            Console.Clear();
            Console.WriteLine("*** Welcome To New Application ***");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("**********************************");
            Console.WriteLine("****** Power By LightCompany *****");
            Console.WriteLine("**********************************");
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("Support : 011 44 44 60 44");
            Console.WriteLine("Website : LightCompany.ir");
        }
        public static void Waiter(int millisec = 3000)
        {
            Thread.Sleep(millisec);
            Application.DoEvents();
        }

        public static void Error(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void Success(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void Warning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void Message(string message)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(message);
        }
        public static void info(string message)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void Primary(string message)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
