using System;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;

namespace MockPassport
{
    public class Socket
    {
        public static void Start()
        {
            RunServer();
        }
        
        private static async void RunServer()
        {
            await Task.Run(() =>
            {
                var ipAddr = Dns.GetHostEntry("localhost").AddressList[0];
                var localEndPoint = new IPEndPoint(ipAddr, 443);

                var listener = new System.Net.Sockets.Socket(ipAddr.AddressFamily,
                    SocketType.Stream, ProtocolType.Tcp);

                try
                {
                    listener.Bind(localEndPoint);
                    listener.Listen(10);

                    while (true)
                    {
                        var clientSocket = listener.Accept();

                        clientSocket.Shutdown(SocketShutdown.Both);
                        clientSocket.Close();
                    }
                }

                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            });
        }
    }
}
