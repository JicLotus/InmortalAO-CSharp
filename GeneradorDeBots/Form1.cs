using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeneradorDeBots
{
    public partial class Form1 : Form
    {

        List<TcpClient> connections; 
        public Form1()
        {
            InitializeComponent();
            
        }

        private void createConnections()
        {
            Int32 port = 7777;
            string server = "localhost";
            int countConnections = 500;
            connections = new List<TcpClient>();

            for (int i=0;i<countConnections;i++)
            {
                connections.Add(new TcpClient(server, port));
                connectRandomClient(connections[i]);
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (TcpClient client in connections) {
                connectRandomClient(client);
            }
        }

        private void connectRandomClient(TcpClient client)
        {
            string input = RandomString(9);
            byte[] array = Encoding.ASCII.GetBytes(input);

            List<byte> data = new List<byte>();
            data.AddRange(new byte[] { 3, Convert.ToByte(input.Length), 0 });
            data.AddRange(array);
            data.AddRange(new byte[] { 1, 0, 98, 1, 1, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 18, 18, 10, 6, 18, 1, 0, 0, 1, 0 });
            NetworkStream stream = client.GetStream();

            stream.Write(data.ToArray(), 0, data.ToArray().Length);
            Thread.Sleep(100);
        }


        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            createConnections();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            while (true) { 
                foreach (TcpClient client in connections)
                {
                    NetworkStream stream = client.GetStream();
                    byte[] data = { 6, 1 };
                    stream.Write(data.ToArray(), 0, data.ToArray().Length);
                }

                Thread.Sleep(300);

                foreach (TcpClient client in connections)
                {
                    NetworkStream stream = client.GetStream();
                    byte[] data = { 6, 3 };
                    stream.Write(data.ToArray(), 0, data.ToArray().Length);
                }

                Thread.Sleep(1000);
            }
        }
    }
}
