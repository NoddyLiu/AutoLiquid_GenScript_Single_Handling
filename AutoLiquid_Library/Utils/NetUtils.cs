using System;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;

namespace AutoLiquid_Library.Utils
{
    public class NetUtils
    {
        /// <summary>
        /// 测试网络是否连通
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public static bool TestNet(string ip)
        {
            Ping ping = new Ping();
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    if (ping.Send(ip, 500).Status == IPStatus.Success) return true;
                }
                catch (PingException e)
                {
                    return false;
                }
            }
            return false;
        }
    }
}