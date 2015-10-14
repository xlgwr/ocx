using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;

namespace TestOCX
{
    //[Guid("7cbebdca-776e-4238-9191-1554c6872b8b")]
    public class TestOCX  //:ActiveXControl
    {
        public string GetMacAddress()
        {
            var mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            var mos = mc.GetInstances();
            var sb = new StringBuilder();

            foreach (ManagementObject mo in mos)
            {
                var macAddress = mo["MacAddress"];

                if (macAddress != null)
                    sb.AppendLine(macAddress.ToString());
            }

            return sb.ToString();
        }
    }
}
