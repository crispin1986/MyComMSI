using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyComMSI
{
    public class MyComMSI
    {
        // Public COM interface
        [Guid("EAA4976A-45C3-4BC5-BC0B-E474F4C3C83F")]
        [InterfaceType(ComInterfaceType.InterfaceIsDual)]
        [ComVisible(true)]
        public interface IMyComComponent
        {
            string SayHello(string name);
            int AddNumbers(int a, int b);
        }

        [Guid("7BD20046-DF8C-44A6-AF72-3D39F2D3E8A8")]      // Unique GUID for class
        [ClassInterface(ClassInterfaceType.None)]         // Required for COM
        [ComVisible(true)]
        [ProgId("MyComMSI.MyComComponent")]  // <-- This is the ProgID
        public class MyComComponent : IMyComComponent
        {
            public string SayHello(string name)
            {
                return $"Hello, {name}! This message is from a COM component.";
            }

            public int AddNumbers(int a, int b)
            {
                return a + b;
            }
        }
    }
}
