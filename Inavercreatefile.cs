using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace navercreatefile
{
    [ComVisible(true),
    GuidAttribute("626A3D58-32EC-4B21-B8D6-7B5CD6763EB2")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    interface Inavercreatefile
    {
        string createFile(string accesskey, string timestamp, string signature, string fileName);
    }
}
