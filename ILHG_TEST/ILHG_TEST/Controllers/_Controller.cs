using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ILHG_TEST.Controllers
{
    public class _Controller : Controller
    {
        public void OpenOdt(Byte[] buffer, string name)
        {
            Response.ContentType = "application/vnd.oasis.opendocument.text";
            Response.AddHeader("content-length", buffer.Length.ToString());
            Response.AddHeader("Content-Disposition", "attachment; filename=" + name);
            Response.BinaryWrite(buffer);
        }
    }

}