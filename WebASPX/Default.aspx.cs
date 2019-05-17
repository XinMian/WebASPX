using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebASPX
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void OnClickClickMeButton(object sender, EventArgs e)
        {
            string uniqeFileName = Guid.NewGuid().ToString() + ".xlsx";
            FileUpload1.SaveAs("D:\\TempFile\\" + uniqeFileName);
        }


    }
}
