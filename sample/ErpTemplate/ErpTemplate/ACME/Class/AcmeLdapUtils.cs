using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;


//
using System.DirectoryServices;
using System.Runtime.InteropServices;


/// <summary>
/// AcmeLdapUtils 的摘要描述
/// </summary>
public class AcmeLdapUtils
{
	public AcmeLdapUtils()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

    public static string GetLDAPPath()
    {
     //   return "LDAP://serv.acmepoint.com";

        return "LDAP://10.10.1.56";
    }


    public static bool CheckAdUser(string UserName, string Password)
    {
        string strPath = GetLDAPPath();

        DirectoryEntry entry = null;

        try
        {

            entry = new DirectoryEntry(strPath, UserName, Password);

        }
        catch (COMException Ex)
        {
            //MessageBox.Show("Couldnt connect to the specified Active Directory Path" + "\n" + "Error = " + Ex.Message + Ex.InnerException,
            //    "AD Tester");
            return false;

        }

        DirectorySearcher mySearcher = new DirectorySearcher(entry);

        TimeSpan waitTime;

        try
        {
            waitTime = new TimeSpan(0, 0, Convert.ToInt32(60));
            mySearcher.ClientTimeout = waitTime;
        }
        catch (Exception Ex)
        {
            //MessageBox.Show("Error = " + Ex.Message + Ex.InnerException,
            //    "AD Tester");
            return false;
        }

        string strCat = "(objectCategory=user)";
        mySearcher.Filter = strCat;

        try
        {
            SearchResult result = mySearcher.FindOne();

            //if exception was not thrown, means it connected successfuly
            //MessageBox.Show("Successfuly connected to Active Directory",
            //    "AD Tester");

            return true;

        }
        catch (COMException Ex)
        {
            //MessageBox.Show("Couldnt connect to Active Directory" + "\n" + "Error = " + Ex.Message + Ex.InnerException,
            //    "AD Tester");
            //this.Cursor = Cursors.Arrow; //change back the cursor to normal
            return false;
        }
        catch (Exception Ex)
        {
            //MessageBox.Show("Error = " + Ex.Message + Ex.InnerException,
            //    "AD Tester");
            //this.Cursor = Cursors.Arrow; //change back the cursor to normal
            return false;

        }


    }
}
