using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;

namespace ACME
{
    class global_Solar
    {

        public static string SolarPath  = @"\\acmesrv01\Public\SolarProject\";
        public static string globalPrjCode;
        public static SqlConnection SapConnection;
    }
}
