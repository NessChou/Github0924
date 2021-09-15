using System;
using System.IO;

namespace ACMELoggers
{
    public class Loggers
    {
        public static void log(Exception EMessage)
        {
            string txtFilePath = @"D:\log\log.txt";
            StreamWriter txtstreamwriter = null;
            try
            {
                if (!File.Exists(txtFilePath)) File.Create(txtFilePath);
                txtstreamwriter = new StreamWriter(txtFilePath, true, System.Text.UTF8Encoding.UTF8);
                txtstreamwriter.WriteLine(EMessage);
            }
            catch (Exception ex) { }
            finally
            {
                txtstreamwriter.Close();
            }
        }
    }
}
