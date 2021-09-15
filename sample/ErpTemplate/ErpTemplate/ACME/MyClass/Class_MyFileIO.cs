using System;
using System.Collections.Generic;
using System.Text;
using System.IO;            //�s�W�R�W�Ŷ�
using System.Windows.Forms; //�ϥ�Application.StartupPath�һݤޥΪ��R�W�Ŷ�
using System.Collections;   //

namespace My
{
    public class MyFileIO
    {

        #region ���o�ɮ׬�����T

        /// <summary>
        /// ���o�ɮ׬�����T
        /// </summary>
        /// <param name="FileName">�ɮצW��,�]�t���|</param>
        /// <param name="InfoStr">�d���ɮת��Ѽ�
        /// "Directory"      '���o�ؿ��W��
        /// "DirectoryName"  '���o������|�W��
        /// "CreationTime"   '���o�إ��ɮ׮ɶ�
        /// "Exists"         '�ˬd�ɮ׬O�_�s�b
        /// "Extension"      '���o���ɦW(�p: .DOC)
        /// "FullName"       '���o������|���ɮצW��
        /// "Name"           '���o�ɮצW��
        /// "Length"         '���o�ɮפj�p
        /// "LastAccessTime" '�W���s���ɶ�
        /// "LastWriteTime"  '�W���g�J�ɶ�
        /// </param>
        /// <returns>�^�Ǧr���</returns>
        public static string FileInformation(string FileName, string InfoStr)
        {
            FileInfo file1 = new FileInfo(FileName);

            switch (InfoStr)
            {
                case "Directory":   //���o�ؿ��W��
                    return file1.Directory.ToString();
                case "DirectoryName"://���o������|�W��
                    return file1.DirectoryName;
                case "CreationTime"://���o�إ��ɮ׮ɶ�
                    return file1.CreationTime.ToString();
                case "Exists"://�ˬd�ɮ׬O�_�s�b
                    return file1.Exists.ToString();
                case "Extension"://���o���ɦW(�p: .DOC)
                    return file1.Extension;
                case "FullName"://���o������|���ɮצW��
                    return file1.FullName;
                case "Name"://���o�ɮצW��
                    return file1.Name;
                case "Length"://���o�ɮפj�p
                    return file1.Length.ToString();
                case "LastAccessTime"://�W���s���ɶ�
                    return file1.LastAccessTime.ToString();
                case "LastWriteTime"://�W���g�J�ɶ�
                    return file1.LastWriteTime.ToString();
                default:
                    return "Error";
            }

        }


        #endregion


        #region �إ��ɮ�

        /// <summary>
        /// �إ��ɮ�
        /// </summary>
        /// <param name="FileName">�ɮצW��</param>
        /// <returns>�إ��ɮצ��\�^��True,�إ��ɮץ��Ѧ^��False</returns>
        public static bool FileCreate(string FileName)
        {
            //string filePath = Path.GetTempFileName();//�q�Ȧs���H�����ͤ@�ӼȦs��

            FileInfo file1 = new FileInfo(FileName);

            if (file1.Exists == false)
            {
                file1.Create();
                return true;
            }
            else
            {
                return false;//����ɮפw�g�s�b
            }

        }

        #endregion


        #region �ɮ׫���

        /// <summary>
        /// �ɮ׫���
        /// </summary>
        /// <param name="SourceFile">�ƻs�ӷ��ɮ�</param>
        /// <param name="TargetFile">�ƻs�ت��ɮ�</param>
        /// <returns>�ƻs���\�^��True , �ƻs���Ѧ^��False</returns>
        public static bool FileCopy(string SourceFile, string TargetFile)
        {
            FileInfo file1 = new FileInfo(SourceFile);
            try
            {
                file1.CopyTo(TargetFile);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("���~�T����:" + ex.Message.ToString());
                return false;
            }
        }

        #endregion


        #region �R���ɮ�

        /// <summary>
        /// �R���ɮ�
        /// </summary>
        /// <param name="FileName">��J�n�R�����ɮצW�٥]�t���|</param>
        /// <returns>�R�����\�^��True , �R�����Ѧ^��False</returns>
        public static bool FileDelete(string FileName)
        {
            FileInfo file1 = new FileInfo(FileName);

            if (file1.Exists == true)
            {
                file1.Delete();
                return true;
            }
            else
            {
                return false;//����ɮפ��s�b,�L�k�i���ɮקR���ʧ@
            }

        }

        #endregion


        #region ����ɦW

        /// <summary>
        /// ����ɦW
        /// </summary>
        /// <param name="OldFileName">���ɮצW��</param>
        /// <param name="NewFileName">����諸�s�ɮצW��</param>
        /// <returns></returns>
        public static bool FileRename(string OldFileName, string NewFileName)
        {
            FileInfo file1 = new FileInfo(OldFileName);
            FileInfo file2 = new FileInfo(NewFileName);

            try
            {
                if (file2.Exists == true)
                {
                    file2.Delete();
                }
                file1.MoveTo(NewFileName);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("���~�T����:" + ex.Message.ToString());
                return false;
            }
        }

        #endregion


        #region �Ǧ^�j�M�ؿ��U���Ҧ��ɮ�

        /// <summary>
        /// �Ǧ^�j�M�ؿ��U���Ҧ��ɮ�
        /// �Ϊk:SearchAllFiles(ref Alist, @"D:\WeiDa\src", "vb");
        /// </summary>
        /// <param name="SaveObj">�N�I�s�B�ҶǤJ��ArrayList�B�z������A�^�Ǧ^�h</param>
        /// <param name="SearchPath">�j�M�ؿ����|</param>
        /// <param name="SearchKeyword">�j�M�ɮצW������r</param>
        public static bool SearchAllFiles(ref ArrayList SaveObj, string SearchPath, string SearchKeyword)
        {
            DirectoryInfo DI = new DirectoryInfo(SearchPath);
            string[] Strbuf;
            int i;

            if (DI.Exists)
            {
                Strbuf = Directory.GetFileSystemEntries(SearchPath, "*" + SearchKeyword + "*");

                if (Strbuf.Length > 0)
                {
                    for (i = 0; i < Strbuf.Length; i++)
                    {
                        SaveObj.Insert(i, Strbuf[i]);
                    }

                    return true;
                }
                else
                {
                    return false;
                }

            }
            else
            {
                return false;
            }

        }

        #endregion

        
        #region �p��ؿ��U�Ҧ��ɮת��j�p

        /// <summary>
        /// �p��ؿ��U�Ҧ��ɮת��j�p
        /// </summary>
        /// <param name="DirPath">���p��Ҧ��ɮפj�p���ؿ����|</param>
        /// <returns>�^�Ǹ�ƫ��A��Double,�Y�^�ǭȬ�0�h��ܰ���p��L�{���o�Ϳ��~ </returns>
        public static double CountDirAllFilesSize(string DirPath)
        {
            DirectoryInfo DI = new DirectoryInfo(DirPath);
            string[] Strbuf;
            int i;
            double AllFileSize = 0;

            if (DI.Exists)
            {
                Strbuf = Directory.GetFileSystemEntries(DirPath);

                if (Strbuf.Length > 0)
                {
                    for (i = 0; i < Strbuf.Length; i++)
                    {
                        FileInfo file1 = new FileInfo(Strbuf[i]);

                        if (file1.Exists == true)//�ɮצs�b�~�i���ɮפj�p�p��
                        {
                            AllFileSize += Convert.ToDouble(FileInformation(Strbuf[i], "Length"));
                        }

                    }

                    return AllFileSize;
                }
                else
                {
                    return 0;
                }

            }
            else
            {
                return 0;
            }


        }

        #endregion


        #region �NByte�ഫ��Bit��KB��MB��GB��TB

        /// <summary>
        /// �NByte�ഫ��Bit��KB��MB��GB��TB
        /// </summary>
        /// <param name="SpaceSize">�Ŷ��j�p,��쬰Byte</param>
        /// <param name="TransferType">�ഫ����,��Ѽ�Bit,KB,MB,GB,TB</param>
        /// <returns>�^���ഫ���G�䫬�A��double</returns>
        public static double ByteToKBMBGBTB(double SpaceSize, string TransferType)
        {
            double result = 0;

            switch (TransferType)
            {
                case "Bit":
                    result = SpaceSize * 8;
                    break;
                case "KB": //Kilo Bytes
                    result = Convert.ToDouble(string.Format("{0:F2}", (SpaceSize / 1024)));
                    break;
                case "MB": //Mega Bytes
                    result = Convert.ToDouble(string.Format("{0:F2}", (SpaceSize / (1048576))));
                    break;
                case "GB": //Giga Bytes
                    result = Convert.ToDouble(string.Format("{0:F2}", (SpaceSize / (1073741824))));
                    break;
                case "TB": //Tera Bytes
                    result = Convert.ToDouble(string.Format("{0:F2}", (SpaceSize / (1099511627776))));
                    break;
                default:
                    break;
            }
            return result;

        }


        #endregion


        #region �إߥؿ��W��

        /// <summary>
        /// �إߥؿ��W��
        /// </summary>
        /// <param name="DirName">�ؿ��W��,�Ҧp:C:\TEMP</param>
        /// <returns>�Y�ؿ����\�إ߫h�^��True,�Y�ؿ��إߥ��ѫh�^��False</returns>
        public static bool DirCreate(string DirName)
        {
            DirectoryInfo dir1 = new DirectoryInfo(DirName);

            if (dir1.Exists == false)
            {
                dir1.Create();
                return true;
            }
            else
            {
                return false;//�ؿ��w�g�s�b
            }

        }

        #endregion



        #region ���ؿ��W��

        /// <summary>
        /// ���ؿ��W��
        /// </summary>
        /// <param name="OldDirName">�¥ؿ��W��</param>
        /// <param name="NewDirName">�s�ؿ��W��</param>
        /// <returns>�Y�ؿ����W�٦��\�إ߫h�^��True,�Y�ؿ����W�٥��ѫh�^��False</returns>
        public static bool DirRename(string OldDirName, string NewDirName)
        {
            DirectoryInfo dir1 = new DirectoryInfo(OldDirName);
            DirectoryInfo dir2 = new DirectoryInfo(NewDirName);
            try
            {
                if (dir2.Exists)
                {
                    return false;
                }
                else
                {
                    dir1.MoveTo(NewDirName);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("�ؿ���W����,����~�T����:" + ex.Message.ToString());
                return false;
            }

        }

        #endregion



        #region �R���ؿ�,�åB�M�w�ӥؿ��U�ɮ׬O�_�j��R��

        /// <summary>
        /// �R���ؿ�,�åB�M�w�ӥؿ��U�ɮ׬O�_�j��R��
        /// </summary>
        /// <param name="DirName">�ؿ����|</param>
        /// <param name="IsDelInDirFiles">�Y�n�j��ӥؿ��U�ɮץ����R���h�ǤJTrue,�_�h�ǤJFalse</param>
        /// <returns>�Y�ؿ��R�����\�h�^��True,�Y�ؿ��R�����ѫh�^��False</returns>
        public static bool DirDelete(string DirName, bool IsDelInDirFiles)
        {
            DirectoryInfo dir1 = new DirectoryInfo(DirName);

            try
            {
                if (dir1.Exists == false)
                {
                    return false;
                }
                else
                {
                    if (IsDelInDirFiles)
                    {
                        string[] bufFilePath = Directory.GetFileSystemEntries(DirName);

                        for (int i = 0; i < bufFilePath.Length; i++)
                        {
                            FileInfo file1 = new FileInfo(bufFilePath[i]);
                            file1.Delete();
                        }
                    }

                    dir1.Delete();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("�ؿ��R������,����~�T����:" + ex.Message.ToString());
                return false;
            }

        }

        #endregion



        #region �d�ߺϺЬ�����T

        /// <summary>
        /// �d�ߺϺЬ�����T
        /// �Ϊk�GDriveInformation("C:\\", "FileSystem").ToString()
        /// </summary>
        /// <param name="Drive">�ǤJ�n�d�ߪ��Ϻо�,�Ҧp:C:\\</param>
        /// <param name="InfoType">�d�߰Ѽ�,�]�t�G
        /// "VolumeLabel"://�ϺмаO
        /// "FileSystem"://�ɮרt��
        /// "TotalFreeSpace"://���ĪŶ��`�q
        /// "TotalSize"://�Ϻо����j�p
        /// </param>
        /// <returns></returns>
        public static object DriveInformation(string Drive, string InfoType)
        {
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            object result;
            if (allDrives.Length > 0)
            {
                foreach (DriveInfo d in allDrives)
                {

                    if (d.DriveType.ToString() == "Fixed" && d.IsReady == true && d.Name == Drive)
                    {
                        switch (InfoType)
                        {
                            case "VolumeLabel"://�ϺмаO
                                result = d.VolumeLabel;
                                break;
                            case "FileSystem"://�ɮרt��
                                result = d.DriveFormat;
                                break;
                            case "TotalFreeSpace"://���ĪŶ��`�q
                                result = d.TotalFreeSpace;
                                break;
                            case "TotalSize"://�Ϻо����j�p
                                result = d.TotalSize;
                                break;
                            default:
                                result = false;
                                break;
                        }
                        return result;
                    }

                }
                result = false;
                return result;

            }
            else
            {
                result = false;
                return result;
            }


        }

        #endregion

    }
}
