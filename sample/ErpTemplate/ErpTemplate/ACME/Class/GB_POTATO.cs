  using System;
  using System.Data;
  using System.Configuration;
  using System.Web;
  using System.Data.SqlClient;
  
  /// <summary>
  /// Summary description for GB_POTATO
  /// 作者:
  /// </summary>
 /// GB_POTATO 資料結構
namespace ACME
{
        
    public class GB_POTATO
    {

        private Int32 _ID;
        private String _GBTYPE;
        private String _OrdName;
        private String _OrdTel;
        private String _OrdCom;
        private String _OrdEMail;
        private String _DelMan;
        private String _DelTel;
        private String _DelAddr;
        private String _ProdID;
        private String _ProdName;
        private Int32 _Qty;
        private Int32 _Price;
        private Int32 _Amount;
        private String _Remark;
        private String _DocType;
        private String _OrdNo;
        private String _InvNo;
        private String _UserCode;
        private String _CreateUser;
        private String _CreateDate;
        private String _CreateTime;
        private String _UpdateUser;
        private String _UpdateDate;
        private String _UpdateTime;
        private String _ShipDate;
        private String _WorkDate;
        private String _PotatoKind;
        private Int32 _BoxQty;
        private Int32 _PotatoWg;
        private String _TransMark;
        private String _DelRemark;
        private String _Flag1;
        private String _Flag2;
        private String _DelCom;
        private String _DELDATE;
        private String _UNIT;
        private String _SERV;
        private String _MEMO;
        private String _SHIPFEE;
        private String _PAYMAN;
        //SHIPFEE
        //PAYMAN
        public Int32 ID { get { return _ID; } set { _ID = value; } }
        public String GBTYPE { get { return _GBTYPE; } set { _GBTYPE = value; } }
        public String OrdName { get { return _OrdName; } set { _OrdName = value; } }
        public String OrdTel { get { return _OrdTel; } set { _OrdTel = value; } }
        public String OrdCom { get { return _OrdCom; } set { _OrdCom = value; } }
        public String OrdEMail { get { return _OrdEMail; } set { _OrdEMail = value; } }
        public String DelMan { get { return _DelMan; } set { _DelMan = value; } }
        public String DelTel { get { return _DelTel; } set { _DelTel = value; } }
        public String DelAddr { get { return _DelAddr; } set { _DelAddr = value; } }
        public String ProdID { get { return _ProdID; } set { _ProdID = value; } }
        public String ProdName { get { return _ProdName; } set { _ProdName = value; } }
        public Int32 Qty { get { return _Qty; } set { _Qty = value; } }
        public Int32 Price { get { return _Price; } set { _Price = value; } }
        public Int32 Amount { get { return _Amount; } set { _Amount = value; } }
        public String Remark { get { return _Remark; } set { _Remark = value; } }
        public String DocType { get { return _DocType; } set { _DocType = value; } }
        public String OrdNo { get { return _OrdNo; } set { _OrdNo = value; } }
        public String InvNo { get { return _InvNo; } set { _InvNo = value; } }
        public String UserCode { get { return _UserCode; } set { _UserCode = value; } }
        public String CreateUser { get { return _CreateUser; } set { _CreateUser = value; } }
        public String CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
        public String CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
        public String UpdateUser { get { return _UpdateUser; } set { _UpdateUser = value; } }
        public String UpdateDate { get { return _UpdateDate; } set { _UpdateDate = value; } }
        public String UpdateTime { get { return _UpdateTime; } set { _UpdateTime = value; } }
        public String ShipDate { get { return _ShipDate; } set { _ShipDate = value; } }
        public String WorkDate { get { return _WorkDate; } set { _WorkDate = value; } }
        public String PotatoKind { get { return _PotatoKind; } set { _PotatoKind = value; } }
        public Int32 BoxQty { get { return _BoxQty; } set { _BoxQty = value; } }
        public Int32 PotatoWg { get { return _PotatoWg; } set { _PotatoWg = value; } }
        public String TransMark { get { return _TransMark; } set { _TransMark = value; } }
        public String DelRemark { get { return _DelRemark; } set { _DelRemark = value; } }
        public String Flag1 { get { return _Flag1; } set { _Flag1 = value; } }
        public String Flag2 { get { return _Flag2; } set { _Flag2 = value; } }
        public String DelCom { get { return _DelCom; } set { _DelCom = value; } }
        public String DELDATE { get { return _DELDATE; } set { _DELDATE = value; } }
        public String UNIT { get { return _UNIT; } set { _UNIT = value; } }
        public String SERV { get { return _SERV; } set { _SERV = value; } }
        public String MEMO { get { return _MEMO; } set { _MEMO = value; } }
        public String SHIPFEE { get { return _SHIPFEE; } set { _SHIPFEE = value; } }
        public String PAYMAN { get { return _PAYMAN; } set { _PAYMAN = value; } }
        public GB_POTATO()
        {
        }
        // GB_POTATO Insert
        public static Int32 AddGB_POTATO(GB_POTATO row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_POTATO(GBTYPE,OrdName,OrdTel,OrdCom,OrdEMail,DelMan,DelTel,DelAddr,ProdID,ProdName,Qty,Price,Amount,Remark,DocType,OrdNo,InvNo,UserCode,CreateUser,CreateDate,CreateTime,UpdateUser,UpdateDate,UpdateTime,ShipDate,WorkDate,PotatoKind,BoxQty,PotatoWg,TransMark,DelRemark,Flag1,Flag2,DelCom,DELDATE,UNIT,SERV,MEMO,SHIPFEE,PAYMAN) values(@GBTYPE,@OrdName,@OrdTel,@OrdCom,@OrdEMail,@DelMan,@DelTel,@DelAddr,@ProdID,@ProdName,@Qty,@Price,@Amount,@Remark,@DocType,@OrdNo,@InvNo,@UserCode,@CreateUser,@CreateDate,@CreateTime,@UpdateUser,@UpdateDate,@UpdateTime,@ShipDate,@WorkDate,@PotatoKind,@BoxQty,@PotatoWg,@TransMark,@DelRemark,@Flag1,@Flag2,@DelCom,@DELDATE,@UNIT,@SERV,@MEMO,@SHIPFEE,@PAYMAN);SELECT @@IDENTITY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@GBTYPE", SqlDbType.VarChar, 10, "GBTYPE"));
            command.Parameters["@GBTYPE"].Value = row.GBTYPE;
            if (String.IsNullOrEmpty(row.GBTYPE))
            {
                command.Parameters["@GBTYPE"].IsNullable = true;
                command.Parameters["GBTYPE"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdName", SqlDbType.NVarChar, 50, "OrdName"));
            command.Parameters["@OrdName"].Value = row.OrdName;
            if (String.IsNullOrEmpty(row.OrdName))
            {
                command.Parameters["@OrdName"].IsNullable = true;
                command.Parameters["@OrdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdTel", SqlDbType.NVarChar, 50, "OrdTel"));
            command.Parameters["@OrdTel"].Value = row.OrdTel;
            if (String.IsNullOrEmpty(row.OrdTel))
            {
                command.Parameters["@OrdTel"].IsNullable = true;
                command.Parameters["@OrdTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdCom", SqlDbType.NVarChar, 100, "OrdCom"));
            command.Parameters["@OrdCom"].Value = row.OrdCom;
            if (String.IsNullOrEmpty(row.OrdCom))
            {
                command.Parameters["@OrdCom"].IsNullable = true;
                command.Parameters["@OrdCom"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdEMail", SqlDbType.NVarChar, 50, "OrdEMail"));
            command.Parameters["@OrdEMail"].Value = row.OrdEMail;
            if (String.IsNullOrEmpty(row.OrdEMail))
            {
                command.Parameters["@OrdEMail"].IsNullable = true;
                command.Parameters["@OrdEMail"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelMan", SqlDbType.NVarChar, 50, "DelMan"));
            command.Parameters["@DelMan"].Value = row.DelMan;
            if (String.IsNullOrEmpty(row.DelMan))
            {
                command.Parameters["@DelMan"].IsNullable = true;
                command.Parameters["@DelMan"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelTel", SqlDbType.NVarChar, 50, "DelTel"));
            command.Parameters["@DelTel"].Value = row.DelTel;
            if (String.IsNullOrEmpty(row.DelTel))
            {
                command.Parameters["@DelTel"].IsNullable = true;
                command.Parameters["@DelTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelAddr", SqlDbType.NVarChar, 50, "DelAddr"));
            command.Parameters["@DelAddr"].Value = row.DelAddr;
            if (String.IsNullOrEmpty(row.DelAddr))
            {
                command.Parameters["@DelAddr"].IsNullable = true;
                command.Parameters["@DelAddr"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdID", SqlDbType.NVarChar, 50, "ProdID"));
            command.Parameters["@ProdID"].Value = row.ProdID;
            if (String.IsNullOrEmpty(row.ProdID))
            {
                command.Parameters["@ProdID"].IsNullable = true;
                command.Parameters["@ProdID"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdName", SqlDbType.NVarChar, 200, "ProdName"));
            command.Parameters["@ProdName"].Value = row.ProdName;
            if (String.IsNullOrEmpty(row.ProdName))
            {
                command.Parameters["@ProdName"].IsNullable = true;
                command.Parameters["@ProdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Qty", row.Qty));
            command.Parameters.Add(new SqlParameter("@Price", row.Price));
            command.Parameters.Add(new SqlParameter("@Amount", row.Amount));
            command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.NVarChar, 250, "Remark"));
            command.Parameters["@Remark"].Value = row.Remark;
            if (String.IsNullOrEmpty(row.Remark))
            {
                command.Parameters["@Remark"].IsNullable = true;
                command.Parameters["@Remark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DocType", SqlDbType.NVarChar, 20, "DocType"));
            command.Parameters["@DocType"].Value = row.DocType;
            if (String.IsNullOrEmpty(row.DocType))
            {
                command.Parameters["@DocType"].IsNullable = true;
                command.Parameters["@DocType"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdNo", SqlDbType.NVarChar, 20, "OrdNo"));
            command.Parameters["@OrdNo"].Value = row.OrdNo;
            if (String.IsNullOrEmpty(row.OrdNo))
            {
                command.Parameters["@OrdNo"].IsNullable = true;
                command.Parameters["@OrdNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@InvNo", SqlDbType.NVarChar, 20, "InvNo"));
            command.Parameters["@InvNo"].Value = row.InvNo;
            if (String.IsNullOrEmpty(row.InvNo))
            {
                command.Parameters["@InvNo"].IsNullable = true;
                command.Parameters["@InvNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.NVarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateUser", SqlDbType.NVarChar, 20, "CreateUser"));
            command.Parameters["@CreateUser"].Value = row.CreateUser;
            if (String.IsNullOrEmpty(row.CreateUser))
            {
                command.Parameters["@CreateUser"].IsNullable = true;
                command.Parameters["@CreateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.NVarChar, 8, "CreateDate"));
            command.Parameters["@CreateDate"].Value = row.CreateDate;
            if (String.IsNullOrEmpty(row.CreateDate))
            {
                command.Parameters["@CreateDate"].IsNullable = true;
                command.Parameters["@CreateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.NVarChar, 6, "CreateTime"));
            command.Parameters["@CreateTime"].Value = row.CreateTime;
            if (String.IsNullOrEmpty(row.CreateTime))
            {
                command.Parameters["@CreateTime"].IsNullable = true;
                command.Parameters["@CreateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.NVarChar, 20, "UpdateUser"));
            command.Parameters["@UpdateUser"].Value = row.UpdateUser;
            if (String.IsNullOrEmpty(row.UpdateUser))
            {
                command.Parameters["@UpdateUser"].IsNullable = true;
                command.Parameters["@UpdateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateDate", SqlDbType.NVarChar, 8, "UpdateDate"));
            command.Parameters["@UpdateDate"].Value = row.UpdateDate;
            if (String.IsNullOrEmpty(row.UpdateDate))
            {
                command.Parameters["@UpdateDate"].IsNullable = true;
                command.Parameters["@UpdateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateTime", SqlDbType.NVarChar, 6, "UpdateTime"));
            command.Parameters["@UpdateTime"].Value = row.UpdateTime;
            if (String.IsNullOrEmpty(row.UpdateTime))
            {
                command.Parameters["@UpdateTime"].IsNullable = true;
                command.Parameters["@UpdateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ShipDate", SqlDbType.NVarChar, 8, "ShipDate"));
            command.Parameters["@ShipDate"].Value = row.ShipDate;
            if (String.IsNullOrEmpty(row.ShipDate))
            {
                command.Parameters["@ShipDate"].IsNullable = true;
                command.Parameters["@ShipDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.NVarChar, 8, "WorkDate"));
            command.Parameters["@WorkDate"].Value = row.WorkDate;
            if (String.IsNullOrEmpty(row.WorkDate))
            {
                command.Parameters["@WorkDate"].IsNullable = true;
                command.Parameters["@WorkDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@PotatoKind", SqlDbType.NVarChar, 1, "PotatoKind"));
            command.Parameters["@PotatoKind"].Value = row.PotatoKind;
            if (String.IsNullOrEmpty(row.PotatoKind))
            {
                command.Parameters["@PotatoKind"].IsNullable = true;
                command.Parameters["@PotatoKind"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@BoxQty", row.BoxQty));
            command.Parameters.Add(new SqlParameter("@PotatoWg", row.PotatoWg));
            command.Parameters.Add(new SqlParameter("@TransMark", SqlDbType.NVarChar, 255, "TransMark"));
            command.Parameters["@TransMark"].Value = row.TransMark;
            if (String.IsNullOrEmpty(row.TransMark))
            {
                command.Parameters["@TransMark"].IsNullable = true;
                command.Parameters["@TransMark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelRemark", SqlDbType.NVarChar, 255, "DelRemark"));
            command.Parameters["@DelRemark"].Value = row.DelRemark;
            if (String.IsNullOrEmpty(row.DelRemark))
            {
                command.Parameters["@DelRemark"].IsNullable = true;
                command.Parameters["@DelRemark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Flag1", SqlDbType.NVarChar, 50, "Flag1"));
            command.Parameters["@Flag1"].Value = row.Flag1;
            if (String.IsNullOrEmpty(row.Flag1))
            {
                command.Parameters["@Flag1"].IsNullable = true;
                command.Parameters["@Flag1"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Flag2", SqlDbType.NVarChar, 50, "Flag2"));
            command.Parameters["@Flag2"].Value = row.Flag2;
            if (String.IsNullOrEmpty(row.Flag2))
            {
                command.Parameters["@Flag2"].IsNullable = true;
                command.Parameters["@Flag2"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelCom", SqlDbType.NVarChar, 100, "DelCom"));
            command.Parameters["@DelCom"].Value = row.DelCom;
            if (String.IsNullOrEmpty(row.DelCom))
            {
                command.Parameters["@DelCom"].IsNullable = true;
                command.Parameters["@DelCom"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DELDATE", SqlDbType.NVarChar, 50, "DELDATE"));
            command.Parameters["@DELDATE"].Value = row.DELDATE;
            if (String.IsNullOrEmpty(row.DelCom))
            {
                command.Parameters["@DELDATE"].IsNullable = true;
                command.Parameters["@DELDATE"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UNIT", SqlDbType.NVarChar, 50, "UNIT"));
            command.Parameters["@UNIT"].Value = row.UNIT;
            if (String.IsNullOrEmpty(row.UNIT))
            {
                command.Parameters["@UNIT"].IsNullable = true;
                command.Parameters["@UNIT"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@SERV", SqlDbType.NVarChar, 50, "SERV"));
            command.Parameters["@SERV"].Value = row.SERV;
            if (String.IsNullOrEmpty(row.SERV))
            {
                command.Parameters["@SERV"].IsNullable = true;
                command.Parameters["@SERV"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@MEMO", SqlDbType.NVarChar, 500, "MEMO"));
            command.Parameters["@MEMO"].Value = row.MEMO;
            if (String.IsNullOrEmpty(row.MEMO))
            {
                command.Parameters["@MEMO"].IsNullable = true;
                command.Parameters["@MEMO"].Value = "";
            }
            //SHIPFEE
            //PAYMAN
            command.Parameters.Add(new SqlParameter("@SHIPFEE", SqlDbType.NVarChar, 50, "SHIPFEE"));
            command.Parameters["@SHIPFEE"].Value = row.SHIPFEE;
            if (String.IsNullOrEmpty(row.SHIPFEE))
            {
                command.Parameters["@SHIPFEE"].IsNullable = true;
                command.Parameters["@SHIPFEE"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@PAYMAN", SqlDbType.NVarChar, 50, "PAYMAN"));
            command.Parameters["@PAYMAN"].Value = row.PAYMAN;
            if (String.IsNullOrEmpty(row.PAYMAN))
            {
                command.Parameters["@PAYMAN"].IsNullable = true;
                command.Parameters["@PAYMAN"].Value = "";
            }
            Int32 AutoNo = 0;
            try
            {
                connection.Open();

                AutoNo = Convert.ToInt32(command.ExecuteScalar());
            }
            finally
            {
                connection.Close();
            }
            return AutoNo;
        }

        // GB_POTATO Update
        public static void UpdateGB_POTATO(GB_POTATO row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE GB_POTATO SET OrdName=@OrdName,OrdTel=@OrdTel,OrdCom=@OrdCom,OrdEMail=@OrdEMail,DelMan=@DelMan,DelTel=@DelTel,DelAddr=@DelAddr,ProdID=@ProdID,ProdName=@ProdName,Qty=@Qty,Price=@Price,Amount=@Amount,Remark=@Remark,DocType=@DocType,OrdNo=@OrdNo,InvNo=@InvNo,UserCode=@UserCode,UpdateUser=@UpdateUser,UpdateDate=@UpdateDate,UpdateTime=@UpdateTime,ShipDate=@ShipDate,WorkDate=@WorkDate,PotatoKind=@PotatoKind,BoxQty=@BoxQty,PotatoWg=@PotatoWg,TransMark=@TransMark,DelRemark=@DelRemark,Flag1=@Flag1,Flag2=@Flag2,DelCom=@DelCom WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@OrdName", SqlDbType.VarChar, 50, "OrdName"));
            command.Parameters["@OrdName"].Value = row.OrdName;
            if (String.IsNullOrEmpty(row.OrdName))
            {
                command.Parameters["@OrdName"].IsNullable = true;
                command.Parameters["@OrdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdTel", SqlDbType.VarChar, 50, "OrdTel"));
            command.Parameters["@OrdTel"].Value = row.OrdTel;
            if (String.IsNullOrEmpty(row.OrdTel))
            {
                command.Parameters["@OrdTel"].IsNullable = true;
                command.Parameters["@OrdTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdCom", SqlDbType.VarChar, 100, "OrdCom"));
            command.Parameters["@OrdCom"].Value = row.OrdCom;
            if (String.IsNullOrEmpty(row.OrdCom))
            {
                command.Parameters["@OrdCom"].IsNullable = true;
                command.Parameters["@OrdCom"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdEMail", SqlDbType.VarChar, 50, "OrdEMail"));
            command.Parameters["@OrdEMail"].Value = row.OrdEMail;
            if (String.IsNullOrEmpty(row.OrdEMail))
            {
                command.Parameters["@OrdEMail"].IsNullable = true;
                command.Parameters["@OrdEMail"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelMan", SqlDbType.VarChar, 50, "DelMan"));
            command.Parameters["@DelMan"].Value = row.DelMan;
            if (String.IsNullOrEmpty(row.DelMan))
            {
                command.Parameters["@DelMan"].IsNullable = true;
                command.Parameters["@DelMan"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelTel", SqlDbType.VarChar, 50, "DelTel"));
            command.Parameters["@DelTel"].Value = row.DelTel;
            if (String.IsNullOrEmpty(row.DelTel))
            {
                command.Parameters["@DelTel"].IsNullable = true;
                command.Parameters["@DelTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelAddr", SqlDbType.VarChar, 50, "DelAddr"));
            command.Parameters["@DelAddr"].Value = row.DelAddr;
            if (String.IsNullOrEmpty(row.DelAddr))
            {
                command.Parameters["@DelAddr"].IsNullable = true;
                command.Parameters["@DelAddr"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdID", SqlDbType.VarChar, 50, "ProdID"));
            command.Parameters["@ProdID"].Value = row.ProdID;
            if (String.IsNullOrEmpty(row.ProdID))
            {
                command.Parameters["@ProdID"].IsNullable = true;
                command.Parameters["@ProdID"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdName", SqlDbType.VarChar, 200, "ProdName"));
            command.Parameters["@ProdName"].Value = row.ProdName;
            if (String.IsNullOrEmpty(row.ProdName))
            {
                command.Parameters["@ProdName"].IsNullable = true;
                command.Parameters["@ProdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Qty", row.Qty));
            command.Parameters.Add(new SqlParameter("@Price", row.Price));
            command.Parameters.Add(new SqlParameter("@Amount", row.Amount));
            command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.VarChar, 250, "Remark"));
            command.Parameters["@Remark"].Value = row.Remark;
            if (String.IsNullOrEmpty(row.Remark))
            {
                command.Parameters["@Remark"].IsNullable = true;
                command.Parameters["@Remark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DocType", SqlDbType.VarChar, 20, "DocType"));
            command.Parameters["@DocType"].Value = row.DocType;
            if (String.IsNullOrEmpty(row.DocType))
            {
                command.Parameters["@DocType"].IsNullable = true;
                command.Parameters["@DocType"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdNo", SqlDbType.VarChar, 20, "OrdNo"));
            command.Parameters["@OrdNo"].Value = row.OrdNo;
            if (String.IsNullOrEmpty(row.OrdNo))
            {
                command.Parameters["@OrdNo"].IsNullable = true;
                command.Parameters["@OrdNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@InvNo", SqlDbType.VarChar, 20, "InvNo"));
            command.Parameters["@InvNo"].Value = row.InvNo;
            if (String.IsNullOrEmpty(row.InvNo))
            {
                command.Parameters["@InvNo"].IsNullable = true;
                command.Parameters["@InvNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
            command.Parameters["@UpdateUser"].Value = row.UpdateUser;
            if (String.IsNullOrEmpty(row.UpdateUser))
            {
                command.Parameters["@UpdateUser"].IsNullable = true;
                command.Parameters["@UpdateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateDate", SqlDbType.VarChar, 8, "UpdateDate"));
            command.Parameters["@UpdateDate"].Value = row.UpdateDate;
            if (String.IsNullOrEmpty(row.UpdateDate))
            {
                command.Parameters["@UpdateDate"].IsNullable = true;
                command.Parameters["@UpdateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateTime", SqlDbType.VarChar, 6, "UpdateTime"));
            command.Parameters["@UpdateTime"].Value = row.UpdateTime;
            if (String.IsNullOrEmpty(row.UpdateTime))
            {
                command.Parameters["@UpdateTime"].IsNullable = true;
                command.Parameters["@UpdateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ShipDate", SqlDbType.VarChar, 8, "ShipDate"));
            command.Parameters["@ShipDate"].Value = row.ShipDate;
            if (String.IsNullOrEmpty(row.ShipDate))
            {
                command.Parameters["@ShipDate"].IsNullable = true;
                command.Parameters["@ShipDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.VarChar, 8, "WorkDate"));
            command.Parameters["@WorkDate"].Value = row.WorkDate;
            if (String.IsNullOrEmpty(row.WorkDate))
            {
                command.Parameters["@WorkDate"].IsNullable = true;
                command.Parameters["@WorkDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@PotatoKind", SqlDbType.VarChar, 1, "PotatoKind"));
            command.Parameters["@PotatoKind"].Value = row.PotatoKind;
            if (String.IsNullOrEmpty(row.PotatoKind))
            {
                command.Parameters["@PotatoKind"].IsNullable = true;
                command.Parameters["@PotatoKind"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@BoxQty", row.BoxQty));
            command.Parameters.Add(new SqlParameter("@PotatoWg", row.PotatoWg));
            command.Parameters.Add(new SqlParameter("@TransMark", SqlDbType.VarChar, 255, "TransMark"));
            command.Parameters["@TransMark"].Value = row.TransMark;
            if (String.IsNullOrEmpty(row.TransMark))
            {
                command.Parameters["@TransMark"].IsNullable = true;
                command.Parameters["@TransMark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelRemark", SqlDbType.VarChar, 255, "DelRemark"));
            command.Parameters["@DelRemark"].Value = row.DelRemark;
            if (String.IsNullOrEmpty(row.DelRemark))
            {
                command.Parameters["@DelRemark"].IsNullable = true;
                command.Parameters["@DelRemark"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Flag1", SqlDbType.VarChar, 50, "Flag1"));
            command.Parameters["@Flag1"].Value = row.Flag1;
            if (String.IsNullOrEmpty(row.Flag1))
            {
                command.Parameters["@Flag1"].IsNullable = true;
                command.Parameters["@Flag1"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Flag2", SqlDbType.VarChar, 50, "Flag2"));
            command.Parameters["@Flag2"].Value = row.Flag2;
            if (String.IsNullOrEmpty(row.Flag2))
            {
                command.Parameters["@Flag2"].IsNullable = true;
                command.Parameters["@Flag2"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@DelCom", SqlDbType.VarChar, 100, "DelCom"));
            command.Parameters["@DelCom"].Value = row.DelCom;
            if (String.IsNullOrEmpty(row.DelCom))
            {
                command.Parameters["@DelCom"].IsNullable = true;
                command.Parameters["@DelCom"].Value = "";
            }
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        public static void UpdateGB_POTATO_Client(GB_POTATO row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE GB_POTATO SET OrdName=@OrdName,OrdTel=@OrdTel,OrdCom=@OrdCom,OrdEMail=@OrdEMail,DelMan=@DelMan,DelTel=@DelTel,DelAddr=@DelAddr,ProdID=@ProdID,ProdName=@ProdName,Qty=@Qty,Price=@Price,Amount=@Amount,Remark=@Remark,UserCode=@UserCode,UpdateUser=@UpdateUser,UpdateDate=@UpdateDate,UpdateTime=@UpdateTime WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@OrdName", SqlDbType.VarChar, 50, "OrdName"));
            command.Parameters["@OrdName"].Value = row.OrdName;
            if (String.IsNullOrEmpty(row.OrdName))
            {
                command.Parameters["@OrdName"].IsNullable = true;
                command.Parameters["@OrdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdTel", SqlDbType.VarChar, 50, "OrdTel"));
            command.Parameters["@OrdTel"].Value = row.OrdTel;
            if (String.IsNullOrEmpty(row.OrdTel))
            {
                command.Parameters["@OrdTel"].IsNullable = true;
                command.Parameters["@OrdTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdCom", SqlDbType.VarChar, 100, "OrdCom"));
            command.Parameters["@OrdCom"].Value = row.OrdCom;
            if (String.IsNullOrEmpty(row.OrdCom))
            {
                command.Parameters["@OrdCom"].IsNullable = true;
                command.Parameters["@OrdCom"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@OrdEMail", SqlDbType.VarChar, 50, "OrdEMail"));
            command.Parameters["@OrdEMail"].Value = row.OrdEMail;
            if (String.IsNullOrEmpty(row.OrdEMail))
            {
                command.Parameters["@OrdEMail"].IsNullable = true;
                command.Parameters["@OrdEMail"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelMan", SqlDbType.VarChar, 50, "DelMan"));
            command.Parameters["@DelMan"].Value = row.DelMan;
            if (String.IsNullOrEmpty(row.DelMan))
            {
                command.Parameters["@DelMan"].IsNullable = true;
                command.Parameters["@DelMan"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelTel", SqlDbType.VarChar, 50, "DelTel"));
            command.Parameters["@DelTel"].Value = row.DelTel;
            if (String.IsNullOrEmpty(row.DelTel))
            {
                command.Parameters["@DelTel"].IsNullable = true;
                command.Parameters["@DelTel"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DelAddr", SqlDbType.VarChar, 50, "DelAddr"));
            command.Parameters["@DelAddr"].Value = row.DelAddr;
            if (String.IsNullOrEmpty(row.DelAddr))
            {
                command.Parameters["@DelAddr"].IsNullable = true;
                command.Parameters["@DelAddr"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdID", SqlDbType.VarChar, 50, "ProdID"));
            command.Parameters["@ProdID"].Value = row.ProdID;
            if (String.IsNullOrEmpty(row.ProdID))
            {
                command.Parameters["@ProdID"].IsNullable = true;
                command.Parameters["@ProdID"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ProdName", SqlDbType.VarChar, 200, "ProdName"));
            command.Parameters["@ProdName"].Value = row.ProdName;
            if (String.IsNullOrEmpty(row.ProdName))
            {
                command.Parameters["@ProdName"].IsNullable = true;
                command.Parameters["@ProdName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Qty", row.Qty));
            command.Parameters.Add(new SqlParameter("@Price", row.Price));
            command.Parameters.Add(new SqlParameter("@Amount", row.Amount));
            command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.VarChar, 250, "Remark"));
            command.Parameters["@Remark"].Value = row.Remark;
            if (String.IsNullOrEmpty(row.Remark))
            {
                command.Parameters["@Remark"].IsNullable = true;
                command.Parameters["@Remark"].Value = "";
            }
            //command.Parameters.Add(new SqlParameter("@DocType", SqlDbType.VarChar, 20, "DocType"));
            //command.Parameters["@DocType"].Value = row.DocType;
            //if (String.IsNullOrEmpty(row.DocType))
            //{
            //    command.Parameters["@DocType"].IsNullable = true;
            //    command.Parameters["@DocType"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@OrdNo", SqlDbType.VarChar, 20, "OrdNo"));
            //command.Parameters["@OrdNo"].Value = row.OrdNo;
            //if (String.IsNullOrEmpty(row.OrdNo))
            //{
            //    command.Parameters["@OrdNo"].IsNullable = true;
            //    command.Parameters["@OrdNo"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@InvNo", SqlDbType.VarChar, 20, "InvNo"));
            //command.Parameters["@InvNo"].Value = row.InvNo;
            //if (String.IsNullOrEmpty(row.InvNo))
            //{
            //    command.Parameters["@InvNo"].IsNullable = true;
            //    command.Parameters["@InvNo"].Value = "";
            //}
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
            command.Parameters["@UpdateUser"].Value = row.UpdateUser;
            if (String.IsNullOrEmpty(row.UpdateUser))
            {
                command.Parameters["@UpdateUser"].IsNullable = true;
                command.Parameters["@UpdateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateDate", SqlDbType.VarChar, 8, "UpdateDate"));
            command.Parameters["@UpdateDate"].Value = row.UpdateDate;
            if (String.IsNullOrEmpty(row.UpdateDate))
            {
                command.Parameters["@UpdateDate"].IsNullable = true;
                command.Parameters["@UpdateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateTime", SqlDbType.VarChar, 6, "UpdateTime"));
            command.Parameters["@UpdateTime"].Value = row.UpdateTime;
            if (String.IsNullOrEmpty(row.UpdateTime))
            {
                command.Parameters["@UpdateTime"].IsNullable = true;
                command.Parameters["@UpdateTime"].Value = "";
            }
            //command.Parameters.Add(new SqlParameter("@ShipDate", SqlDbType.VarChar, 8, "ShipDate"));
            //command.Parameters["@ShipDate"].Value = row.ShipDate;
            //if (String.IsNullOrEmpty(row.ShipDate))
            //{
            //    command.Parameters["@ShipDate"].IsNullable = true;
            //    command.Parameters["@ShipDate"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.VarChar, 8, "WorkDate"));
            //command.Parameters["@WorkDate"].Value = row.WorkDate;
            //if (String.IsNullOrEmpty(row.WorkDate))
            //{
            //    command.Parameters["@WorkDate"].IsNullable = true;
            //    command.Parameters["@WorkDate"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@PotatoKind", SqlDbType.VarChar, 1, "PotatoKind"));
            //command.Parameters["@PotatoKind"].Value = row.PotatoKind;
            //if (String.IsNullOrEmpty(row.PotatoKind))
            //{
            //    command.Parameters["@PotatoKind"].IsNullable = true;
            //    command.Parameters["@PotatoKind"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@BoxQty", row.BoxQty));
            //command.Parameters.Add(new SqlParameter("@PotatoWg", row.PotatoWg));
            //command.Parameters.Add(new SqlParameter("@TransMark", SqlDbType.VarChar, 255, "TransMark"));
            //command.Parameters["@TransMark"].Value = row.TransMark;
            //if (String.IsNullOrEmpty(row.TransMark))
            //{
            //    command.Parameters["@TransMark"].IsNullable = true;
            //    command.Parameters["@TransMark"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@DelRemark", SqlDbType.VarChar, 255, "DelRemark"));
            //command.Parameters["@DelRemark"].Value = row.DelRemark;
            //if (String.IsNullOrEmpty(row.DelRemark))
            //{
            //    command.Parameters["@DelRemark"].IsNullable = true;
            //    command.Parameters["@DelRemark"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@Flag1", SqlDbType.VarChar, 50, "Flag1"));
            //command.Parameters["@Flag1"].Value = row.Flag1;
            //if (String.IsNullOrEmpty(row.Flag1))
            //{
            //    command.Parameters["@Flag1"].IsNullable = true;
            //    command.Parameters["@Flag1"].Value = "";
            //}
            //command.Parameters.Add(new SqlParameter("@Flag2", SqlDbType.VarChar, 50, "Flag2"));
            //command.Parameters["@Flag2"].Value = row.Flag2;
            //if (String.IsNullOrEmpty(row.Flag2))
            //{
            //    command.Parameters["@Flag2"].IsNullable = true;
            //    command.Parameters["@Flag2"].Value = "";
            //}

            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        // GB_POTATO Delete
        public static void DeleteGB_POTATO(GB_POTATO row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE GB_POTATO WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        // GB_POTATO Select
        public static DataTable GetGB_POTATO(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT * FROM GB_POTATO WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "GB_POTATO");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["GB_POTATO"];
        }
        // Condition 版本
        public static DataTable GetGB_POTATO_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT * FROM GB_POTATO WHERE 1=1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "GB_POTATO");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["GB_POTATO"];
        }
    }

}