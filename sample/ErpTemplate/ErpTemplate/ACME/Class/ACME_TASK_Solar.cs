 using System;
 using System.Data;
 using System.Configuration;
 using System.Web;
 using System.Data.SqlClient;
 
 /// <summary>
 /// Summary description for ACME_TASK_Solar
 /// 作者:
 /// </summary>
// ACME_TASK_Solar 資料結構
public class ACME_TASK_Solar
{
    public static string AcmesqlSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
    private string _PrjCode;
    private string _PrjName;
    private string _CardCode;
    private string _CardName;
    private string _SlpCode;
    private string _SlpName;
    private string _SignDate;
    private string _SignDate_Memo;
    private string _StartDate_Sign;
    private string _EndDate_Sign;
    private string _AcDate_Sign;
    private string _PayDate1;
    private string _PayDate1_Memo;
    private string _StartDate_PayDate1;
    private string _EndDate_PayDate1;
    private string _AcDate_PayDate1;
    private string _Ep1;
    private string _Ep2;
    private string _Ep3;
    private string _EpDate1;
    private string _EpDate2;
    private string _Ep_Memo;
    private string _StartDate_Ep;
    private string _EndDate_Ep;
    private string _AcDate_Ep;
    private string _PayDate2;
    private string _PayDate2_Memo;
    private string _StartDate_PayDate2;
    private string _EndDate_PayDate2;
    private string _AcDate_PayDate2;
    private string _Fi1;
    private string _Fi2;
    private string _Fi3;
    private string _FiDate1;
    private string _FiDate2;
    private string _Fi_Memo;
    private string _StartDate_Fi;
    private string _EndDate_Fi;
    private string _AcDate_Fi;
    private string _Pr1;
    private string _Pr1Qty;
    private string _Pr2;
    private string _Pr2Qty;
    private string _Pr3;
    private string _Pr3Qty;
    private string _Pr4;
    private string _Pr4Qty;
    private string _Pr5;
    private string _Pr5Qty;
    private string _Pr_Memo;
    private string _StartDate_Pr;
    private string _EndDate_Pr;
    private string _AcDate_Pr;
    private string _WkDate1;
    private string _WkDate2;
    private string _WkDate3;
    private string _WkDate4;
    private string _WkDate5;
    private string _Wk_Memo;
    private string _StartDate_wk;
    private string _EndDate_wk;
    private string _AcDate_wk;
    private string _St1;
    private string _St2;
    private string _St3;
    private string _St4;
    private string _St5;
    private string _St6;
    private string _St_Memo;
    private string _StartDate_St;
    private string _EndDate_St;
    private string _AcDate_St;
    private string _ClDate1;
    private string _ClDate2;
    private string _ClDate3;
    private string _Cl_Memo;
    private string _StartDate_Cl;
    private string _EndDate_Cl;
    private string _AcDate_Cl;
    private string _PayDate3;
    private string _PayDate3_Memo;
    private string _StartDate_PayDate3;
    private string _EndDate_PayDate3;
    private string _AcDate_PayDate3;
    private string _Remark;
    private string _CreateDate;
    private string _CreateTime;
    private string _CreateUser;
    private string _UpdateDate;
    private string _UpdateTime;
    private string _UpdateUser;
    private string _PrjPercent;
    private string _CurrentStage;


    public string PrjCode { get { return _PrjCode; } set { _PrjCode = value; } }
    public string PrjName { get { return _PrjName; } set { _PrjName = value; } }
    public string CardCode { get { return _CardCode; } set { _CardCode = value; } }
    public string CardName { get { return _CardName; } set { _CardName = value; } }
    public string SlpCode { get { return _SlpCode; } set { _SlpCode = value; } }
    public string SlpName { get { return _SlpName; } set { _SlpName = value; } }
    public string SignDate { get { return _SignDate; } set { _SignDate = value; } }
    public string SignDate_Memo { get { return _SignDate_Memo; } set { _SignDate_Memo = value; } }
    public string StartDate_Sign { get { return _StartDate_Sign; } set { _StartDate_Sign = value; } }
    public string EndDate_Sign { get { return _EndDate_Sign; } set { _EndDate_Sign = value; } }
    public string AcDate_Sign { get { return _AcDate_Sign; } set { _AcDate_Sign = value; } }
    public string PayDate1 { get { return _PayDate1; } set { _PayDate1 = value; } }
    public string PayDate1_Memo { get { return _PayDate1_Memo; } set { _PayDate1_Memo = value; } }
    public string StartDate_PayDate1 { get { return _StartDate_PayDate1; } set { _StartDate_PayDate1 = value; } }
    public string EndDate_PayDate1 { get { return _EndDate_PayDate1; } set { _EndDate_PayDate1 = value; } }
    public string AcDate_PayDate1 { get { return _AcDate_PayDate1; } set { _AcDate_PayDate1 = value; } }
    public string Ep1 { get { return _Ep1; } set { _Ep1 = value; } }
    public string Ep2 { get { return _Ep2; } set { _Ep2 = value; } }
    public string Ep3 { get { return _Ep3; } set { _Ep3 = value; } }
    public string EpDate1 { get { return _EpDate1; } set { _EpDate1 = value; } }
    public string EpDate2 { get { return _EpDate2; } set { _EpDate2 = value; } }
    public string Ep_Memo { get { return _Ep_Memo; } set { _Ep_Memo = value; } }
    public string StartDate_Ep { get { return _StartDate_Ep; } set { _StartDate_Ep = value; } }
    public string EndDate_Ep { get { return _EndDate_Ep; } set { _EndDate_Ep = value; } }
    public string AcDate_Ep { get { return _AcDate_Ep; } set { _AcDate_Ep = value; } }
    public string PayDate2 { get { return _PayDate2; } set { _PayDate2 = value; } }
    public string PayDate2_Memo { get { return _PayDate2_Memo; } set { _PayDate2_Memo = value; } }
    public string StartDate_PayDate2 { get { return _StartDate_PayDate2; } set { _StartDate_PayDate2 = value; } }
    public string EndDate_PayDate2 { get { return _EndDate_PayDate2; } set { _EndDate_PayDate2 = value; } }
    public string AcDate_PayDate2 { get { return _AcDate_PayDate2; } set { _AcDate_PayDate2 = value; } }
    public string Fi1 { get { return _Fi1; } set { _Fi1 = value; } }
    public string Fi2 { get { return _Fi2; } set { _Fi2 = value; } }
    public string Fi3 { get { return _Fi3; } set { _Fi3 = value; } }
    public string FiDate1 { get { return _FiDate1; } set { _FiDate1 = value; } }
    public string FiDate2 { get { return _FiDate2; } set { _FiDate2 = value; } }
    public string Fi_Memo { get { return _Fi_Memo; } set { _Fi_Memo = value; } }
    public string StartDate_Fi { get { return _StartDate_Fi; } set { _StartDate_Fi = value; } }
    public string EndDate_Fi { get { return _EndDate_Fi; } set { _EndDate_Fi = value; } }
    public string AcDate_Fi { get { return _AcDate_Fi; } set { _AcDate_Fi = value; } }
    public string Pr1 { get { return _Pr1; } set { _Pr1 = value; } }
    public string Pr1Qty { get { return _Pr1Qty; } set { _Pr1Qty = value; } }
    public string Pr2 { get { return _Pr2; } set { _Pr2 = value; } }
    public string Pr2Qty { get { return _Pr2Qty; } set { _Pr2Qty = value; } }
    public string Pr3 { get { return _Pr3; } set { _Pr3 = value; } }
    public string Pr3Qty { get { return _Pr3Qty; } set { _Pr3Qty = value; } }
    public string Pr4 { get { return _Pr4; } set { _Pr4 = value; } }
    public string Pr4Qty { get { return _Pr4Qty; } set { _Pr4Qty = value; } }
    public string Pr5 { get { return _Pr5; } set { _Pr5 = value; } }
    public string Pr5Qty { get { return _Pr5Qty; } set { _Pr5Qty = value; } }
    public string Pr_Memo { get { return _Pr_Memo; } set { _Pr_Memo = value; } }
    public string StartDate_Pr { get { return _StartDate_Pr; } set { _StartDate_Pr = value; } }
    public string EndDate_Pr { get { return _EndDate_Pr; } set { _EndDate_Pr = value; } }
    public string AcDate_Pr { get { return _AcDate_Pr; } set { _AcDate_Pr = value; } }
    public string WkDate1 { get { return _WkDate1; } set { _WkDate1 = value; } }
    public string WkDate2 { get { return _WkDate2; } set { _WkDate2 = value; } }
    public string WkDate3 { get { return _WkDate3; } set { _WkDate3 = value; } }
    public string WkDate4 { get { return _WkDate4; } set { _WkDate4 = value; } }
    public string WkDate5 { get { return _WkDate5; } set { _WkDate5 = value; } }
    public string Wk_Memo { get { return _Wk_Memo; } set { _Wk_Memo = value; } }
    public string StartDate_wk { get { return _StartDate_wk; } set { _StartDate_wk = value; } }
    public string EndDate_wk { get { return _EndDate_wk; } set { _EndDate_wk = value; } }
    public string AcDate_wk { get { return _AcDate_wk; } set { _AcDate_wk = value; } }
    public string St1 { get { return _St1; } set { _St1 = value; } }
    public string St2 { get { return _St2; } set { _St2 = value; } }
    public string St3 { get { return _St3; } set { _St3 = value; } }
    public string St4 { get { return _St4; } set { _St4 = value; } }
    public string St5 { get { return _St5; } set { _St5 = value; } }
    public string St6 { get { return _St6; } set { _St6 = value; } }
    public string St_Memo { get { return _St_Memo; } set { _St_Memo = value; } }
    public string StartDate_St { get { return _StartDate_St; } set { _StartDate_St = value; } }
    public string EndDate_St { get { return _EndDate_St; } set { _EndDate_St = value; } }
    public string AcDate_St { get { return _AcDate_St; } set { _AcDate_St = value; } }
    public string ClDate1 { get { return _ClDate1; } set { _ClDate1 = value; } }
    public string ClDate2 { get { return _ClDate2; } set { _ClDate2 = value; } }
    public string ClDate3 { get { return _ClDate3; } set { _ClDate3 = value; } }
    public string Cl_Memo { get { return _Cl_Memo; } set { _Cl_Memo = value; } }
    public string StartDate_Cl { get { return _StartDate_Cl; } set { _StartDate_Cl = value; } }
    public string EndDate_Cl { get { return _EndDate_Cl; } set { _EndDate_Cl = value; } }
    public string AcDate_Cl { get { return _AcDate_Cl; } set { _AcDate_Cl = value; } }
    public string PayDate3 { get { return _PayDate3; } set { _PayDate3 = value; } }
    public string PayDate3_Memo { get { return _PayDate3_Memo; } set { _PayDate3_Memo = value; } }
    public string StartDate_PayDate3 { get { return _StartDate_PayDate3; } set { _StartDate_PayDate3 = value; } }
    public string EndDate_PayDate3 { get { return _EndDate_PayDate3; } set { _EndDate_PayDate3 = value; } }
    public string AcDate_PayDate3 { get { return _AcDate_PayDate3; } set { _AcDate_PayDate3 = value; } }
    public string Remark { get { return _Remark; } set { _Remark = value; } }
    public string PrjPercent { get { return _PrjPercent; } set { _PrjPercent = value; } }
    public string CurrentStage { get { return _CurrentStage; } set { _CurrentStage = value; } }

    
    public string CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
    public string CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
    public string CreateUser { get { return _CreateUser; } set { _CreateUser = value; } }
    public string UpdateDate { get { return _UpdateDate; } set { _UpdateDate = value; } }
    public string UpdateTime { get { return _UpdateTime; } set { _UpdateTime = value; } }
    public string UpdateUser { get { return _UpdateUser; } set { _UpdateUser = value; } }

    public ACME_TASK_Solar(string PrjCode, string PrjName, string CardCode, string CardName, string SlpCode, string SlpName, string
  SignDate, string SignDate_Memo, string StartDate_Sign, string EndDate_Sign, string AcDate_Sign, string PayDate1, string
  PayDate1_Memo, string StartDate_PayDate1, string EndDate_PayDate1, string AcDate_PayDate1, string Ep1, string Ep2, string Ep3, string
  EpDate1, string EpDate2, string Ep_Memo, string StartDate_Ep, string EndDate_Ep, string AcDate_Ep, string PayDate2, string
  PayDate2_Memo, string StartDate_PayDate2, string EndDate_PayDate2, string AcDate_PayDate2, string Fi1, string Fi2, string Fi3, string
  FiDate1, string FiDate2, string Fi_Memo, string StartDate_Fi, string EndDate_Fi, string AcDate_Fi, string Pr1, string Pr1Qty, string
  Pr2, string Pr2Qty, string Pr3, string Pr3Qty, string Pr4, string Pr4Qty, string Pr5, string Pr5Qty, string Pr_Memo, string
  StartDate_Pr, string EndDate_Pr, string AcDate_Pr, string WkDate1, string WkDate2, string WkDate3, string WkDate4, string WkDate5, string
  Wk_Memo, string StartDate_wk, string EndDate_wk, string AcDate_wk, string St1, string St2, string St3, string St4, string St5, string
  St6, string St_Memo, string StartDate_St, string EndDate_St, string AcDate_St, string ClDate1, string ClDate2, string ClDate3, string
  Cl_Memo, string StartDate_Cl, string EndDate_Cl, string AcDate_Cl, string PayDate3, string PayDate3_Memo, string
  StartDate_PayDate3, string EndDate_PayDate3, string AcDate_PayDate3, string Remark, string CreateDate, string CreateTime, string
  CreateUser, string UpdateDate, string UpdateTime, string UpdateUser)
    {
        _PrjCode = PrjCode;
        _PrjName = PrjName;
        _CardCode = CardCode;
        _CardName = CardName;
        _SlpCode = SlpCode;
        _SlpName = SlpName;
        _SignDate = SignDate;
        _SignDate_Memo = SignDate_Memo;
        _StartDate_Sign = StartDate_Sign;
        _EndDate_Sign = EndDate_Sign;
        _AcDate_Sign = AcDate_Sign;
        _PayDate1 = PayDate1;
        _PayDate1_Memo = PayDate1_Memo;
        _StartDate_PayDate1 = StartDate_PayDate1;
        _EndDate_PayDate1 = EndDate_PayDate1;
        _AcDate_PayDate1 = AcDate_PayDate1;
        _Ep1 = Ep1;
        _Ep2 = Ep2;
        _Ep3 = Ep3;
        _EpDate1 = EpDate1;
        _EpDate2 = EpDate2;
        _Ep_Memo = Ep_Memo;
        _StartDate_Ep = StartDate_Ep;
        _EndDate_Ep = EndDate_Ep;
        _AcDate_Ep = AcDate_Ep;
        _PayDate2 = PayDate2;
        _PayDate2_Memo = PayDate2_Memo;
        _StartDate_PayDate2 = StartDate_PayDate2;
        _EndDate_PayDate2 = EndDate_PayDate2;
        _AcDate_PayDate2 = AcDate_PayDate2;
        _Fi1 = Fi1;
        _Fi2 = Fi2;
        _Fi3 = Fi3;
        _FiDate1 = FiDate1;
        _FiDate2 = FiDate2;
        _Fi_Memo = Fi_Memo;
        _StartDate_Fi = StartDate_Fi;
        _EndDate_Fi = EndDate_Fi;
        _AcDate_Fi = AcDate_Fi;
        _Pr1 = Pr1;
        _Pr1Qty = Pr1Qty;
        _Pr2 = Pr2;
        _Pr2Qty = Pr2Qty;
        _Pr3 = Pr3;
        _Pr3Qty = Pr3Qty;
        _Pr4 = Pr4;
        _Pr4Qty = Pr4Qty;
        _Pr5 = Pr5;
        _Pr5Qty = Pr5Qty;
        _Pr_Memo = Pr_Memo;
        _StartDate_Pr = StartDate_Pr;
        _EndDate_Pr = EndDate_Pr;
        _AcDate_Pr = AcDate_Pr;
        _WkDate1 = WkDate1;
        _WkDate2 = WkDate2;
        _WkDate3 = WkDate3;
        _WkDate4 = WkDate4;
        _WkDate5 = WkDate5;
        _Wk_Memo = Wk_Memo;
        _StartDate_wk = StartDate_wk;
        _EndDate_wk = EndDate_wk;
        _AcDate_wk = AcDate_wk;
        _St1 = St1;
        _St2 = St2;
        _St3 = St3;
        _St4 = St4;
        _St5 = St5;
        _St6 = St6;
        _St_Memo = St_Memo;
        _StartDate_St = StartDate_St;
        _EndDate_St = EndDate_St;
        _AcDate_St = AcDate_St;
        _ClDate1 = ClDate1;
        _ClDate2 = ClDate2;
        _ClDate3 = ClDate3;
        _Cl_Memo = Cl_Memo;
        _StartDate_Cl = StartDate_Cl;
        _EndDate_Cl = EndDate_Cl;
        _AcDate_Cl = AcDate_Cl;
        _PayDate3 = PayDate3;
        _PayDate3_Memo = PayDate3_Memo;
        _StartDate_PayDate3 = StartDate_PayDate3;
        _EndDate_PayDate3 = EndDate_PayDate3;
        _AcDate_PayDate3 = AcDate_PayDate3;
        _Remark = Remark;
        _CreateDate = CreateDate;
        _CreateTime = CreateTime;
        _CreateUser = CreateUser;
        _UpdateDate = UpdateDate;
        _UpdateTime = UpdateTime;
        _UpdateUser = UpdateUser;
       
    }
    public ACME_TASK_Solar()
    {
    }
    // ACME_TASK_Solar Insert
    public static void AddACME_TASK_Solar(ACME_TASK_Solar row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        SqlCommand command = new SqlCommand("Insert into ACME_TASK_Solar(PrjCode,PrjName,CardCode,CardName,SlpCode,SlpName,SignDate,SignDate_Memo,StartDate_Sign,EndDate_Sign,AcDate_Sign,PayDate1,PayDate1_Memo,StartDate_PayDate1,EndDate_PayDate1,AcDate_PayDate1,Ep1,Ep2,Ep3,EpDate1,EpDate2,Ep_Memo,StartDate_Ep,EndDate_Ep,AcDate_Ep,PayDate2,PayDate2_Memo,StartDate_PayDate2,EndDate_PayDate2,AcDate_PayDate2,Fi1,Fi2,Fi3,FiDate1,FiDate2,Fi_Memo,StartDate_Fi,EndDate_Fi,AcDate_Fi,Pr1,Pr1Qty,Pr2,Pr2Qty,Pr3,Pr3Qty,Pr4,Pr4Qty,Pr5,Pr5Qty,Pr_Memo,StartDate_Pr,EndDate_Pr,AcDate_Pr,WkDate1,WkDate2,WkDate3,WkDate4,WkDate5,Wk_Memo,StartDate_wk,EndDate_wk,AcDate_wk,St1,St2,St3,St4,St5,St6,St_Memo,StartDate_St,EndDate_St,AcDate_St,ClDate1,ClDate2,ClDate3,Cl_Memo,StartDate_Cl,EndDate_Cl,AcDate_Cl,PayDate3,PayDate3_Memo,StartDate_PayDate3,EndDate_PayDate3,AcDate_PayDate3,Remark,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser) values(@PrjCode,@PrjName,@CardCode,@CardName,@SlpCode,@SlpName,@SignDate,@SignDate_Memo,@StartDate_Sign,@EndDate_Sign,@AcDate_Sign,@PayDate1,@PayDate1_Memo,@StartDate_PayDate1,@EndDate_PayDate1,@AcDate_PayDate1,@Ep1,@Ep2,@Ep3,@EpDate1,@EpDate2,@Ep_Memo,@StartDate_Ep,@EndDate_Ep,@AcDate_Ep,@PayDate2,@PayDate2_Memo,@StartDate_PayDate2,@EndDate_PayDate2,@AcDate_PayDate2,@Fi1,@Fi2,@Fi3,@FiDate1,@FiDate2,@Fi_Memo,@StartDate_Fi,@EndDate_Fi,@AcDate_Fi,@Pr1,@Pr1Qty,@Pr2,@Pr2Qty,@Pr3,@Pr3Qty,@Pr4,@Pr4Qty,@Pr5,@Pr5Qty,@Pr_Memo,@StartDate_Pr,@EndDate_Pr,@AcDate_Pr,@WkDate1,@WkDate2,@WkDate3,@WkDate4,@WkDate5,@Wk_Memo,@StartDate_wk,@EndDate_wk,@AcDate_wk,@St1,@St2,@St3,@St4,@St5,@St6,@St_Memo,@StartDate_St,@EndDate_St,@AcDate_St,@ClDate1,@ClDate2,@ClDate3,@Cl_Memo,@StartDate_Cl,@EndDate_Cl,@AcDate_Cl,@PayDate3,@PayDate3_Memo,@StartDate_PayDate3,@EndDate_PayDate3,@AcDate_PayDate3,@Remark,@CreateDate,@CreateTime,@CreateUser,@UpdateDate,@UpdateTime,@UpdateUser)", connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
        command.Parameters["@PrjCode"].Value = row.PrjCode;
        if (String.IsNullOrEmpty(row.PrjCode))
        {
            command.Parameters["@PrjCode"].IsNullable = true;
            command.Parameters["@PrjCode"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PrjName", SqlDbType.VarChar, 50, "PrjName"));
        command.Parameters["@PrjName"].Value = row.PrjName;
        if (String.IsNullOrEmpty(row.PrjName))
        {
            command.Parameters["@PrjName"].IsNullable = true;
            command.Parameters["@PrjName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CardCode", SqlDbType.VarChar, 50, "CardCode"));
        command.Parameters["@CardCode"].Value = row.CardCode;
        if (String.IsNullOrEmpty(row.CardCode))
        {
            command.Parameters["@CardCode"].IsNullable = true;
            command.Parameters["@CardCode"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CardName", SqlDbType.VarChar, 50, "CardName"));
        command.Parameters["@CardName"].Value = row.CardName;
        if (String.IsNullOrEmpty(row.CardName))
        {
            command.Parameters["@CardName"].IsNullable = true;
            command.Parameters["@CardName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SlpCode", SqlDbType.VarChar, 50, "SlpCode"));
        command.Parameters["@SlpCode"].Value = row.SlpCode;
        if (String.IsNullOrEmpty(row.SlpCode))
        {
            command.Parameters["@SlpCode"].IsNullable = true;
            command.Parameters["@SlpCode"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SlpName", SqlDbType.VarChar, 50, "SlpName"));
        command.Parameters["@SlpName"].Value = row.SlpName;
        if (String.IsNullOrEmpty(row.SlpName))
        {
            command.Parameters["@SlpName"].IsNullable = true;
            command.Parameters["@SlpName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SignDate", SqlDbType.VarChar, 8, "SignDate"));
        command.Parameters["@SignDate"].Value = row.SignDate;
        if (String.IsNullOrEmpty(row.SignDate))
        {
            command.Parameters["@SignDate"].IsNullable = true;
            command.Parameters["@SignDate"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SignDate_Memo", SqlDbType.VarChar, 250, "SignDate_Memo"));
        command.Parameters["@SignDate_Memo"].Value = row.SignDate_Memo;
        if (String.IsNullOrEmpty(row.SignDate_Memo))
        {
            command.Parameters["@SignDate_Memo"].IsNullable = true;
            command.Parameters["@SignDate_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Sign", SqlDbType.VarChar, 8, "StartDate_Sign"));
        command.Parameters["@StartDate_Sign"].Value = row.StartDate_Sign;
        if (String.IsNullOrEmpty(row.StartDate_Sign))
        {
            command.Parameters["@StartDate_Sign"].IsNullable = true;
            command.Parameters["@StartDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Sign", SqlDbType.VarChar, 8, "EndDate_Sign"));
        command.Parameters["@EndDate_Sign"].Value = row.EndDate_Sign;
        if (String.IsNullOrEmpty(row.EndDate_Sign))
        {
            command.Parameters["@EndDate_Sign"].IsNullable = true;
            command.Parameters["@EndDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Sign", SqlDbType.VarChar, 8, "AcDate_Sign"));
        command.Parameters["@AcDate_Sign"].Value = row.AcDate_Sign;
        if (String.IsNullOrEmpty(row.AcDate_Sign))
        {
            command.Parameters["@AcDate_Sign"].IsNullable = true;
            command.Parameters["@AcDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate1", SqlDbType.VarChar, 8, "PayDate1"));
        command.Parameters["@PayDate1"].Value = row.PayDate1;
        if (String.IsNullOrEmpty(row.PayDate1))
        {
            command.Parameters["@PayDate1"].IsNullable = true;
            command.Parameters["@PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate1_Memo", SqlDbType.VarChar, 250, "PayDate1_Memo"));
        command.Parameters["@PayDate1_Memo"].Value = row.PayDate1_Memo;
        if (String.IsNullOrEmpty(row.PayDate1_Memo))
        {
            command.Parameters["@PayDate1_Memo"].IsNullable = true;
            command.Parameters["@PayDate1_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate1", SqlDbType.VarChar, 8, "StartDate_PayDate1"));
        command.Parameters["@StartDate_PayDate1"].Value = row.StartDate_PayDate1;
        if (String.IsNullOrEmpty(row.StartDate_PayDate1))
        {
            command.Parameters["@StartDate_PayDate1"].IsNullable = true;
            command.Parameters["@StartDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate1", SqlDbType.VarChar, 8, "EndDate_PayDate1"));
        command.Parameters["@EndDate_PayDate1"].Value = row.EndDate_PayDate1;
        if (String.IsNullOrEmpty(row.EndDate_PayDate1))
        {
            command.Parameters["@EndDate_PayDate1"].IsNullable = true;
            command.Parameters["@EndDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate1", SqlDbType.VarChar, 8, "AcDate_PayDate1"));
        command.Parameters["@AcDate_PayDate1"].Value = row.AcDate_PayDate1;
        if (String.IsNullOrEmpty(row.AcDate_PayDate1))
        {
            command.Parameters["@AcDate_PayDate1"].IsNullable = true;
            command.Parameters["@AcDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep1", SqlDbType.VarChar, 50, "Ep1"));
        command.Parameters["@Ep1"].Value = row.Ep1;
        if (String.IsNullOrEmpty(row.Ep1))
        {
            command.Parameters["@Ep1"].IsNullable = true;
            command.Parameters["@Ep1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep2", SqlDbType.VarChar, 50, "Ep2"));
        command.Parameters["@Ep2"].Value = row.Ep2;
        if (String.IsNullOrEmpty(row.Ep2))
        {
            command.Parameters["@Ep2"].IsNullable = true;
            command.Parameters["@Ep2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep3", SqlDbType.VarChar, 50, "Ep3"));
        command.Parameters["@Ep3"].Value = row.Ep3;
        if (String.IsNullOrEmpty(row.Ep3))
        {
            command.Parameters["@Ep3"].IsNullable = true;
            command.Parameters["@Ep3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EpDate1", SqlDbType.VarChar, 8, "EpDate1"));
        command.Parameters["@EpDate1"].Value = row.EpDate1;
        if (String.IsNullOrEmpty(row.EpDate1))
        {
            command.Parameters["@EpDate1"].IsNullable = true;
            command.Parameters["@EpDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EpDate2", SqlDbType.VarChar, 8, "EpDate2"));
        command.Parameters["@EpDate2"].Value = row.EpDate2;
        if (String.IsNullOrEmpty(row.EpDate2))
        {
            command.Parameters["@EpDate2"].IsNullable = true;
            command.Parameters["@EpDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep_Memo", SqlDbType.VarChar, 250, "Ep_Memo"));
        command.Parameters["@Ep_Memo"].Value = row.Ep_Memo;
        if (String.IsNullOrEmpty(row.Ep_Memo))
        {
            command.Parameters["@Ep_Memo"].IsNullable = true;
            command.Parameters["@Ep_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Ep", SqlDbType.VarChar, 8, "StartDate_Ep"));
        command.Parameters["@StartDate_Ep"].Value = row.StartDate_Ep;
        if (String.IsNullOrEmpty(row.StartDate_Ep))
        {
            command.Parameters["@StartDate_Ep"].IsNullable = true;
            command.Parameters["@StartDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Ep", SqlDbType.VarChar, 8, "EndDate_Ep"));
        command.Parameters["@EndDate_Ep"].Value = row.EndDate_Ep;
        if (String.IsNullOrEmpty(row.EndDate_Ep))
        {
            command.Parameters["@EndDate_Ep"].IsNullable = true;
            command.Parameters["@EndDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Ep", SqlDbType.VarChar, 8, "AcDate_Ep"));
        command.Parameters["@AcDate_Ep"].Value = row.AcDate_Ep;
        if (String.IsNullOrEmpty(row.AcDate_Ep))
        {
            command.Parameters["@AcDate_Ep"].IsNullable = true;
            command.Parameters["@AcDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate2", SqlDbType.VarChar, 8, "PayDate2"));
        command.Parameters["@PayDate2"].Value = row.PayDate2;
        if (String.IsNullOrEmpty(row.PayDate2))
        {
            command.Parameters["@PayDate2"].IsNullable = true;
            command.Parameters["@PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate2_Memo", SqlDbType.VarChar, 250, "PayDate2_Memo"));
        command.Parameters["@PayDate2_Memo"].Value = row.PayDate2_Memo;
        if (String.IsNullOrEmpty(row.PayDate2_Memo))
        {
            command.Parameters["@PayDate2_Memo"].IsNullable = true;
            command.Parameters["@PayDate2_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate2", SqlDbType.VarChar, 8, "StartDate_PayDate2"));
        command.Parameters["@StartDate_PayDate2"].Value = row.StartDate_PayDate2;
        if (String.IsNullOrEmpty(row.StartDate_PayDate2))
        {
            command.Parameters["@StartDate_PayDate2"].IsNullable = true;
            command.Parameters["@StartDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate2", SqlDbType.VarChar, 8, "EndDate_PayDate2"));
        command.Parameters["@EndDate_PayDate2"].Value = row.EndDate_PayDate2;
        if (String.IsNullOrEmpty(row.EndDate_PayDate2))
        {
            command.Parameters["@EndDate_PayDate2"].IsNullable = true;
            command.Parameters["@EndDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate2", SqlDbType.VarChar, 8, "AcDate_PayDate2"));
        command.Parameters["@AcDate_PayDate2"].Value = row.AcDate_PayDate2;
        if (String.IsNullOrEmpty(row.AcDate_PayDate2))
        {
            command.Parameters["@AcDate_PayDate2"].IsNullable = true;
            command.Parameters["@AcDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi1", SqlDbType.VarChar, 50, "Fi1"));
        command.Parameters["@Fi1"].Value = row.Fi1;
        if (String.IsNullOrEmpty(row.Fi1))
        {
            command.Parameters["@Fi1"].IsNullable = true;
            command.Parameters["@Fi1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi2", SqlDbType.VarChar, 50, "Fi2"));
        command.Parameters["@Fi2"].Value = row.Fi2;
        if (String.IsNullOrEmpty(row.Fi2))
        {
            command.Parameters["@Fi2"].IsNullable = true;
            command.Parameters["@Fi2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi3", SqlDbType.VarChar, 50, "Fi3"));
        command.Parameters["@Fi3"].Value = row.Fi3;
        if (String.IsNullOrEmpty(row.Fi3))
        {
            command.Parameters["@Fi3"].IsNullable = true;
            command.Parameters["@Fi3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@FiDate1", SqlDbType.VarChar, 8, "FiDate1"));
        command.Parameters["@FiDate1"].Value = row.FiDate1;
        if (String.IsNullOrEmpty(row.FiDate1))
        {
            command.Parameters["@FiDate1"].IsNullable = true;
            command.Parameters["@FiDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@FiDate2", SqlDbType.VarChar, 8, "FiDate2"));
        command.Parameters["@FiDate2"].Value = row.FiDate2;
        if (String.IsNullOrEmpty(row.FiDate2))
        {
            command.Parameters["@FiDate2"].IsNullable = true;
            command.Parameters["@FiDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi_Memo", SqlDbType.VarChar, 250, "Fi_Memo"));
        command.Parameters["@Fi_Memo"].Value = row.Fi_Memo;
        if (String.IsNullOrEmpty(row.Fi_Memo))
        {
            command.Parameters["@Fi_Memo"].IsNullable = true;
            command.Parameters["@Fi_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Fi", SqlDbType.VarChar, 8, "StartDate_Fi"));
        command.Parameters["@StartDate_Fi"].Value = row.StartDate_Fi;
        if (String.IsNullOrEmpty(row.StartDate_Fi))
        {
            command.Parameters["@StartDate_Fi"].IsNullable = true;
            command.Parameters["@StartDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Fi", SqlDbType.VarChar, 8, "EndDate_Fi"));
        command.Parameters["@EndDate_Fi"].Value = row.EndDate_Fi;
        if (String.IsNullOrEmpty(row.EndDate_Fi))
        {
            command.Parameters["@EndDate_Fi"].IsNullable = true;
            command.Parameters["@EndDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Fi", SqlDbType.VarChar, 8, "AcDate_Fi"));
        command.Parameters["@AcDate_Fi"].Value = row.AcDate_Fi;
        if (String.IsNullOrEmpty(row.AcDate_Fi))
        {
            command.Parameters["@AcDate_Fi"].IsNullable = true;
            command.Parameters["@AcDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr1", SqlDbType.VarChar, 50, "Pr1"));
        command.Parameters["@Pr1"].Value = row.Pr1;
        if (String.IsNullOrEmpty(row.Pr1))
        {
            command.Parameters["@Pr1"].IsNullable = true;
            command.Parameters["@Pr1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr1Qty", SqlDbType.VarChar, 50, "Pr1Qty"));
        command.Parameters["@Pr1Qty"].Value = row.Pr1Qty;
        if (String.IsNullOrEmpty(row.Pr1Qty))
        {
            command.Parameters["@Pr1Qty"].IsNullable = true;
            command.Parameters["@Pr1Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr2", SqlDbType.VarChar, 50, "Pr2"));
        command.Parameters["@Pr2"].Value = row.Pr2;
        if (String.IsNullOrEmpty(row.Pr2))
        {
            command.Parameters["@Pr2"].IsNullable = true;
            command.Parameters["@Pr2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr2Qty", SqlDbType.VarChar, 50, "Pr2Qty"));
        command.Parameters["@Pr2Qty"].Value = row.Pr2Qty;
        if (String.IsNullOrEmpty(row.Pr2Qty))
        {
            command.Parameters["@Pr2Qty"].IsNullable = true;
            command.Parameters["@Pr2Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr3", SqlDbType.VarChar, 50, "Pr3"));
        command.Parameters["@Pr3"].Value = row.Pr3;
        if (String.IsNullOrEmpty(row.Pr3))
        {
            command.Parameters["@Pr3"].IsNullable = true;
            command.Parameters["@Pr3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr3Qty", SqlDbType.VarChar, 50, "Pr3Qty"));
        command.Parameters["@Pr3Qty"].Value = row.Pr3Qty;
        if (String.IsNullOrEmpty(row.Pr3Qty))
        {
            command.Parameters["@Pr3Qty"].IsNullable = true;
            command.Parameters["@Pr3Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr4", SqlDbType.VarChar, 50, "Pr4"));
        command.Parameters["@Pr4"].Value = row.Pr4;
        if (String.IsNullOrEmpty(row.Pr4))
        {
            command.Parameters["@Pr4"].IsNullable = true;
            command.Parameters["@Pr4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr4Qty", SqlDbType.VarChar, 50, "Pr4Qty"));
        command.Parameters["@Pr4Qty"].Value = row.Pr4Qty;
        if (String.IsNullOrEmpty(row.Pr4Qty))
        {
            command.Parameters["@Pr4Qty"].IsNullable = true;
            command.Parameters["@Pr4Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr5", SqlDbType.VarChar, 50, "Pr5"));
        command.Parameters["@Pr5"].Value = row.Pr5;
        if (String.IsNullOrEmpty(row.Pr5))
        {
            command.Parameters["@Pr5"].IsNullable = true;
            command.Parameters["@Pr5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr5Qty", SqlDbType.VarChar, 50, "Pr5Qty"));
        command.Parameters["@Pr5Qty"].Value = row.Pr5Qty;
        if (String.IsNullOrEmpty(row.Pr5Qty))
        {
            command.Parameters["@Pr5Qty"].IsNullable = true;
            command.Parameters["@Pr5Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr_Memo", SqlDbType.VarChar, 250, "Pr_Memo"));
        command.Parameters["@Pr_Memo"].Value = row.Pr_Memo;
        if (String.IsNullOrEmpty(row.Pr_Memo))
        {
            command.Parameters["@Pr_Memo"].IsNullable = true;
            command.Parameters["@Pr_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Pr", SqlDbType.VarChar, 8, "StartDate_Pr"));
        command.Parameters["@StartDate_Pr"].Value = row.StartDate_Pr;
        if (String.IsNullOrEmpty(row.StartDate_Pr))
        {
            command.Parameters["@StartDate_Pr"].IsNullable = true;
            command.Parameters["@StartDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Pr", SqlDbType.VarChar, 8, "EndDate_Pr"));
        command.Parameters["@EndDate_Pr"].Value = row.EndDate_Pr;
        if (String.IsNullOrEmpty(row.EndDate_Pr))
        {
            command.Parameters["@EndDate_Pr"].IsNullable = true;
            command.Parameters["@EndDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Pr", SqlDbType.VarChar, 8, "AcDate_Pr"));
        command.Parameters["@AcDate_Pr"].Value = row.AcDate_Pr;
        if (String.IsNullOrEmpty(row.AcDate_Pr))
        {
            command.Parameters["@AcDate_Pr"].IsNullable = true;
            command.Parameters["@AcDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate1", SqlDbType.VarChar, 8, "WkDate1"));
        command.Parameters["@WkDate1"].Value = row.WkDate1;
        if (String.IsNullOrEmpty(row.WkDate1))
        {
            command.Parameters["@WkDate1"].IsNullable = true;
            command.Parameters["@WkDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate2", SqlDbType.VarChar, 8, "WkDate2"));
        command.Parameters["@WkDate2"].Value = row.WkDate2;
        if (String.IsNullOrEmpty(row.WkDate2))
        {
            command.Parameters["@WkDate2"].IsNullable = true;
            command.Parameters["@WkDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate3", SqlDbType.VarChar, 8, "WkDate3"));
        command.Parameters["@WkDate3"].Value = row.WkDate3;
        if (String.IsNullOrEmpty(row.WkDate3))
        {
            command.Parameters["@WkDate3"].IsNullable = true;
            command.Parameters["@WkDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate4", SqlDbType.VarChar, 8, "WkDate4"));
        command.Parameters["@WkDate4"].Value = row.WkDate4;
        if (String.IsNullOrEmpty(row.WkDate4))
        {
            command.Parameters["@WkDate4"].IsNullable = true;
            command.Parameters["@WkDate4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate5", SqlDbType.VarChar, 8, "WkDate5"));
        command.Parameters["@WkDate5"].Value = row.WkDate5;
        if (String.IsNullOrEmpty(row.WkDate5))
        {
            command.Parameters["@WkDate5"].IsNullable = true;
            command.Parameters["@WkDate5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Wk_Memo", SqlDbType.VarChar, 250, "Wk_Memo"));
        command.Parameters["@Wk_Memo"].Value = row.Wk_Memo;
        if (String.IsNullOrEmpty(row.Wk_Memo))
        {
            command.Parameters["@Wk_Memo"].IsNullable = true;
            command.Parameters["@Wk_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_wk", SqlDbType.VarChar, 8, "StartDate_wk"));
        command.Parameters["@StartDate_wk"].Value = row.StartDate_wk;
        if (String.IsNullOrEmpty(row.StartDate_wk))
        {
            command.Parameters["@StartDate_wk"].IsNullable = true;
            command.Parameters["@StartDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_wk", SqlDbType.VarChar, 8, "EndDate_wk"));
        command.Parameters["@EndDate_wk"].Value = row.EndDate_wk;
        if (String.IsNullOrEmpty(row.EndDate_wk))
        {
            command.Parameters["@EndDate_wk"].IsNullable = true;
            command.Parameters["@EndDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_wk", SqlDbType.VarChar, 8, "AcDate_wk"));
        command.Parameters["@AcDate_wk"].Value = row.AcDate_wk;
        if (String.IsNullOrEmpty(row.AcDate_wk))
        {
            command.Parameters["@AcDate_wk"].IsNullable = true;
            command.Parameters["@AcDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St1", SqlDbType.VarChar, 50, "St1"));
        command.Parameters["@St1"].Value = row.St1;
        if (String.IsNullOrEmpty(row.St1))
        {
            command.Parameters["@St1"].IsNullable = true;
            command.Parameters["@St1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St2", SqlDbType.VarChar, 50, "St2"));
        command.Parameters["@St2"].Value = row.St2;
        if (String.IsNullOrEmpty(row.St2))
        {
            command.Parameters["@St2"].IsNullable = true;
            command.Parameters["@St2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St3", SqlDbType.VarChar, 50, "St3"));
        command.Parameters["@St3"].Value = row.St3;
        if (String.IsNullOrEmpty(row.St3))
        {
            command.Parameters["@St3"].IsNullable = true;
            command.Parameters["@St3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St4", SqlDbType.VarChar, 50, "St4"));
        command.Parameters["@St4"].Value = row.St4;
        if (String.IsNullOrEmpty(row.St4))
        {
            command.Parameters["@St4"].IsNullable = true;
            command.Parameters["@St4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St5", SqlDbType.VarChar, 50, "St5"));
        command.Parameters["@St5"].Value = row.St5;
        if (String.IsNullOrEmpty(row.St5))
        {
            command.Parameters["@St5"].IsNullable = true;
            command.Parameters["@St5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St6", SqlDbType.VarChar, 50, "St6"));
        command.Parameters["@St6"].Value = row.St6;
        if (String.IsNullOrEmpty(row.St6))
        {
            command.Parameters["@St6"].IsNullable = true;
            command.Parameters["@St6"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St_Memo", SqlDbType.VarChar, 250, "St_Memo"));
        command.Parameters["@St_Memo"].Value = row.St_Memo;
        if (String.IsNullOrEmpty(row.St_Memo))
        {
            command.Parameters["@St_Memo"].IsNullable = true;
            command.Parameters["@St_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_St", SqlDbType.VarChar, 8, "StartDate_St"));
        command.Parameters["@StartDate_St"].Value = row.StartDate_St;
        if (String.IsNullOrEmpty(row.StartDate_St))
        {
            command.Parameters["@StartDate_St"].IsNullable = true;
            command.Parameters["@StartDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_St", SqlDbType.VarChar, 8, "EndDate_St"));
        command.Parameters["@EndDate_St"].Value = row.EndDate_St;
        if (String.IsNullOrEmpty(row.EndDate_St))
        {
            command.Parameters["@EndDate_St"].IsNullable = true;
            command.Parameters["@EndDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_St", SqlDbType.VarChar, 8, "AcDate_St"));
        command.Parameters["@AcDate_St"].Value = row.AcDate_St;
        if (String.IsNullOrEmpty(row.AcDate_St))
        {
            command.Parameters["@AcDate_St"].IsNullable = true;
            command.Parameters["@AcDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate1", SqlDbType.VarChar, 8, "ClDate1"));
        command.Parameters["@ClDate1"].Value = row.ClDate1;
        if (String.IsNullOrEmpty(row.ClDate1))
        {
            command.Parameters["@ClDate1"].IsNullable = true;
            command.Parameters["@ClDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate2", SqlDbType.VarChar, 8, "ClDate2"));
        command.Parameters["@ClDate2"].Value = row.ClDate2;
        if (String.IsNullOrEmpty(row.ClDate2))
        {
            command.Parameters["@ClDate2"].IsNullable = true;
            command.Parameters["@ClDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate3", SqlDbType.VarChar, 8, "ClDate3"));
        command.Parameters["@ClDate3"].Value = row.ClDate3;
        if (String.IsNullOrEmpty(row.ClDate3))
        {
            command.Parameters["@ClDate3"].IsNullable = true;
            command.Parameters["@ClDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Cl_Memo", SqlDbType.VarChar, 250, "Cl_Memo"));
        command.Parameters["@Cl_Memo"].Value = row.Cl_Memo;
        if (String.IsNullOrEmpty(row.Cl_Memo))
        {
            command.Parameters["@Cl_Memo"].IsNullable = true;
            command.Parameters["@Cl_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Cl", SqlDbType.VarChar, 8, "StartDate_Cl"));
        command.Parameters["@StartDate_Cl"].Value = row.StartDate_Cl;
        if (String.IsNullOrEmpty(row.StartDate_Cl))
        {
            command.Parameters["@StartDate_Cl"].IsNullable = true;
            command.Parameters["@StartDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Cl", SqlDbType.VarChar, 8, "EndDate_Cl"));
        command.Parameters["@EndDate_Cl"].Value = row.EndDate_Cl;
        if (String.IsNullOrEmpty(row.EndDate_Cl))
        {
            command.Parameters["@EndDate_Cl"].IsNullable = true;
            command.Parameters["@EndDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Cl", SqlDbType.VarChar, 8, "AcDate_Cl"));
        command.Parameters["@AcDate_Cl"].Value = row.AcDate_Cl;
        if (String.IsNullOrEmpty(row.AcDate_Cl))
        {
            command.Parameters["@AcDate_Cl"].IsNullable = true;
            command.Parameters["@AcDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate3", SqlDbType.VarChar, 8, "PayDate3"));
        command.Parameters["@PayDate3"].Value = row.PayDate3;
        if (String.IsNullOrEmpty(row.PayDate3))
        {
            command.Parameters["@PayDate3"].IsNullable = true;
            command.Parameters["@PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate3_Memo", SqlDbType.VarChar, 50, "PayDate3_Memo"));
        command.Parameters["@PayDate3_Memo"].Value = row.PayDate3_Memo;
        if (String.IsNullOrEmpty(row.PayDate3_Memo))
        {
            command.Parameters["@PayDate3_Memo"].IsNullable = true;
            command.Parameters["@PayDate3_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate3", SqlDbType.VarChar, 8, "StartDate_PayDate3"));
        command.Parameters["@StartDate_PayDate3"].Value = row.StartDate_PayDate3;
        if (String.IsNullOrEmpty(row.StartDate_PayDate3))
        {
            command.Parameters["@StartDate_PayDate3"].IsNullable = true;
            command.Parameters["@StartDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate3", SqlDbType.VarChar, 8, "EndDate_PayDate3"));
        command.Parameters["@EndDate_PayDate3"].Value = row.EndDate_PayDate3;
        if (String.IsNullOrEmpty(row.EndDate_PayDate3))
        {
            command.Parameters["@EndDate_PayDate3"].IsNullable = true;
            command.Parameters["@EndDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate3", SqlDbType.VarChar, 8, "AcDate_PayDate3"));
        command.Parameters["@AcDate_PayDate3"].Value = row.AcDate_PayDate3;
        if (String.IsNullOrEmpty(row.AcDate_PayDate3))
        {
            command.Parameters["@AcDate_PayDate3"].IsNullable = true;
            command.Parameters["@AcDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.VarChar, 250, "Remark"));
        command.Parameters["@Remark"].Value = row.Remark;
        if (String.IsNullOrEmpty(row.Remark))
        {
            command.Parameters["@Remark"].IsNullable = true;
            command.Parameters["@Remark"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
        command.Parameters["@CreateDate"].Value = row.CreateDate;
        if (String.IsNullOrEmpty(row.CreateDate))
        {
            command.Parameters["@CreateDate"].IsNullable = true;
            command.Parameters["@CreateDate"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 6, "CreateTime"));
        command.Parameters["@CreateTime"].Value = row.CreateTime;
        if (String.IsNullOrEmpty(row.CreateTime))
        {
            command.Parameters["@CreateTime"].IsNullable = true;
            command.Parameters["@CreateTime"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateUser", SqlDbType.VarChar, 20, "CreateUser"));
        command.Parameters["@CreateUser"].Value = row.CreateUser;
        if (String.IsNullOrEmpty(row.CreateUser))
        {
            command.Parameters["@CreateUser"].IsNullable = true;
            command.Parameters["@CreateUser"].Value = "";
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
        command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
        command.Parameters["@UpdateUser"].Value = row.UpdateUser;
        if (String.IsNullOrEmpty(row.UpdateUser))
        {
            command.Parameters["@UpdateUser"].IsNullable = true;
            command.Parameters["@UpdateUser"].Value = "";
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

    // ACME_TASK_Solar Update
    public static void UpdateACME_TASK_Solar(ACME_TASK_Solar row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "UPDATE ACME_TASK_Solar SET PrjName = @PrjName,CurrentStage=@CurrentStage,PrjPercent=@PrjPercent,CardCode = @CardCode,CardName = @CardName,SlpCode = @SlpCode,SlpName = @SlpName,SignDate = @SignDate,SignDate_Memo = @SignDate_Memo,StartDate_Sign = @StartDate_Sign,EndDate_Sign = @EndDate_Sign,AcDate_Sign = @AcDate_Sign,PayDate1 = @PayDate1,PayDate1_Memo = @PayDate1_Memo,StartDate_PayDate1 = @StartDate_PayDate1,EndDate_PayDate1 = @EndDate_PayDate1,AcDate_PayDate1 = @AcDate_PayDate1,Ep1 = @Ep1,Ep2 = @Ep2,Ep3 = @Ep3,EpDate1 = @EpDate1,EpDate2 = @EpDate2,Ep_Memo = @Ep_Memo,StartDate_Ep = @StartDate_Ep,EndDate_Ep = @EndDate_Ep,AcDate_Ep = @AcDate_Ep,PayDate2 = @PayDate2,PayDate2_Memo = @PayDate2_Memo,StartDate_PayDate2 = @StartDate_PayDate2,EndDate_PayDate2 = @EndDate_PayDate2,AcDate_PayDate2 = @AcDate_PayDate2,Fi1 = @Fi1,Fi2 = @Fi2,Fi3 = @Fi3,FiDate1 = @FiDate1,FiDate2 = @FiDate2,Fi_Memo = @Fi_Memo,StartDate_Fi = @StartDate_Fi,EndDate_Fi = @EndDate_Fi,AcDate_Fi = @AcDate_Fi,Pr1 = @Pr1,Pr1Qty = @Pr1Qty,Pr2 = @Pr2,Pr2Qty = @Pr2Qty,Pr3 = @Pr3,Pr3Qty = @Pr3Qty,Pr4 = @Pr4,Pr4Qty = @Pr4Qty,Pr5 = @Pr5,Pr5Qty = @Pr5Qty,Pr_Memo = @Pr_Memo,StartDate_Pr = @StartDate_Pr,EndDate_Pr = @EndDate_Pr,AcDate_Pr = @AcDate_Pr,WkDate1 = @WkDate1,WkDate2 = @WkDate2,WkDate3 = @WkDate3,WkDate4 = @WkDate4,WkDate5 = @WkDate5,Wk_Memo = @Wk_Memo,StartDate_wk = @StartDate_wk,EndDate_wk = @EndDate_wk,AcDate_wk = @AcDate_wk,St1 = @St1,St2 = @St2,St3 = @St3,St4 = @St4,St5 = @St5,St6 = @St6,St_Memo = @St_Memo,StartDate_St = @StartDate_St,EndDate_St = @EndDate_St,AcDate_St = @AcDate_St,ClDate1 = @ClDate1,ClDate2 = @ClDate2,ClDate3 = @ClDate3,Cl_Memo = @Cl_Memo,StartDate_Cl = @StartDate_Cl,EndDate_Cl = @EndDate_Cl,AcDate_Cl = @AcDate_Cl,PayDate3 = @PayDate3,PayDate3_Memo = @PayDate3_Memo,StartDate_PayDate3 = @StartDate_PayDate3,EndDate_PayDate3 = @EndDate_PayDate3,AcDate_PayDate3 = @AcDate_PayDate3,Remark = @Remark,CreateDate = @CreateDate,CreateTime = @CreateTime,CreateUser = @CreateUser,UpdateDate = @UpdateDate,UpdateTime = @UpdateTime,UpdateUser = @UpdateUser WHERE PrjCode=@PrjCode";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
        command.Parameters["@PrjCode"].Value = row.PrjCode;
        if (String.IsNullOrEmpty(row.PrjCode))
        {
            command.Parameters["@PrjCode"].IsNullable = true;
            command.Parameters["@PrjCode"].Value = "";
        }

        command.Parameters.Add(new SqlParameter("@CurrentStage", SqlDbType.VarChar, 50, "CurrentStage"));
        command.Parameters["@CurrentStage"].Value = row.CurrentStage;
        if (String.IsNullOrEmpty(row.CurrentStage))
        {
            command.Parameters["@CurrentStage"].IsNullable = true;
            command.Parameters["@CurrentStage"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PrjPercent", SqlDbType.VarChar, 10, "PrjPercent"));
        command.Parameters["@PrjPercent"].Value = row.PrjPercent;
        if (String.IsNullOrEmpty(row.PrjPercent))
        {
            command.Parameters["@PrjPercent"].IsNullable = true;
            command.Parameters["@PrjPercent"].Value = "";
        }

        command.Parameters.Add(new SqlParameter("@PrjName", SqlDbType.VarChar, 50, "PrjName"));
        command.Parameters["@PrjName"].Value = row.PrjName;
        if (String.IsNullOrEmpty(row.PrjName))
        {
            command.Parameters["@PrjName"].IsNullable = true;
            command.Parameters["@PrjName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CardCode", SqlDbType.VarChar, 50, "CardCode"));
        command.Parameters["@CardCode"].Value = row.CardCode;
        if (String.IsNullOrEmpty(row.CardCode))
        {
            command.Parameters["@CardCode"].IsNullable = true;
            command.Parameters["@CardCode"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CardName", SqlDbType.VarChar, 50, "CardName"));
        command.Parameters["@CardName"].Value = row.CardName;
        if (String.IsNullOrEmpty(row.CardName))
        {
            command.Parameters["@CardName"].IsNullable = true;
            command.Parameters["@CardName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SlpCode", SqlDbType.VarChar, 50, "SlpCode"));
        command.Parameters["@SlpCode"].Value = row.SlpCode;
        if (String.IsNullOrEmpty(row.SlpCode))
        {
            command.Parameters["@SlpCode"].IsNullable = true;
            command.Parameters["@SlpCode"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SlpName", SqlDbType.VarChar, 50, "SlpName"));
        command.Parameters["@SlpName"].Value = row.SlpName;
        if (String.IsNullOrEmpty(row.SlpName))
        {
            command.Parameters["@SlpName"].IsNullable = true;
            command.Parameters["@SlpName"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SignDate", SqlDbType.VarChar, 8, "SignDate"));
        command.Parameters["@SignDate"].Value = row.SignDate;
        if (String.IsNullOrEmpty(row.SignDate))
        {
            command.Parameters["@SignDate"].IsNullable = true;
            command.Parameters["@SignDate"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SignDate_Memo", SqlDbType.VarChar, 250, "SignDate_Memo"));
        command.Parameters["@SignDate_Memo"].Value = row.SignDate_Memo;
        if (String.IsNullOrEmpty(row.SignDate_Memo))
        {
            command.Parameters["@SignDate_Memo"].IsNullable = true;
            command.Parameters["@SignDate_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Sign", SqlDbType.VarChar, 8, "StartDate_Sign"));
        command.Parameters["@StartDate_Sign"].Value = row.StartDate_Sign;
        if (String.IsNullOrEmpty(row.StartDate_Sign))
        {
            command.Parameters["@StartDate_Sign"].IsNullable = true;
            command.Parameters["@StartDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Sign", SqlDbType.VarChar, 8, "EndDate_Sign"));
        command.Parameters["@EndDate_Sign"].Value = row.EndDate_Sign;
        if (String.IsNullOrEmpty(row.EndDate_Sign))
        {
            command.Parameters["@EndDate_Sign"].IsNullable = true;
            command.Parameters["@EndDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Sign", SqlDbType.VarChar, 8, "AcDate_Sign"));
        command.Parameters["@AcDate_Sign"].Value = row.AcDate_Sign;
        if (String.IsNullOrEmpty(row.AcDate_Sign))
        {
            command.Parameters["@AcDate_Sign"].IsNullable = true;
            command.Parameters["@AcDate_Sign"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate1", SqlDbType.VarChar, 8, "PayDate1"));
        command.Parameters["@PayDate1"].Value = row.PayDate1;
        if (String.IsNullOrEmpty(row.PayDate1))
        {
            command.Parameters["@PayDate1"].IsNullable = true;
            command.Parameters["@PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate1_Memo", SqlDbType.VarChar, 250, "PayDate1_Memo"));
        command.Parameters["@PayDate1_Memo"].Value = row.PayDate1_Memo;
        if (String.IsNullOrEmpty(row.PayDate1_Memo))
        {
            command.Parameters["@PayDate1_Memo"].IsNullable = true;
            command.Parameters["@PayDate1_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate1", SqlDbType.VarChar, 8, "StartDate_PayDate1"));
        command.Parameters["@StartDate_PayDate1"].Value = row.StartDate_PayDate1;
        if (String.IsNullOrEmpty(row.StartDate_PayDate1))
        {
            command.Parameters["@StartDate_PayDate1"].IsNullable = true;
            command.Parameters["@StartDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate1", SqlDbType.VarChar, 8, "EndDate_PayDate1"));
        command.Parameters["@EndDate_PayDate1"].Value = row.EndDate_PayDate1;
        if (String.IsNullOrEmpty(row.EndDate_PayDate1))
        {
            command.Parameters["@EndDate_PayDate1"].IsNullable = true;
            command.Parameters["@EndDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate1", SqlDbType.VarChar, 8, "AcDate_PayDate1"));
        command.Parameters["@AcDate_PayDate1"].Value = row.AcDate_PayDate1;
        if (String.IsNullOrEmpty(row.AcDate_PayDate1))
        {
            command.Parameters["@AcDate_PayDate1"].IsNullable = true;
            command.Parameters["@AcDate_PayDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep1", SqlDbType.VarChar, 50, "Ep1"));
        command.Parameters["@Ep1"].Value = row.Ep1;
        if (String.IsNullOrEmpty(row.Ep1))
        {
            command.Parameters["@Ep1"].IsNullable = true;
            command.Parameters["@Ep1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep2", SqlDbType.VarChar, 50, "Ep2"));
        command.Parameters["@Ep2"].Value = row.Ep2;
        if (String.IsNullOrEmpty(row.Ep2))
        {
            command.Parameters["@Ep2"].IsNullable = true;
            command.Parameters["@Ep2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep3", SqlDbType.VarChar, 50, "Ep3"));
        command.Parameters["@Ep3"].Value = row.Ep3;
        if (String.IsNullOrEmpty(row.Ep3))
        {
            command.Parameters["@Ep3"].IsNullable = true;
            command.Parameters["@Ep3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EpDate1", SqlDbType.VarChar, 8, "EpDate1"));
        command.Parameters["@EpDate1"].Value = row.EpDate1;
        if (String.IsNullOrEmpty(row.EpDate1))
        {
            command.Parameters["@EpDate1"].IsNullable = true;
            command.Parameters["@EpDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EpDate2", SqlDbType.VarChar, 8, "EpDate2"));
        command.Parameters["@EpDate2"].Value = row.EpDate2;
        if (String.IsNullOrEmpty(row.EpDate2))
        {
            command.Parameters["@EpDate2"].IsNullable = true;
            command.Parameters["@EpDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Ep_Memo", SqlDbType.VarChar, 250, "Ep_Memo"));
        command.Parameters["@Ep_Memo"].Value = row.Ep_Memo;
        if (String.IsNullOrEmpty(row.Ep_Memo))
        {
            command.Parameters["@Ep_Memo"].IsNullable = true;
            command.Parameters["@Ep_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Ep", SqlDbType.VarChar, 8, "StartDate_Ep"));
        command.Parameters["@StartDate_Ep"].Value = row.StartDate_Ep;
        if (String.IsNullOrEmpty(row.StartDate_Ep))
        {
            command.Parameters["@StartDate_Ep"].IsNullable = true;
            command.Parameters["@StartDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Ep", SqlDbType.VarChar, 8, "EndDate_Ep"));
        command.Parameters["@EndDate_Ep"].Value = row.EndDate_Ep;
        if (String.IsNullOrEmpty(row.EndDate_Ep))
        {
            command.Parameters["@EndDate_Ep"].IsNullable = true;
            command.Parameters["@EndDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Ep", SqlDbType.VarChar, 8, "AcDate_Ep"));
        command.Parameters["@AcDate_Ep"].Value = row.AcDate_Ep;
        if (String.IsNullOrEmpty(row.AcDate_Ep))
        {
            command.Parameters["@AcDate_Ep"].IsNullable = true;
            command.Parameters["@AcDate_Ep"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate2", SqlDbType.VarChar, 8, "PayDate2"));
        command.Parameters["@PayDate2"].Value = row.PayDate2;
        if (String.IsNullOrEmpty(row.PayDate2))
        {
            command.Parameters["@PayDate2"].IsNullable = true;
            command.Parameters["@PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate2_Memo", SqlDbType.VarChar, 250, "PayDate2_Memo"));
        command.Parameters["@PayDate2_Memo"].Value = row.PayDate2_Memo;
        if (String.IsNullOrEmpty(row.PayDate2_Memo))
        {
            command.Parameters["@PayDate2_Memo"].IsNullable = true;
            command.Parameters["@PayDate2_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate2", SqlDbType.VarChar, 8, "StartDate_PayDate2"));
        command.Parameters["@StartDate_PayDate2"].Value = row.StartDate_PayDate2;
        if (String.IsNullOrEmpty(row.StartDate_PayDate2))
        {
            command.Parameters["@StartDate_PayDate2"].IsNullable = true;
            command.Parameters["@StartDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate2", SqlDbType.VarChar, 8, "EndDate_PayDate2"));
        command.Parameters["@EndDate_PayDate2"].Value = row.EndDate_PayDate2;
        if (String.IsNullOrEmpty(row.EndDate_PayDate2))
        {
            command.Parameters["@EndDate_PayDate2"].IsNullable = true;
            command.Parameters["@EndDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate2", SqlDbType.VarChar, 8, "AcDate_PayDate2"));
        command.Parameters["@AcDate_PayDate2"].Value = row.AcDate_PayDate2;
        if (String.IsNullOrEmpty(row.AcDate_PayDate2))
        {
            command.Parameters["@AcDate_PayDate2"].IsNullable = true;
            command.Parameters["@AcDate_PayDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi1", SqlDbType.VarChar, 50, "Fi1"));
        command.Parameters["@Fi1"].Value = row.Fi1;
        if (String.IsNullOrEmpty(row.Fi1))
        {
            command.Parameters["@Fi1"].IsNullable = true;
            command.Parameters["@Fi1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi2", SqlDbType.VarChar, 50, "Fi2"));
        command.Parameters["@Fi2"].Value = row.Fi2;
        if (String.IsNullOrEmpty(row.Fi2))
        {
            command.Parameters["@Fi2"].IsNullable = true;
            command.Parameters["@Fi2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi3", SqlDbType.VarChar, 50, "Fi3"));
        command.Parameters["@Fi3"].Value = row.Fi3;
        if (String.IsNullOrEmpty(row.Fi3))
        {
            command.Parameters["@Fi3"].IsNullable = true;
            command.Parameters["@Fi3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@FiDate1", SqlDbType.VarChar, 8, "FiDate1"));
        command.Parameters["@FiDate1"].Value = row.FiDate1;
        if (String.IsNullOrEmpty(row.FiDate1))
        {
            command.Parameters["@FiDate1"].IsNullable = true;
            command.Parameters["@FiDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@FiDate2", SqlDbType.VarChar, 8, "FiDate2"));
        command.Parameters["@FiDate2"].Value = row.FiDate2;
        if (String.IsNullOrEmpty(row.FiDate2))
        {
            command.Parameters["@FiDate2"].IsNullable = true;
            command.Parameters["@FiDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Fi_Memo", SqlDbType.VarChar, 250, "Fi_Memo"));
        command.Parameters["@Fi_Memo"].Value = row.Fi_Memo;
        if (String.IsNullOrEmpty(row.Fi_Memo))
        {
            command.Parameters["@Fi_Memo"].IsNullable = true;
            command.Parameters["@Fi_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Fi", SqlDbType.VarChar, 8, "StartDate_Fi"));
        command.Parameters["@StartDate_Fi"].Value = row.StartDate_Fi;
        if (String.IsNullOrEmpty(row.StartDate_Fi))
        {
            command.Parameters["@StartDate_Fi"].IsNullable = true;
            command.Parameters["@StartDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Fi", SqlDbType.VarChar, 8, "EndDate_Fi"));
        command.Parameters["@EndDate_Fi"].Value = row.EndDate_Fi;
        if (String.IsNullOrEmpty(row.EndDate_Fi))
        {
            command.Parameters["@EndDate_Fi"].IsNullable = true;
            command.Parameters["@EndDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Fi", SqlDbType.VarChar, 8, "AcDate_Fi"));
        command.Parameters["@AcDate_Fi"].Value = row.AcDate_Fi;
        if (String.IsNullOrEmpty(row.AcDate_Fi))
        {
            command.Parameters["@AcDate_Fi"].IsNullable = true;
            command.Parameters["@AcDate_Fi"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr1", SqlDbType.VarChar, 50, "Pr1"));
        command.Parameters["@Pr1"].Value = row.Pr1;
        if (String.IsNullOrEmpty(row.Pr1))
        {
            command.Parameters["@Pr1"].IsNullable = true;
            command.Parameters["@Pr1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr1Qty", SqlDbType.VarChar, 50, "Pr1Qty"));
        command.Parameters["@Pr1Qty"].Value = row.Pr1Qty;
        if (String.IsNullOrEmpty(row.Pr1Qty))
        {
            command.Parameters["@Pr1Qty"].IsNullable = true;
            command.Parameters["@Pr1Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr2", SqlDbType.VarChar, 50, "Pr2"));
        command.Parameters["@Pr2"].Value = row.Pr2;
        if (String.IsNullOrEmpty(row.Pr2))
        {
            command.Parameters["@Pr2"].IsNullable = true;
            command.Parameters["@Pr2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr2Qty", SqlDbType.VarChar, 50, "Pr2Qty"));
        command.Parameters["@Pr2Qty"].Value = row.Pr2Qty;
        if (String.IsNullOrEmpty(row.Pr2Qty))
        {
            command.Parameters["@Pr2Qty"].IsNullable = true;
            command.Parameters["@Pr2Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr3", SqlDbType.VarChar, 50, "Pr3"));
        command.Parameters["@Pr3"].Value = row.Pr3;
        if (String.IsNullOrEmpty(row.Pr3))
        {
            command.Parameters["@Pr3"].IsNullable = true;
            command.Parameters["@Pr3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr3Qty", SqlDbType.VarChar, 50, "Pr3Qty"));
        command.Parameters["@Pr3Qty"].Value = row.Pr3Qty;
        if (String.IsNullOrEmpty(row.Pr3Qty))
        {
            command.Parameters["@Pr3Qty"].IsNullable = true;
            command.Parameters["@Pr3Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr4", SqlDbType.VarChar, 50, "Pr4"));
        command.Parameters["@Pr4"].Value = row.Pr4;
        if (String.IsNullOrEmpty(row.Pr4))
        {
            command.Parameters["@Pr4"].IsNullable = true;
            command.Parameters["@Pr4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr4Qty", SqlDbType.VarChar, 50, "Pr4Qty"));
        command.Parameters["@Pr4Qty"].Value = row.Pr4Qty;
        if (String.IsNullOrEmpty(row.Pr4Qty))
        {
            command.Parameters["@Pr4Qty"].IsNullable = true;
            command.Parameters["@Pr4Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr5", SqlDbType.VarChar, 50, "Pr5"));
        command.Parameters["@Pr5"].Value = row.Pr5;
        if (String.IsNullOrEmpty(row.Pr5))
        {
            command.Parameters["@Pr5"].IsNullable = true;
            command.Parameters["@Pr5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr5Qty", SqlDbType.VarChar, 50, "Pr5Qty"));
        command.Parameters["@Pr5Qty"].Value = row.Pr5Qty;
        if (String.IsNullOrEmpty(row.Pr5Qty))
        {
            command.Parameters["@Pr5Qty"].IsNullable = true;
            command.Parameters["@Pr5Qty"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Pr_Memo", SqlDbType.VarChar, 250, "Pr_Memo"));
        command.Parameters["@Pr_Memo"].Value = row.Pr_Memo;
        if (String.IsNullOrEmpty(row.Pr_Memo))
        {
            command.Parameters["@Pr_Memo"].IsNullable = true;
            command.Parameters["@Pr_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Pr", SqlDbType.VarChar, 8, "StartDate_Pr"));
        command.Parameters["@StartDate_Pr"].Value = row.StartDate_Pr;
        if (String.IsNullOrEmpty(row.StartDate_Pr))
        {
            command.Parameters["@StartDate_Pr"].IsNullable = true;
            command.Parameters["@StartDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Pr", SqlDbType.VarChar, 8, "EndDate_Pr"));
        command.Parameters["@EndDate_Pr"].Value = row.EndDate_Pr;
        if (String.IsNullOrEmpty(row.EndDate_Pr))
        {
            command.Parameters["@EndDate_Pr"].IsNullable = true;
            command.Parameters["@EndDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Pr", SqlDbType.VarChar, 8, "AcDate_Pr"));
        command.Parameters["@AcDate_Pr"].Value = row.AcDate_Pr;
        if (String.IsNullOrEmpty(row.AcDate_Pr))
        {
            command.Parameters["@AcDate_Pr"].IsNullable = true;
            command.Parameters["@AcDate_Pr"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate1", SqlDbType.VarChar, 8, "WkDate1"));
        command.Parameters["@WkDate1"].Value = row.WkDate1;
        if (String.IsNullOrEmpty(row.WkDate1))
        {
            command.Parameters["@WkDate1"].IsNullable = true;
            command.Parameters["@WkDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate2", SqlDbType.VarChar, 8, "WkDate2"));
        command.Parameters["@WkDate2"].Value = row.WkDate2;
        if (String.IsNullOrEmpty(row.WkDate2))
        {
            command.Parameters["@WkDate2"].IsNullable = true;
            command.Parameters["@WkDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate3", SqlDbType.VarChar, 8, "WkDate3"));
        command.Parameters["@WkDate3"].Value = row.WkDate3;
        if (String.IsNullOrEmpty(row.WkDate3))
        {
            command.Parameters["@WkDate3"].IsNullable = true;
            command.Parameters["@WkDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate4", SqlDbType.VarChar, 8, "WkDate4"));
        command.Parameters["@WkDate4"].Value = row.WkDate4;
        if (String.IsNullOrEmpty(row.WkDate4))
        {
            command.Parameters["@WkDate4"].IsNullable = true;
            command.Parameters["@WkDate4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@WkDate5", SqlDbType.VarChar, 8, "WkDate5"));
        command.Parameters["@WkDate5"].Value = row.WkDate5;
        if (String.IsNullOrEmpty(row.WkDate5))
        {
            command.Parameters["@WkDate5"].IsNullable = true;
            command.Parameters["@WkDate5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Wk_Memo", SqlDbType.VarChar, 250, "Wk_Memo"));
        command.Parameters["@Wk_Memo"].Value = row.Wk_Memo;
        if (String.IsNullOrEmpty(row.Wk_Memo))
        {
            command.Parameters["@Wk_Memo"].IsNullable = true;
            command.Parameters["@Wk_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_wk", SqlDbType.VarChar, 8, "StartDate_wk"));
        command.Parameters["@StartDate_wk"].Value = row.StartDate_wk;
        if (String.IsNullOrEmpty(row.StartDate_wk))
        {
            command.Parameters["@StartDate_wk"].IsNullable = true;
            command.Parameters["@StartDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_wk", SqlDbType.VarChar, 8, "EndDate_wk"));
        command.Parameters["@EndDate_wk"].Value = row.EndDate_wk;
        if (String.IsNullOrEmpty(row.EndDate_wk))
        {
            command.Parameters["@EndDate_wk"].IsNullable = true;
            command.Parameters["@EndDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_wk", SqlDbType.VarChar, 8, "AcDate_wk"));
        command.Parameters["@AcDate_wk"].Value = row.AcDate_wk;
        if (String.IsNullOrEmpty(row.AcDate_wk))
        {
            command.Parameters["@AcDate_wk"].IsNullable = true;
            command.Parameters["@AcDate_wk"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St1", SqlDbType.VarChar, 50, "St1"));
        command.Parameters["@St1"].Value = row.St1;
        if (String.IsNullOrEmpty(row.St1))
        {
            command.Parameters["@St1"].IsNullable = true;
            command.Parameters["@St1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St2", SqlDbType.VarChar, 50, "St2"));
        command.Parameters["@St2"].Value = row.St2;
        if (String.IsNullOrEmpty(row.St2))
        {
            command.Parameters["@St2"].IsNullable = true;
            command.Parameters["@St2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St3", SqlDbType.VarChar, 50, "St3"));
        command.Parameters["@St3"].Value = row.St3;
        if (String.IsNullOrEmpty(row.St3))
        {
            command.Parameters["@St3"].IsNullable = true;
            command.Parameters["@St3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St4", SqlDbType.VarChar, 50, "St4"));
        command.Parameters["@St4"].Value = row.St4;
        if (String.IsNullOrEmpty(row.St4))
        {
            command.Parameters["@St4"].IsNullable = true;
            command.Parameters["@St4"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St5", SqlDbType.VarChar, 50, "St5"));
        command.Parameters["@St5"].Value = row.St5;
        if (String.IsNullOrEmpty(row.St5))
        {
            command.Parameters["@St5"].IsNullable = true;
            command.Parameters["@St5"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St6", SqlDbType.VarChar, 50, "St6"));
        command.Parameters["@St6"].Value = row.St6;
        if (String.IsNullOrEmpty(row.St6))
        {
            command.Parameters["@St6"].IsNullable = true;
            command.Parameters["@St6"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@St_Memo", SqlDbType.VarChar, 250, "St_Memo"));
        command.Parameters["@St_Memo"].Value = row.St_Memo;
        if (String.IsNullOrEmpty(row.St_Memo))
        {
            command.Parameters["@St_Memo"].IsNullable = true;
            command.Parameters["@St_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_St", SqlDbType.VarChar, 8, "StartDate_St"));
        command.Parameters["@StartDate_St"].Value = row.StartDate_St;
        if (String.IsNullOrEmpty(row.StartDate_St))
        {
            command.Parameters["@StartDate_St"].IsNullable = true;
            command.Parameters["@StartDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_St", SqlDbType.VarChar, 8, "EndDate_St"));
        command.Parameters["@EndDate_St"].Value = row.EndDate_St;
        if (String.IsNullOrEmpty(row.EndDate_St))
        {
            command.Parameters["@EndDate_St"].IsNullable = true;
            command.Parameters["@EndDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_St", SqlDbType.VarChar, 8, "AcDate_St"));
        command.Parameters["@AcDate_St"].Value = row.AcDate_St;
        if (String.IsNullOrEmpty(row.AcDate_St))
        {
            command.Parameters["@AcDate_St"].IsNullable = true;
            command.Parameters["@AcDate_St"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate1", SqlDbType.VarChar, 8, "ClDate1"));
        command.Parameters["@ClDate1"].Value = row.ClDate1;
        if (String.IsNullOrEmpty(row.ClDate1))
        {
            command.Parameters["@ClDate1"].IsNullable = true;
            command.Parameters["@ClDate1"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate2", SqlDbType.VarChar, 8, "ClDate2"));
        command.Parameters["@ClDate2"].Value = row.ClDate2;
        if (String.IsNullOrEmpty(row.ClDate2))
        {
            command.Parameters["@ClDate2"].IsNullable = true;
            command.Parameters["@ClDate2"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ClDate3", SqlDbType.VarChar, 8, "ClDate3"));
        command.Parameters["@ClDate3"].Value = row.ClDate3;
        if (String.IsNullOrEmpty(row.ClDate3))
        {
            command.Parameters["@ClDate3"].IsNullable = true;
            command.Parameters["@ClDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Cl_Memo", SqlDbType.VarChar, 250, "Cl_Memo"));
        command.Parameters["@Cl_Memo"].Value = row.Cl_Memo;
        if (String.IsNullOrEmpty(row.Cl_Memo))
        {
            command.Parameters["@Cl_Memo"].IsNullable = true;
            command.Parameters["@Cl_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_Cl", SqlDbType.VarChar, 8, "StartDate_Cl"));
        command.Parameters["@StartDate_Cl"].Value = row.StartDate_Cl;
        if (String.IsNullOrEmpty(row.StartDate_Cl))
        {
            command.Parameters["@StartDate_Cl"].IsNullable = true;
            command.Parameters["@StartDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_Cl", SqlDbType.VarChar, 8, "EndDate_Cl"));
        command.Parameters["@EndDate_Cl"].Value = row.EndDate_Cl;
        if (String.IsNullOrEmpty(row.EndDate_Cl))
        {
            command.Parameters["@EndDate_Cl"].IsNullable = true;
            command.Parameters["@EndDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_Cl", SqlDbType.VarChar, 8, "AcDate_Cl"));
        command.Parameters["@AcDate_Cl"].Value = row.AcDate_Cl;
        if (String.IsNullOrEmpty(row.AcDate_Cl))
        {
            command.Parameters["@AcDate_Cl"].IsNullable = true;
            command.Parameters["@AcDate_Cl"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate3", SqlDbType.VarChar, 8, "PayDate3"));
        command.Parameters["@PayDate3"].Value = row.PayDate3;
        if (String.IsNullOrEmpty(row.PayDate3))
        {
            command.Parameters["@PayDate3"].IsNullable = true;
            command.Parameters["@PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@PayDate3_Memo", SqlDbType.VarChar, 50, "PayDate3_Memo"));
        command.Parameters["@PayDate3_Memo"].Value = row.PayDate3_Memo;
        if (String.IsNullOrEmpty(row.PayDate3_Memo))
        {
            command.Parameters["@PayDate3_Memo"].IsNullable = true;
            command.Parameters["@PayDate3_Memo"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate_PayDate3", SqlDbType.VarChar, 8, "StartDate_PayDate3"));
        command.Parameters["@StartDate_PayDate3"].Value = row.StartDate_PayDate3;
        if (String.IsNullOrEmpty(row.StartDate_PayDate3))
        {
            command.Parameters["@StartDate_PayDate3"].IsNullable = true;
            command.Parameters["@StartDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@EndDate_PayDate3", SqlDbType.VarChar, 8, "EndDate_PayDate3"));
        command.Parameters["@EndDate_PayDate3"].Value = row.EndDate_PayDate3;
        if (String.IsNullOrEmpty(row.EndDate_PayDate3))
        {
            command.Parameters["@EndDate_PayDate3"].IsNullable = true;
            command.Parameters["@EndDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@AcDate_PayDate3", SqlDbType.VarChar, 8, "AcDate_PayDate3"));
        command.Parameters["@AcDate_PayDate3"].Value = row.AcDate_PayDate3;
        if (String.IsNullOrEmpty(row.AcDate_PayDate3))
        {
            command.Parameters["@AcDate_PayDate3"].IsNullable = true;
            command.Parameters["@AcDate_PayDate3"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.VarChar, 250, "Remark"));
        command.Parameters["@Remark"].Value = row.Remark;
        if (String.IsNullOrEmpty(row.Remark))
        {
            command.Parameters["@Remark"].IsNullable = true;
            command.Parameters["@Remark"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
        command.Parameters["@CreateDate"].Value = row.CreateDate;
        if (String.IsNullOrEmpty(row.CreateDate))
        {
            command.Parameters["@CreateDate"].IsNullable = true;
            command.Parameters["@CreateDate"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 6, "CreateTime"));
        command.Parameters["@CreateTime"].Value = row.CreateTime;
        if (String.IsNullOrEmpty(row.CreateTime))
        {
            command.Parameters["@CreateTime"].IsNullable = true;
            command.Parameters["@CreateTime"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CreateUser", SqlDbType.VarChar, 20, "CreateUser"));
        command.Parameters["@CreateUser"].Value = row.CreateUser;
        if (String.IsNullOrEmpty(row.CreateUser))
        {
            command.Parameters["@CreateUser"].IsNullable = true;
            command.Parameters["@CreateUser"].Value = "";
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
        command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
        command.Parameters["@UpdateUser"].Value = row.UpdateUser;
        if (String.IsNullOrEmpty(row.UpdateUser))
        {
            command.Parameters["@UpdateUser"].IsNullable = true;
            command.Parameters["@UpdateUser"].Value = "";
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

    // ACME_TASK_Solar Delete
    public static void DeleteACME_TASK_Solar(ACME_TASK_Solar row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "DELETE ACME_TASK_Solar WHERE PrjCode=@PrjCode";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@PrjCode", row.PrjCode));
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

    // ACME_TASK_Solar Select
    public static DataTable GetACME_TASK_Solar(ACME_TASK_Solar row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT currentStage,PrjPercent,PrjCode,PrjName,CardCode,CardName,SlpCode,SlpName,SignDate,SignDate_Memo,StartDate_Sign,EndDate_Sign,AcDate_Sign,PayDate1,PayDate1_Memo,StartDate_PayDate1,EndDate_PayDate1,AcDate_PayDate1,Ep1,Ep2,Ep3,EpDate1,EpDate2,Ep_Memo,StartDate_Ep,EndDate_Ep,AcDate_Ep,PayDate2,PayDate2_Memo,StartDate_PayDate2,EndDate_PayDate2,AcDate_PayDate2,Fi1,Fi2,Fi3,FiDate1,FiDate2,Fi_Memo,StartDate_Fi,EndDate_Fi,AcDate_Fi,Pr1,Pr1Qty,Pr2,Pr2Qty,Pr3,Pr3Qty,Pr4,Pr4Qty,Pr5,Pr5Qty,Pr_Memo,StartDate_Pr,EndDate_Pr,AcDate_Pr,WkDate1,WkDate2,WkDate3,WkDate4,WkDate5,Wk_Memo,StartDate_wk,EndDate_wk,AcDate_wk,St1,St2,St3,St4,St5,St6,St_Memo,StartDate_St,EndDate_St,AcDate_St,ClDate1,ClDate2,ClDate3,Cl_Memo,StartDate_Cl,EndDate_Cl,AcDate_Cl,PayDate3,PayDate3_Memo,StartDate_PayDate3,EndDate_PayDate3,AcDate_PayDate3,Remark,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Solar WHERE 1= 1  AND PrjCode=@PrjCode";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@PrjCode", row.PrjCode));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Solar");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Solar"];
    }

    // ACME_TASK_Solar Select
    public static DataTable GetACME_TASK_Solar(string PrjCode)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT currentStage,PrjCode,PrjName,PrjPercent,CardCode,CardName,SlpCode,SlpName,SignDate,SignDate_Memo,StartDate_Sign,EndDate_Sign,AcDate_Sign,PayDate1,PayDate1_Memo,StartDate_PayDate1,EndDate_PayDate1,AcDate_PayDate1,Ep1,Ep2,Ep3,EpDate1,EpDate2,Ep_Memo,StartDate_Ep,EndDate_Ep,AcDate_Ep,PayDate2,PayDate2_Memo,StartDate_PayDate2,EndDate_PayDate2,AcDate_PayDate2,Fi1,Fi2,Fi3,FiDate1,FiDate2,Fi_Memo,StartDate_Fi,EndDate_Fi,AcDate_Fi,Pr1,Pr1Qty,Pr2,Pr2Qty,Pr3,Pr3Qty,Pr4,Pr4Qty,Pr5,Pr5Qty,Pr_Memo,StartDate_Pr,EndDate_Pr,AcDate_Pr,WkDate1,WkDate2,WkDate3,WkDate4,WkDate5,Wk_Memo,StartDate_wk,EndDate_wk,AcDate_wk,St1,St2,St3,St4,St5,St6,St_Memo,StartDate_St,EndDate_St,AcDate_St,ClDate1,ClDate2,ClDate3,Cl_Memo,StartDate_Cl,EndDate_Cl,AcDate_Cl,PayDate3,PayDate3_Memo,StartDate_PayDate3,EndDate_PayDate3,AcDate_PayDate3,Remark,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Solar WHERE 1= 1  AND PrjCode=@PrjCode";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Solar");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Solar"];
    }
    // Condition 版本
    public static DataTable GetACME_TASK_Solar_Condition(string Condition)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT PrjCode,PrjName,PrjPercent,CardCode,CardName,SlpCode,SlpName,SignDate,SignDate_Memo,StartDate_Sign,EndDate_Sign,AcDate_Sign,PayDate1,PayDate1_Memo,StartDate_PayDate1,EndDate_PayDate1,AcDate_PayDate1,Ep1,Ep2,Ep3,EpDate1,EpDate2,Ep_Memo,StartDate_Ep,EndDate_Ep,AcDate_Ep,PayDate2,PayDate2_Memo,StartDate_PayDate2,EndDate_PayDate2,AcDate_PayDate2,Fi1,Fi2,Fi3,FiDate1,FiDate2,Fi_Memo,StartDate_Fi,EndDate_Fi,AcDate_Fi,Pr1,Pr1Qty,Pr2,Pr2Qty,Pr3,Pr3Qty,Pr4,Pr4Qty,Pr5,Pr5Qty,Pr_Memo,StartDate_Pr,EndDate_Pr,AcDate_Pr,WkDate1,WkDate2,WkDate3,WkDate4,WkDate5,Wk_Memo,StartDate_wk,EndDate_wk,AcDate_wk,St1,St2,St3,St4,St5,St6,St_Memo,StartDate_St,EndDate_St,AcDate_St,ClDate1,ClDate2,ClDate3,Cl_Memo,StartDate_Cl,EndDate_Cl,AcDate_Cl,PayDate3,PayDate3_Memo,StartDate_PayDate3,EndDate_PayDate3,AcDate_PayDate3,Remark,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Solar WHERE 1= 1 ";
        sql += Condition;
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Solar");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Solar"];
    }

}
