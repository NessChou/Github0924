namespace ACME.CRM 
{

    using ACME.CRM.CRMTableAdapters;

    public partial class CRM 
    {


        public static ACME_STAGETableAdapter MyACME_STAGE = new ACME_STAGETableAdapter();
        partial class ACME_STAGEDataTable : ITableUpdate
        {
            void ITableUpdate.Fill()
            {
                MyACME_STAGE.Connection = globals.Connection;
                MyACME_STAGE.Fill(this,globals.UserID);
            }
            void ITableUpdate.Update()
            {
                MyACME_STAGE.Update(this);
            }

        }


        public static ACME_LEADTableAdapter MyACME_LEAD = new ACME_LEADTableAdapter();
        partial class ACME_LEADDataTable : ITableUpdate
        {
           
            
            void ITableUpdate.Fill()
            {
                MyACME_LEAD.Connection = globals.Connection;
                
                MyACME_LEAD.Fill(this, globals.UserID);
            }
            void ITableUpdate.Update()
            {
                MyACME_LEAD.Update(this);
            }

        }

        public static ACME_MISTableAdapter MyACME_MIS = new ACME_MISTableAdapter();
        partial class ACME_MISDataTable : ITableUpdate
        {


            void ITableUpdate.Fill()
            {
                MyACME_MIS.Connection = globals.Connection;

                MyACME_MIS.Fill(this);
            }
            void ITableUpdate.Update()
            {
                MyACME_MIS.Update(this);
            }

        }


        public static ACME_STAGE_DTableAdapter MyACME_STAGE_D = new ACME_STAGE_DTableAdapter();
        partial class ACME_STAGE_DDataTable : ITableUpdate
        {
            void ITableUpdate.Fill()
            {
                MyACME_STAGE_D.Connection = globals.Connection;
                MyACME_STAGE_D.Fill(this);
            }
            void ITableUpdate.Update()
            {
                MyACME_STAGE_D.Update(this);
            }
            
        }
    }
     
}
