using System;
using System.ComponentModel;
using System.Collections;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Reflection;
using Microsoft.Win32;
using System.IO;
using Srvtools;
using System.Security;
using System.Security.Permissions;
using System.Threading;

namespace ServerPackage
{
    /// <summary>
    /// Summary description for Component.
    /// </summary>
    public class Component : DataModule
    {
        private ServiceManager serviceManager;
        private InfoConnection InfoConnection1;
        private InfoCommand C_company;
        private UpdateComponent ucC_company;
        private InfoCommand View_C_company;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components;

        public Component(System.ComponentModel.IContainer container)
        {
            ///
            /// Required for Windows.Forms Class Composition Designer support
            ///
            container.Add(this);
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        public Component()
        {
            ///
            /// This call is required by the Windows.Forms Designer.
            ///
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Srvtools.KeyItem keyItem1 = new Srvtools.KeyItem();
            Srvtools.FieldAttr fieldAttr1 = new Srvtools.FieldAttr();
            Srvtools.FieldAttr fieldAttr2 = new Srvtools.FieldAttr();
            Srvtools.FieldAttr fieldAttr3 = new Srvtools.FieldAttr();
            Srvtools.FieldAttr fieldAttr4 = new Srvtools.FieldAttr();
            Srvtools.FieldAttr fieldAttr5 = new Srvtools.FieldAttr();
            Srvtools.FieldAttr fieldAttr6 = new Srvtools.FieldAttr();
            Srvtools.KeyItem keyItem2 = new Srvtools.KeyItem();
            this.serviceManager = new Srvtools.ServiceManager(this.components);
            this.InfoConnection1 = new Srvtools.InfoConnection(this.components);
            this.C_company = new Srvtools.InfoCommand(this.components);
            this.ucC_company = new Srvtools.UpdateComponent(this.components);
            this.View_C_company = new Srvtools.InfoCommand(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.InfoConnection1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.C_company)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.View_C_company)).BeginInit();
            // 
            // InfoConnection1
            // 
            this.InfoConnection1.EEPAlias = "EIP";
            // 
            // C_company
            // 
            this.C_company.CommandText = "SELECT [C_company].* FROM [C_company]";
            this.C_company.CommandTimeout = 30;
            this.C_company.CommandType = System.Data.CommandType.Text;
            this.C_company.DynamicTableName = false;
            this.C_company.EEPAlias = null;
            this.C_company.InfoConnection = this.InfoConnection1;
            keyItem1.KeyName = "UniqueID";
            this.C_company.KeyFields.Add(keyItem1);
            this.C_company.MultiSetWhere = false;
            this.C_company.Name = "C_company";
            this.C_company.NotificationAutoEnlist = false;
            this.C_company.SecExcept = null;
            this.C_company.SecFieldName = null;
            this.C_company.SecStyle = Srvtools.SecurityStyle.ssByNone;
            this.C_company.SelectTop = 0;
            this.C_company.SiteControl = false;
            this.C_company.SiteFieldName = null;
            this.C_company.UpdatedRowSource = System.Data.UpdateRowSource.None;
            // 
            // ucC_company
            // 
            this.ucC_company.AutoTrans = true;
            this.ucC_company.ExceptJoin = false;
            fieldAttr1.CheckNull = false;
            fieldAttr1.DataField = "UniqueID";
            fieldAttr1.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr1.DefaultValue = null;
            fieldAttr1.TrimLength = 0;
            fieldAttr1.UpdateEnable = true;
            fieldAttr1.WhereMode = true;
            fieldAttr2.CheckNull = false;
            fieldAttr2.DataField = "sort";
            fieldAttr2.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr2.DefaultValue = null;
            fieldAttr2.TrimLength = 0;
            fieldAttr2.UpdateEnable = true;
            fieldAttr2.WhereMode = true;
            fieldAttr3.CheckNull = false;
            fieldAttr3.DataField = "comp01";
            fieldAttr3.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr3.DefaultValue = null;
            fieldAttr3.TrimLength = 0;
            fieldAttr3.UpdateEnable = true;
            fieldAttr3.WhereMode = true;
            fieldAttr4.CheckNull = false;
            fieldAttr4.DataField = "comp02";
            fieldAttr4.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr4.DefaultValue = null;
            fieldAttr4.TrimLength = 0;
            fieldAttr4.UpdateEnable = true;
            fieldAttr4.WhereMode = true;
            fieldAttr5.CheckNull = false;
            fieldAttr5.DataField = "comp03";
            fieldAttr5.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr5.DefaultValue = null;
            fieldAttr5.TrimLength = 0;
            fieldAttr5.UpdateEnable = true;
            fieldAttr5.WhereMode = true;
            fieldAttr6.CheckNull = false;
            fieldAttr6.DataField = "comp04";
            fieldAttr6.DefaultMode = Srvtools.DefaultModeType.Insert;
            fieldAttr6.DefaultValue = null;
            fieldAttr6.TrimLength = 0;
            fieldAttr6.UpdateEnable = true;
            fieldAttr6.WhereMode = true;
            this.ucC_company.FieldAttrs.Add(fieldAttr1);
            this.ucC_company.FieldAttrs.Add(fieldAttr2);
            this.ucC_company.FieldAttrs.Add(fieldAttr3);
            this.ucC_company.FieldAttrs.Add(fieldAttr4);
            this.ucC_company.FieldAttrs.Add(fieldAttr5);
            this.ucC_company.FieldAttrs.Add(fieldAttr6);
            this.ucC_company.LogInfo = null;
            this.ucC_company.Name = null;
            this.ucC_company.RowAffectsCheck = true;
            this.ucC_company.SelectCmd = this.C_company;
            this.ucC_company.ServerModify = true;
            this.ucC_company.ServerModifyGetMax = false;
            this.ucC_company.TranscationScopeTimeOut = System.TimeSpan.Parse("00:02:00");
            this.ucC_company.TransIsolationLevel = System.Data.IsolationLevel.ReadCommitted;
            this.ucC_company.UseTranscationScope = false;
            this.ucC_company.WhereMode = Srvtools.WhereModeType.Keyfields;
            // 
            // View_C_company
            // 
            this.View_C_company.CommandText = "SELECT * FROM [C_company]";
            this.View_C_company.CommandTimeout = 30;
            this.View_C_company.CommandType = System.Data.CommandType.Text;
            this.View_C_company.DynamicTableName = false;
            this.View_C_company.EEPAlias = null;
            this.View_C_company.InfoConnection = this.InfoConnection1;
            keyItem2.KeyName = "UniqueID";
            this.View_C_company.KeyFields.Add(keyItem2);
            this.View_C_company.MultiSetWhere = false;
            this.View_C_company.Name = "View_C_company";
            this.View_C_company.NotificationAutoEnlist = false;
            this.View_C_company.SecExcept = null;
            this.View_C_company.SecFieldName = null;
            this.View_C_company.SecStyle = Srvtools.SecurityStyle.ssByNone;
            this.View_C_company.SelectTop = 0;
            this.View_C_company.SiteControl = false;
            this.View_C_company.SiteFieldName = null;
            this.View_C_company.UpdatedRowSource = System.Data.UpdateRowSource.None;
            ((System.ComponentModel.ISupportInitialize)(this.InfoConnection1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.C_company)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.View_C_company)).EndInit();

        }

        #endregion
    }
}
