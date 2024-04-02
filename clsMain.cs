using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sales_Target
{
    public class clsMain
    {

        #region Variables
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany = null;

        string[] supportedFormats = new string[] { "dd/MM/yy", "dd/MM/yyyy", "dd-MM-yy", "dd-MM-yyyy", "MM/dd/yyyy", "MM/dd/yy", "ddMMMyyyy", "ddMMyyyy", "yyyyMMdd" };

        string psPath = "";
        SAPbouiCOM.MenuItem oMenuItem;
        SAPbouiCOM.Menus oMenus;
        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
        private SAPbobsCOM.Recordset oRec, oRecConfig, oRec1, oRecdata, oRecordset, oRecordSet, oRecM;
        public SAPbouiCOM.EventFilters oFilters;
        public SAPbouiCOM.EventFilter oFilter;

        public SAPbouiCOM.Form oForm;


        private SAPbouiCOM.CheckBox oCheckBox;
        public static SAPbouiCOM.ComboBox oCombo;
        public SAPbouiCOM.Item oItem;
        public static string DBName, ServerName, UserName, Password, DbUsername, Dbpassword, sPath, DocEntry = null, DocNum = "", Series = "", LicenseServer, SLDServer, DbServerType;
        public static string ProjectName = "Sales Target";
        public static string ProjectCode = "SD";
        private int IntCode;
        private SAPbobsCOM.UserTablesMD oUserTablesMD = null;
        private SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
        private SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
        public static bool CloseFlg, BtnFlag = false, DBSetupFlg = false;

        public static bool Success = false;
        // ManualJE ObjmanualJE = new ManualJE();
        Sales_Target_Setup ObjSalesTS = new Sales_Target_Setup();
        Sales_Achivement ObjSalesAch = new Sales_Achivement();
        // SecondarySchemeReport ObjSecSchRep = new SecondarySchemeReport();
        #endregion

        #region Main
        public clsMain()
        {
            try
            {

                SetApplication();
                Setconnection();
                DBName = SBO_Application.Company.DatabaseName; //Database Name            
                ServerName = SBO_Application.Company.ServerName;//Server Name
                LicenseServer = oCompany.LicenseServer;
                SLDServer = oCompany.SLDServer;
                DbUsername = oCompany.DbUserName;
                Dbpassword = oCompany.DbPassword;
                UserName = oCompany.UserName;
                Password = oCompany.Password;

                UserName = SBO_Application.Company.UserName;

                sPath = System.Windows.Forms.Application.StartupPath.ToString();


                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                //SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);

                if (DataBaseCreate() == true)
                {
                    //clsMain.SBO_Application.SetStatusBarMessage("Database Created successfully ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                AddMenu();
                SBO_Application.MessageBox("Add on " + ProjectName + " Connected !", 1, "Ok", "", "");
                System.Windows.Forms.Application.Run();

            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

        }
        #endregion

        #region Connection
        public void Setconnection()
        {
            try
            {
                string sCookie;
                string sConnectionContext;

                // First initialize the Company object
                oCompany = new SAPbobsCOM.Company();
                if (oCompany.Connected == true)
                    return;

                sCookie = oCompany.GetContextCookie();
                sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
                oCompany.SetSboLoginContext(sConnectionContext);
                oCompany.Connect();
            }
            catch (Exception ex)
            {
                oCompany = SBO_Application.Company.GetDICompany();
                //SBO_Application.MessageBox(ex.ToString().ToString(), 1, "ok", "", "");s
            }
        }
        private void SetFilters()
        {
            oFilters = new SAPbouiCOM.EventFilters();

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            SBO_Application.SetFilter(oFilters);
        }
        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            try
            {
                // If there's no active application the connection will fail
                SboGuiApi.Connect(sConnectionString);
            }
            catch
            { //  Connection failed
                System.Windows.Forms.MessageBox.Show("No SAP Business One Application");
                System.Environment.Exit(0);
            }
            // get an initialized application object
            SBO_Application = SboGuiApi.GetApplication(-1);
        }

        #endregion

        #region Menu
        private void AddMenu()
        {
            try
            {
                oMenus = null;
                oMenuItem = null;
                oCreationPackage = null;
                oMenus = SBO_Application.Menus;
                oCreationPackage = ((SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oMenuItem = oMenus.Item("43520");
                if (!SBO_Application.Menus.Exists("STS"))
                {
                    AddMenu_Items(16, "", BoMenuType.mt_POPUP, "Sales Target", "ST", oMenus);
                    oMenuItem = SBO_Application.Menus.Item("ST");
                    oMenus = oMenuItem.SubMenus;
                    AddMenu_Items(0, "", BoMenuType.mt_STRING, "Sales Target Setup", "sST", oMenus);
                    AddMenu_Items(1, "", BoMenuType.mt_STRING, "Sales Achivement", "sSA", oMenus);
                }
            }
            catch { }
        }

        private void AddMenu_Items(int Position, string ImageName, SAPbouiCOM.BoMenuType Type, string MenuLabel, string MenuId, SAPbouiCOM.Menus oMenus)//, ref string ImageName,ref SAPbouiCOM.BoMenuType Type, ref string MenuLabel,ref string MenuId, ref SAPbouiCOM.Menus oMenus
        {
            try
            {
                oCreationPackage.Type = Type;
                oCreationPackage.String = MenuLabel;
                oCreationPackage.UniqueID = MenuId;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = Position;
                if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
                {
                    oCreationPackage.Image = sPath + "\\" + ImageName;
                }
                if (!oMenus.Exists(MenuId))
                {
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void RemoveMenu()
        {
            oMenus = SBO_Application.Menus;
            oMenuItem = SBO_Application.Menus.Item("frmDB");
            RemoveMenu_Item(oMenus, "frmDB");

            oMenuItem = SBO_Application.Menus.Item("43520");
            RemoveMenu_Item(oMenus, "frmProject");

            oMenus = oMenuItem.SubMenus;
            oMenuItem = SBO_Application.Menus.Item("frmProject");

        }

        private void RemoveMenu_Item(SAPbouiCOM.Menus oMenus, string MenuId)
        {
            if (oMenus.Exists(MenuId) == true)
            {
                oMenuItem = SBO_Application.Menus.Item(MenuId);
                oMenus.Remove(oMenuItem);
            }
        }
        #endregion

        #region ItemEvents
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                switch (FormUID)
                {
                    case "frmSalesTarget":
                        ObjSalesTS.Itemevent(ref pVal, ref BubbleEvent, FormUID);
                        break;
                    case "frmSalesAchivement":
                        ObjSalesAch.Itemevent(ref pVal, ref BubbleEvent, FormUID);
                        break;
                }

                if (pVal.BeforeAction == true)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {

                        if (CloseFlg == true)
                        {
                            BubbleEvent = false;
                            CloseFlg = false;
                            int Msg = clsMain.SBO_Application.MessageBox("Do you want to Save Changes ?", 1, "Yes", "No", "Cancel");
                            if (Msg == 1 || Msg == 2)
                            {
                                clsMain.SBO_Application.Forms.Item(FormUID).Items.Item("2").Click();
                            }

                        }



                    }

                }


            }
            catch (Exception ex)
            {

            }

        }

        #endregion

        #region MenuEvents
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    #region Add Mode

                    if (pVal.MenuUID == "1282" || pVal.MenuUID == "1281")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;

                    }
                    #endregion
                }
                if (pVal.BeforeAction == false)
                {
                    #region Add Mode
                    if (pVal.MenuUID == "1282")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmSalesTarget":
                                ObjSalesTS.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                            case "frmSalesAchivement":
                                ObjSalesAch.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                        }
                    }
                    #endregion

                    #region Find Mode
                    if (pVal.MenuUID == "1281")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmSalesTarget":
                                ObjSalesTS.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                            case "frmSalesAchivement":
                                ObjSalesAch.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                        }
                    }
                    #endregion

                    #region Add Row
                    if (pVal.MenuUID == "1292")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmSalesTarget":
                                ObjSalesTS.MenuEvent(ref pVal, oForm.UniqueID, "AddR");
                                break;

                        }
                    }
                    #endregion

                    #region Sales Target Setup

                    if (pVal.MenuUID == "sST")
                    {
                        for (int i = 0; i < SBO_Application.Forms.Count; i++)
                        {
                            if (SBO_Application.Forms.Item(i).UniqueID == "frmSalesTargetsetup")
                            {
                                SBO_Application.Forms.Item(i).Select();
                                return;
                            }
                        }
                        LoadFromXML("frmSalesTargetsetup");
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        oForm.DataBrowser.BrowseBy = "Item_7";
                        ObjSalesTS.MenuEvent(ref pVal, oForm.UniqueID, "Add");

                        Sales_Target_Setup.FillSalesEmp();
                        Sales_Target_Setup.FillComboFromQuery();
                        Sales_Target_Setup.FillYearCombo();

                    }
                    if (pVal.MenuUID == "sSA")
                    {
                        for (int i = 0; i < SBO_Application.Forms.Count; i++)
                        {
                            if (SBO_Application.Forms.Item(i).UniqueID == "frmSalesAchivement")
                            {
                                SBO_Application.Forms.Item(i).Select();
                                return;
                            }
                        }
                        LoadFromXML("frmSalesAchivement");
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        oForm.DataBrowser.BrowseBy = "cMaingrou";
                        ObjSalesAch.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                        oForm.Items.Item("Item_38").Click();
                    }

                    #endregion

                }
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        #endregion

        #region Right Click Event
        public void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (eventInfo.FormUID)
                {
                    case "frmSalesTarget":
                        ObjSalesTS.RightClickEvent(ref eventInfo, ref BubbleEvent, eventInfo.FormUID);
                        break;

                }

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        #endregion

        #region FormDataEvent
        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "frmSalesTarget":
                        ObjSalesTS.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
                        break;

                }
            }
            catch { }
        }
        #endregion

        #region SAPEvents
        public static void LoadFromXML(string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            // load the content of the XML File
            string sPath = null;
            FileName = FileName + ".xml";
            sPath = System.Windows.Forms.Application.StartupPath.ToString();
            oXmlDoc.Load(sPath + @"\XML\" + FileName);
            // load the form to the SBO application in one batch
            string tmpStr;
            tmpStr = oXmlDoc.InnerXml;
            SBO_Application.LoadBatchActions(ref tmpStr);
            sPath = SBO_Application.GetLastBatchResults();
        }

        public static int GetNextDocNum(ref SAPbouiCOM.EditText oEdit, ref string TableName)
        {
            try
            {
                oEdit.Value = string.Empty;
                if (TableName.Trim() != string.Empty)
                {
                    SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec.DoQuery("Select Max(\"DocEntry\") From \"" + TableName + "\"");
                    int MaxCode = Convert.ToInt32(oRec.Fields.Item(0).Value);
                    if (MaxCode == 0)
                    {
                        oEdit.Value = "1";
                    }
                    else
                    {
                        if (!oRec.EoF)
                        {
                            oEdit.Value = Convert.ToString(MaxCode + 1);
                        }
                        else
                        {
                            oEdit.Value = "1";
                        }
                    }
                    oRec = null;
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Error on GetNextDocNum : Invalid TableName ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                return 0;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error on GetNextDocNum :" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return -1;
            }
        }
        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition || EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown || EventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
            {
                //RemoveMenu();
                UnHandleEvents();
                SBO_Application = null;
                oCompany = null;
                foreach (var process in Process.GetProcessesByName("WhatsApp"))
                {
                    process.Kill();
                }
                System.Windows.Forms.Application.Exit();
            }
        }
        private void UnHandleEvents()
        {
            try
            {
                SBO_Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.AppEvent -= new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.FormDataEvent -= new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                SBO_Application.RightClickEvent -= new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
            }
            catch (Exception ex)
            {
            }
        }
        #endregion

        #region Other Methods

        public static void SetCode(String FormId, String OBJ, String Type)
        {
            try
            {
                SAPbobsCOM.Recordset oRecordSet;
                SAPbouiCOM.Form oForm;
                SAPbouiCOM.EditText oEdit;
                string Table = null;
                DateTime now = DateTime.Now;
                oForm = SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                Table = "@" + OBJ;

                GetNextDocNum(ref oEdit, ref Table);
                //D means Document of UDO we created.
                if (Type == "D")
                {
                    oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("Select * From OUDO Where \"MngSeries\" = 'N' and \"Code\" = '" + OBJ + "' ");
                    oForm.Items.Item("tDocDate").Specific.value = DateTime.Now.ToString("yyyyMMdd");
                    if (oRecordSet.RecordCount > 0)
                    {
                        oForm.Items.Item("tDocNum").Specific.value = oEdit.Value;
                    }
                    else
                    {

                        oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.value, OBJ);
                    }

                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Set Code :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Create DataBase

        private bool DataBaseCreate()
        {
            try
            {
                //AddMenu();
                oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery("Select * From \"@LICENSEMST\" Where \"Name\" = '" + clsMain.ProjectCode + "'");
                if (oRecordset.RecordCount == 0)
                {
                    DBSetupFlg = false;
                }
                else if (oRecordset.Fields.Item("U_DBCreate").Value == "N")
                {
                    DBSetupFlg = true;
                }
            }
            catch
            {
                DBSetupFlg = true;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (DBSetupFlg == true)
            {
                try
                {
                    clsMain.SBO_Application.SetStatusBarMessage("Database structure creation in progress...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    CreateDataBase();
                    oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery("Select * From \"@LICENSEMST\" Where \"Name\" = '" + clsMain.ProjectCode + "'");
                    if (oRecordset.RecordCount > 0)
                    {
                        oRecordset.DoQuery("Update \"@LICENSEMST\" Set \"U_DBCreate\" = 'N' Where \"Name\" = '" + clsMain.ProjectCode + "'");
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return true;
        }

        private bool CreateDataBase()
        {
            try
            {
                //#region Licence field creation

                //CreateTable("LICENSEMST", "License Master", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //CreateFields("@LICENSEMST");
                //#endregion

                #region Sales Target

                CreateTable("STSA", "Sales Target Setup", SAPbobsCOM.BoUTBTableType.bott_Document);
                CreateTable("TSA1", "SalesTarget Matrix1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                CreateTable("TSA2", "SalesTarget Matrix2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

                CreateFields("@STSA");

                string[] FindColoums = new string[2];
                FindColoums[0] = "U_stEmpCode";
                FindColoums[1] = "U_stYear";


                string[] ChildTable = new string[2];
                ChildTable[0] = "TSA1";
                ChildTable[1] = "TSA2";

                CreateUserObject("STSA", "Sales Target Setup", "STSA", ChildTable, FindColoums, SAPbobsCOM.BoUDOObjType.boud_Document, "N");
                #endregion


                //#region Secondary Scheme Grid

                //CreateTable("OSEG", "SecondarySch Grid", SAPbobsCOM.BoUTBTableType.bott_Document);
                //CreateFields("@OSEG");
                //string[] FindColoum1 = new string[0];
                //string[] ChildTables1 = new string[0];

                //CreateUserObject("OSEG", "SecondarySch Grid", "OSEG", ChildTables1, FindColoum1, SAPbobsCOM.BoUDOObjType.boud_Document, "N");
                //#endregion

                #region SystemTable

                CreateFields("SystemTable");

                #endregion

                //#region AutoFMS
                //AutoFMS();
                //#endregion

                return true;
            }
            catch (Exception ex)
            {
                // SBO_Application.StatusBar.SetText(ex.ToString().ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private bool CreateFields(string TableName)
        {
            try
            {
                #region SystemTable
                if (TableName == "SystemTable")
                {
                    FieldDetails("OINV", "RMSCHAMT", "Remaining Scheme Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");

                }
                #endregion

                #region Sales Target
                if (TableName == "@STSA")
                {
                    FieldDetails("@STSA", "sCode", "Sales Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@STSA", "sName", "ASM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                   // FieldDetails("@STSA", "stPrj", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");
                    FieldDetails("@STSA", "sYear", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");
                    FieldDetails("@STSA", "sRmk", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "");
                    //FieldDetails("@OCSH", "SInactive", "In Active", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", "N");

                    FieldDetails("@TSA1", "stPrj", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");
                    FieldDetails("@TSA1", "stJan", "January Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stFeb", "February Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stMar", "March Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stApr", "April Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stMay", "May Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stJun", "June Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stJul", "July Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stAug", "August Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stSep", "September Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stOct", "October Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stNov", "November Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");
                    FieldDetails("@TSA1", "stDec", "December Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", "");

                    FieldDetails("@TSA2", "sCCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@TSA2", "sCName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@TSA2", "stPrj", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");


                }
                #endregion

                //#region Secondary Scheme Grid
                //if (TableName == "@OSEG")
                //{
                //    FieldDetails("@OSEG", "SPoj", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");
                //    FieldDetails("@OSEG", "SYear", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");

                //}
                //#endregion


                return true;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Create Field: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        private bool CreateTable(string TableName, string TableDesc, SAPbobsCOM.BoUTBTableType TableType)
        {
            try
            {
                int errCode;
                string ErrMsg = null;
                //oUserTablesMD = null;
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                //oUserTablesMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserTablesMD == null)
                {
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                }
                if (oUserTablesMD.GetByKey(TableName) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserTablesMD.TableName = TableName;
                oUserTablesMD.TableDescription = TableDesc;
                oUserTablesMD.TableType = TableType;
                long err = oUserTablesMD.Add();
                if (err != 0)
                {
                    clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                }
                if (err == 0)
                {
                    clsMain.SBO_Application.StatusBar.SetText("Table Created : " + TableDesc + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //CreateFields(TableName);
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool FieldDetails(string TableName, string FieldName, string FieldDesc, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int FieldSize, string ValidValues, string DefaultVal)
        {
            if (FieldExist(TableName, FieldName) == false)
            {
                string ErrMsg;
                int errCode;
                int IRetCode;
                oUserFieldsMD = null;
                try
                {
                    GC.Collect();
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                    oUserFieldsMD.TableName = TableName;
                    oUserFieldsMD.Name = FieldName;
                    oUserFieldsMD.Description = FieldDesc;

                    oUserFieldsMD.Type = FieldType;
                    oUserFieldsMD.SubType = FieldSubType;
                    oUserFieldsMD.EditSize = FieldSize;

                    if (ValidValues != "")
                    {
                        switch (ValidValues)
                        {
                            case "NY":
                                oUserFieldsMD.ValidValues.Value = "Y";
                                oUserFieldsMD.ValidValues.Description = "Yes";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "N";
                                oUserFieldsMD.ValidValues.Description = "No";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "N";
                                break;

                            case "DocType":
                                oUserFieldsMD.ValidValues.Value = "13";
                                oUserFieldsMD.ValidValues.Description = "A/R Invoice";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "17";
                                oUserFieldsMD.ValidValues.Description = "Sales Order";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "23";
                                oUserFieldsMD.ValidValues.Description = "Sales Quotation";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "203";
                                oUserFieldsMD.ValidValues.Description = "A/R Downpayment Request";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "-1";
                                oUserFieldsMD.ValidValues.Description = "WhatsApp Bulk";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "13";
                                break;

                            case "Currency":
                                oUserFieldsMD.ValidValues.Value = "INR";
                                oUserFieldsMD.ValidValues.Description = "INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "NOT_INR";
                                oUserFieldsMD.ValidValues.Description = "Other Than INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "INR";
                                break;

                            case "SAct":
                                oUserFieldsMD.ValidValues.Value = "INR";
                                oUserFieldsMD.ValidValues.Description = "INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "NOT_INR";
                                oUserFieldsMD.ValidValues.Description = "Other Than INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "INR";
                                break;

                        }
                    }
                    if (DefaultVal != "")
                    {
                        oUserFieldsMD.DefaultValue = DefaultVal;
                    }
                    // Add the field to the table
                    IRetCode = oUserFieldsMD.Add();
                    if (IRetCode != 0)
                    {
                        if (IRetCode == -2035 || IRetCode == -1120)
                        {
                            clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            return false;
                        }
                        else
                        {
                            clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            clsMain.SBO_Application.SetStatusBarMessage("Error : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("Field Created in : " + TableName + " As : " + FieldDesc, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    ErrMsg = null; errCode = 0; IRetCode = 0;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return true;
        }

        private bool CreateUserObject(string CodeID, string Name, string TableName, string[] ChildTableName, string[] FindColoums, SAPbobsCOM.BoUDOObjType ObjectType, string ManageSeries) //used for registration of user defined table
        {
            try
            {
                int lRetCode = 0, code = 0;
                string sErrMsg = null;
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                if (oUserObjectMD.GetByKey(CodeID) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.TableName = TableName;

                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                if (ManageSeries == "Y")
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                else
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                }

                if (ChildTableName != null)
                {
                    for (int i = 0; i <= ChildTableName.Length - 1; i++)
                    {
                        if (ChildTableName[i] != null)
                        {
                            if (ChildTableName[i].Trim() != string.Empty)
                            {
                                oUserObjectMD.ChildTables.TableName = ChildTableName[i];
                                oUserObjectMD.ChildTables.Add();
                            }
                        }
                    }
                }
                if (FindColoums != null)
                {
                    for (int i = 0; i <= FindColoums.Length - 1; i++)
                    {
                        if (FindColoums[i] != null)
                        {
                            if (FindColoums[i].Trim() != string.Empty)
                            {
                                oUserObjectMD.FindColumns.ColumnAlias = FindColoums[i];
                                oUserObjectMD.FindColumns.Add();
                            }
                        }
                    }
                }
                // check for errors in the process
                lRetCode = oUserObjectMD.Add();

                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { clsMain.oCompany.GetLastError(out lRetCode, out sErrMsg); }
                    else
                    { clsMain.oCompany.GetLastError(out lRetCode, out sErrMsg); }
                else
                {
                    SAPbobsCOM.Recordset oRecordset, oRec;
                    oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    if (TableName != "LICAUTHMST")
                    {
                        oRecordset.DoQuery("Select \"Code\" from \"@OBJECTMST\" Where \"U_ObjCode\" = '" + TableName + "' ");
                        if (oRecordset.RecordCount > 0)
                        {
                            oRecordset.DoQuery("Select \"Code\" from \"@OBJECTMST\" Where \"U_Module\" like '%" + clsMain.ProjectCode + "%' And \"U_ObjCode\" = '" + TableName + "' ");
                            if (oRecordset.RecordCount == 0)
                            {
                                oRecordset.DoQuery("Update \"@OBJECTMST\" set \"U_Module\" = '" + clsMain.ProjectCode + "' + ',' + ISNULL(\"U_Module\",'') where \"U_ObjCode\" = '" + TableName + "'");
                            }
                        }
                        else
                        {
                            oRec.DoQuery("Select max(cast(\"Code\" as integer)) From \"@OBJECTMST\"");
                            if (oRec.RecordCount > 0)
                            {
                                IntCode = Convert.ToInt32(oRec.Fields.Item(0).Value) + 1;
                            }
                            else
                            {
                                IntCode = 1;
                            }
                            oRecordset.DoQuery("Insert into \"@OBJECTMST\" (\"Code\",\"Name\",\"U_ObjCode\",\"U_ObjName\",\"U_Module\") Values ('" + IntCode + "','" + IntCode + "','" + TableName + "','" + Name + "','" + clsMain.ProjectCode + "') ");
                        }
                        clsMain.SBO_Application.StatusBar.SetText("Object Registered : " + CodeID + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    oRec = null;
                    oRecordset = null;
                    IntCode = 0;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        private bool FieldExist(string TableName, string ColumnName)
        {
            SAPbobsCOM.Recordset oRecordSet = default(SAPbobsCOM.Recordset);
            oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'");
                if ((Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0))
                {
                    oRecordset = null;
                    return false;
                }
                else
                {
                    oRecordset = null;
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion
    }
}
