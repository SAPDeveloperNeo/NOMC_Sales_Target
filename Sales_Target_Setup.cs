using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sales_Target
{
    class Sales_Target_Setup
    {
        #region Variable

        private SAPbouiCOM.Form oForm, oForm1, oFormCFL;
        private SAPbouiCOM.Button oBtn = null;
        private SAPbouiCOM.Item oItem, oItem1, oItem2, oItem3;
        private SAPbouiCOM.Grid oGrid;

        private SAPbouiCOM.Matrix oMatrix, oMatrixA;
        private Boolean ACTION = false;
        private SAPbobsCOM.Recordset oRecordSet, oRec1, oRec, oRecordSetRP, oRecordSetMng;
        private int Mode;
        private int i, DelLine;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        private SAPbouiCOM.ChooseFromList oCFL = null;
        public static bool CloseFlg, BtnFlag = false, DBSetupFlg = false;
        private string MatrixName = string.Empty;
        string Query = "", MatName = "";


        #endregion

        #region MenuEvent
        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type)
        {
            bool bevent = true;
            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(FormId);

                if (Type == "Add")
                {
                    AddMatrixRow("Matrix1", "Col_0");
                    AddMatrixRow("mCustomer", "Col_0");
                    oForm.Items.Item("sEMP").Enabled = true;
                    oForm.Items.Item("Item_3").Enabled = true;
                    oForm.Items.Item("Item_17").Click();
                    return true;
                }
                else if (Type == "Find")
                {
                    oForm.Items.Item("sEMP").Enabled = true;
                    oForm.Items.Item("Item_3").Enabled = true;

                    //oForm.Items.Item("tCode").Specific.value = null;
                    //oForm.Items.Item("tRmk").Click();
                    return true;
                }
                else if (Type == "AddR")
                {
                    if (MatName == "Matrix1")
                    {
                        AddMatrixRow("Matrix1", "Col_0");
                    }
                    if (MatName == "mCustomer")
                    {
                        AddMatrixRow("mCustomer", "Col_0");
                    }

                }

                return bevent;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Menu Event : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        #endregion

        #region  Item event
        public bool Itemevent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {

            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(FormId);
                oMatrix = oForm.Items.Item("mCustomer").Specific;
                oMatrixA = oForm.Items.Item("Matrix1").Specific;
                switch (pVal.EventType)
                {
                    //cfl
                    #region ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.BeforeAction)
                        {
                            try
                            {
                                if (pVal.ItemUID == "1" && (pVal.FormMode == 3 || pVal.FormMode == 2))
                                {
                                    Mode = pVal.FormMode;
                                    if (Validation() == false)
                                    {
                                        BubbleEvent = false;
                                        return false;

                                    }
                                }

                                else if (pVal.ItemUID == "1" && (pVal.FormMode == 2 || pVal.FormMode == 3))
                                {
                                    oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string queryV = "Select \"U_stEmpCode\",\"U_stYear\",\"U_stPrj\" From \"@STSA\"";
                                    oRecordSet.DoQuery(queryV);
                                    if (!oRecordSet.EoF)
                                    {
                                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                                        {
                                            if (oForm.Items.Item("sEMP").Specific.Value.ToString() == oRecordSet.Fields.Item("U_stEmpCode").Value.ToString() &&
                                                oForm.Items.Item("Item_3").Specific.Value.ToString() == oRecordSet.Fields.Item("U_stYear").Value.ToString()

                                                )
                                            {
                                                return true;

                                            }

                                            oRecordSet.MoveNext();

                                        }

                                    }



                                }


                                //else if (pVal.ColUID == "#" && (pVal.FormMode == 3 || pVal.FormMode == 2) && pVal.ItemUID == "Matrix1")
                                //{
                                //    AddMatrixRow("Matrix1", "Col_0");
                                //}
                                //else if (pVal.ColUID == "#" && (pVal.FormMode == 3) && pVal.ItemUID == "mCustomer")
                                //{
                                //    AddMatrixRow("mCustomer", "Col_0");
                                //}


                            }

                            catch (Exception ex)
                            {
                                throw;
                            }
                        }


                        //}

                        break;

                    #endregion

                    #region LOST_FOCUS
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        if (pVal.BeforeAction)
                        {
                            try
                            {

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {
                            try
                            {
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }

                        break;

                    #endregion

                    #region COMBO_SELECT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        if (pVal.BeforeAction)
                        {
                            try
                            {
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {

                            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "sEMP")
                            {
                                oRecordSetMng = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("sEMP").Specific;
                                SAPbouiCOM.EditText oTextBox = (SAPbouiCOM.EditText)oForm.Items.Item("tName").Specific;
                                string selectedValue = oCombo.Selected.Value;
                                string queryForManager = "SELECT A.\"salesPrson\" AS \"Sales PersonID\", " +
                                               "C.\"SlpName\" AS \"Sales PersonName\", " +
                                               "IFNULL(B.\"firstName\", '') || ' ' || IFNULL(B.\"middleName\", '') || ' ' || IFNULL(B.\"lastName\", '') AS \"Manager\" " +
                                               "FROM ohem A " +
                                               "LEFT JOIN ohem B ON A.\"manager\" = B.\"empID\" " +
                                               "LEFT JOIN oslp C ON A.\"salesPrson\" = C.\"SlpCode\" " +
                                               "WHERE A.\"salesPrson\" = '" + selectedValue + "' AND A.\"salesPrson\" IS NOT NULL " +
                                               "ORDER BY A.\"salesPrson\"";
                                oRecordSetMng.DoQuery(queryForManager);
                                oTextBox.Value = oRecordSetMng.Fields.Item("Manager").Value;

                            }
                        }

                        break;

                    #endregion

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ColUID == "Col_0")
                            {
                                CFLCondition("CFL_CUS");
                            }
                            if (pVal.ColUID == "cMaingrou" && pVal.ItemUID == "Matrix1")
                            {
                                CFLCondition("CFL_PRJ");

                            }
                            if (pVal.ColUID == "stPrj" && pVal.ItemUID == "mCustomer")
                            {
                                CFLCondition("CFL_PRJ");

                            }
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = null;
                            sCFL_ID = oCFLEvento.ChooseFromListUID;
                            string val1 = null;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (pVal.ColUID == "Col_0")
                            {
                                try
                                {
                                    int i = pVal.Row;
                                    oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific.Value = oDataTable.GetValue("CardCode", 0).ToString();
                                    oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific.Value = oDataTable.GetValue("CardName", 0).ToString();
                                }
                                catch (Exception ex)
                                {
                                }
                                try
                                {


                                }
                                catch { }
                            }
                            if (pVal.ColUID == "cMaingrou")
                            {
                                try
                                {
                                    int i = pVal.Row;
                                    oMatrixA.Columns.Item("cMaingrou").Cells.Item(i).Specific.Value = oDataTable.GetValue("PrjCode", 0).ToString();
                                    //oMatrixA.Columns.Item("cMaingrou").Cells.Item(i).Specific.Value = oDataTable.GetValue("PrjName", 0).ToString();
                                }
                                catch (Exception ex)
                                {
                                }
                                try
                                {


                                }
                                catch { }
                            }
                            if (pVal.ColUID == "stPrj")
                            {
                                try
                                {
                                    int i = pVal.Row;
                                    oMatrix.Columns.Item("stPrj").Cells.Item(i).Specific.Value = oDataTable.GetValue("PrjCode", 0).ToString();
                                    //oMatrixA.Columns.Item("cMaingrou").Cells.Item(i).Specific.Value = oDataTable.GetValue("PrjName", 0).ToString();
                                }
                                catch (Exception ex)
                                {
                                }
                                try
                                {


                                }
                                catch { }
                            }


                        }


                        break;


                    #endregion

                    #region CLICK
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "Item_22" && oForm.Items.Item("Item_22").Enabled == true)
                                {
                                    OpenUserDefinedForm();
                                }



                            }

                            else
                            {

                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }

                        break;

                        #endregion
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
                                clsMain.SBO_Application.Forms.Item(FormId).Items.Item("2").Click();
                            }

                        }


                    }

                }
                return BubbleEvent;
            }
            catch (Exception Ex)
            {
                oForm.Freeze(false);
                return false;
            }
        }
        #endregion

        #region FormDataEvent
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {

            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                string psFormId = BusinessObjectInfo.FormUID;
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (BusinessObjectInfo.BeforeAction == true)
                        {
                        }
                        else
                        {
                            oForm.Items.Item("stSAEMP").Enabled = false;
                            oForm.Items.Item("Item_3").Enabled = false;

                        }

                        break;
                }
            }
            catch { }
        }
        #endregion

        #region Right Click Event
        public bool RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent, string FormId)
        {
            try
            {
                if (eventInfo.FormUID == FormId && eventInfo.BeforeAction == true)
                {
                    MatName = null;

                    oForm.EnableMenu("1294", false);//Duplicate Row
                    oForm.EnableMenu("1299", false);//Close Row
                    oForm.EnableMenu("1284", false);//Cancle
                    oForm.EnableMenu("1287", false);//Close Row
                    oForm.EnableMenu("771", false);//Cut Row
                    oForm.EnableMenu("774", false);//Delete Row
                    oForm.EnableMenu("775", false);//Row
                    oForm.EnableMenu("8802", false);//Maximize Row
                    oForm.EnableMenu("8801", false);//
                    oForm.EnableMenu("1292", false);//Add Row
                    oForm.EnableMenu("1293", false);//Delete Row
                    oForm.EnableMenu("784", false);//copy table Row

                    oForm.EnableMenu("772", false);//Copy RoweventInfo.ItemUID == "matDetail"
                    oForm.EnableMenu("773", false);//Paste Row


                    if (eventInfo.ItemUID == "Matrix1" || eventInfo.ItemUID == "mCustomer")
                    {
                        oForm.EnableMenu("784", true);//copy table Row
                        oForm.EnableMenu("772", true);//Copy Row
                        oForm.EnableMenu("773", true);//Paste Row
                        oForm.EnableMenu("1292", true);//Add Row
                        oForm.EnableMenu("1293", true);//Delete Row
                    }

                    MatName = eventInfo.ItemUID;
                    DelLine = eventInfo.Row;
                }
                return true;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Right Click Event : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }
        #endregion

        #region Validation
        public bool Validation()
        {
            try
            {

                if (string.IsNullOrEmpty(oForm.Items.Item("sEMP").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Sales Employee Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("sEMP").Click();
                    return false;

                }
                if (string.IsNullOrEmpty(oForm.Items.Item("tName").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("ASM Name Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("tName").Click();
                    return false;

                }
                if (string.IsNullOrEmpty(oForm.Items.Item("Item_3").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Year Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Item_3").Click();
                    return false;

                }
                if (oMatrixA.RowCount > 0)
                {
                    for (i = 1; i <= oMatrixA.RowCount; i++)
                    {
                        if (string.IsNullOrEmpty(oMatrixA.Columns.Item("cMaingrou").Cells.Item(i).Specific.value))
                        {
                            clsMain.SBO_Application.SetStatusBarMessage("Project is missing", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            // oForm.Items.Item("frmSalesTarget").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oMatrixA.Columns.Item("cMaingrou").Cells.Item(i).Click();
                            return false;
                        }
                    }
                }


            }
            catch (Exception ex)
            {

                clsMain.SBO_Application.SetStatusBarMessage("Validation : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;

            }
            return true;
        }
        #endregion

        #region FillComboFromQuery
        public static void FillComboFromQuery()
        {
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.ComboBox oCombo = oForm.Items.Item("cMaingrou").Specific;

                if (oCombo.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "SELECT DISTINCT \"PrjCode\", \"PrjName\" FROM OPRJ";
                    oRec.DoQuery(query);

                    if (!oRec.EoF)
                    {
                        while (!oRec.EoF)
                        {
                            string prjCode = oRec.Fields.Item("PrjCode").Value.ToString();
                            string prjName = oRec.Fields.Item("PrjName").Value.ToString();

                            oCombo.ValidValues.Add(prjCode, prjName);

                            oRec.MoveNext();
                        }
                    }
                }

                oForm.Items.Item("cMaingrou").DisplayDesc = true;
                // oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception ex)
            {
                // Handle the exception appropriately, e.g., display an error message
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region FillYearCombo
        public static void FillYearCombo()
        {
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.ComboBox oCombo = oForm.Items.Item("Item_3").Specific;

                if (oCombo.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "SELECT DISTINCT YEAR(\"F_RefDate\")AS F_RefYear, YEAR(\"F_RefDate\") AS F_RefYear1 FROM OFPR ORDER BY F_RefYear DESC";
                    oRec.DoQuery(query);

                    if (!oRec.EoF)
                    {
                        while (!oRec.EoF)
                        {
                            string Year = oRec.Fields.Item("F_RefYear").Value.ToString();
                            string Year1 = oRec.Fields.Item("F_RefYear1").Value.ToString();

                            oCombo.ValidValues.Add(Year, Year1);

                            oRec.MoveNext();
                        }
                    }
                }

                oForm.Items.Item("Item_3").DisplayDesc = true;
                //oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception ex)
            {
                // Handle the exception appropriately, e.g., display an error message
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region Fill Sales Employee


        public static void FillSalesEmp()
        {
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.ComboBox oCombo = oForm.Items.Item("sEMP").Specific;

                if (oCombo.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // string query = "Select Distinct \"salesPrson\",(\"firstName\"||' '||\"middleName\"||' '||\"lastName\") as \"Name\" from OHEM";
                    //string query = "Select IFNULL(T0.\"firstName\",'')||' '||IFNULL(T0.\"middleName\",'')||' '||IFNULL(T0.\"lastName\",'') as \"ASM Name\",T1.\"SlpName\" As \"SO Name\" from ohem T0 left JOIN OSLP T1 ON T0.\"salesPrson\" = T1.\"SlpCode\"";
                    string query = "SELECT A.\"salesPrson\" AS \"SalesPersonID\", " +
                                      "C.\"SlpName\" AS \"SalesPersonName\", " +
                                      "IFNULL(B.\"firstName\", '') || ' ' || IFNULL(B.\"middleName\", '') || ' ' || IFNULL(B.\"lastName\", '') AS \"Manager\" " +
                                      "FROM ohem A " +
                                      "LEFT JOIN ohem B ON A.\"manager\" = B.\"empID\" " +
                                      "LEFT JOIN oslp C ON A.\"salesPrson\" = C.\"SlpCode\" " +
                                      "WHERE A.\"salesPrson\" IS NOT NULL " +
                                      "ORDER BY A.\"salesPrson\"";


                    oRec.DoQuery(query);

                    //if (!oRec.EoF)
                    //{
                    //    while (!oRec.EoF)
                    //    {

                    //        string salesEmpCode = oRec.Fields.Item("SO Name").Value.ToString();
                    //        string salesEmpCodeName = oRec.Fields.Item("ASM Name").Value.ToString();

                    //        oCombo.ValidValues.Add(salesEmpCode, salesEmpCodeName);
                    //        oRec.MoveNext();
                    //    }
                    //}
                    if (!oRec.EoF)
                    {
                        while (!oRec.EoF)
                        {

                            string salesEmpCode = oRec.Fields.Item("SalesPersonID").Value.ToString();
                            string salesEmpCodeName = oRec.Fields.Item("SalesPersonName").Value.ToString();

                            oCombo.ValidValues.Add(salesEmpCode, salesEmpCodeName);
                            oRec.MoveNext();
                        }
                    }
                }

                oForm.Items.Item("sEMP").DisplayDesc = true;
                // oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region CFL
        public object CFLCondition(string CFL)
        {
            try
            {
                oCFL = oForm.ChooseFromLists.Item(CFL);
                oConds = clsMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (CFL == "CFL_CUS")
                {
                    oRecordSet.DoQuery("SELECT \"CardCode\" FROM \"OCRD\" WHERE \"CardType\"='C'AND \"validFor\" = 'Y'");
                    if (oRecordSet.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "CardCode";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                            if (i != oRecordSet.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("No Record found", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        oCond = oConds.Add();
                        oCond.Alias = "CardCode";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = null;
                        oCFL.SetConditions(oConds);
                        return true;
                    }

                }
                if (CFL == "CFL_PRJ")
                {
                    oRecordSet.DoQuery("Select \"PrjCode\" From OPRJ");
                    if (oRecordSet.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "PrjCode";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                            if (i != oRecordSet.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("No Record found", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        oCond = oConds.Add();
                        oCond.Alias = "PrjCode";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = null;
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                }
                //if (CFL == "CFL_PRJ1")
                //{
                //    oRecordSet.DoQuery("Select \"PrjCode\" From OPRJ");
                //    if (oRecordSet.RecordCount > 0)
                //    {
                //        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                //        {
                //            oCond = oConds.Add();
                //            oCond.Alias = "PrjCode";
                //            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                //            if (i != oRecordSet.RecordCount)
                //                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                //            oRecordSet.MoveNext();
                //        }
                //        oCFL.SetConditions(oConds);
                //        return true;
                //    }
                //    else
                //    {
                //        clsMain.SBO_Application.SetStatusBarMessage("No Record found", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                //        oCond = oConds.Add();
                //        oCond.Alias = "PrjCode";
                //        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //        oCond.CondVal = null;
                //        oCFL.SetConditions(oConds);
                //        return true;
                //    }
                //}
                return true;

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("CFL : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }
        #endregion

        #region OpenUserDefinedForm

        public void OpenUserDefinedForm()
        {

            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbobsCOM.Recordset oRecordSetRP = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string fileName = "frmSalesAchivement";

                clsMain.LoadFromXML(fileName);
                //clsMain.SBO_Application_MenuEvent(ref pVal, ref BubbleEvent, "")


            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region Add Matrix
        public void AddMatrixRow(string MatrixName, string ColName)
        {
            try
            {

                oForm.Freeze(true);
                oMatrix = oForm.Items.Item(MatrixName).Specific;

                if (oMatrix.RowCount == 0)
                {
                    oMatrix.AddRow();
                }
                else
                {
                    int lastRowIndex = oMatrix.RowCount;
                    string lastCellValue = oMatrix.Columns.Item("Col_0").Cells.Item(lastRowIndex).Specific.Value;

                    if (!string.IsNullOrEmpty(lastCellValue))
                    {
                        oMatrix.AddRow();
                    }
                }

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
        }

        #endregion
    }
}
