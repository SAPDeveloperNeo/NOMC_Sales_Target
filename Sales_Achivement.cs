using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sales_Target
{
    class Sales_Achivement
    {

        #region Variable

        private SAPbouiCOM.Form oForm, oForm1, oFormCFL;
        private SAPbouiCOM.Button oBtn = null;
        private SAPbouiCOM.Item oItem, oItem1, oItem2, oItem3;

        private SAPbouiCOM.Grid oGrid;




        private SAPbouiCOM.Matrix oMatrix, oMat;
        private Boolean ACTION = false;
        private SAPbobsCOM.Recordset oRecordSet, oRec1, oRec, oRecordSetforGrid;
        private int Mode;
        private int i, DelLine;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        private SAPbouiCOM.ChooseFromList oCFL = null;

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
                                    //if (Validation() == false)
                                    //{
                                    //    BubbleEvent = false;
                                    //    return false;

                                    //}


                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {

                        }

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
                        }

                        break;

                    #endregion

                    #region COMBO_SELECT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        if (pVal.BeforeAction == true)
                        {
                            try
                            {

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "cMaingrou" || pVal.ItemUID == "Item_3")
                            {
                                fillgrid();
                                oForm.Items.Item("Item_38").Click();

                            }



                        }

                        break;

                    #endregion

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {

                        }

                        break;

                    #endregion

                    #region CLICK
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "cMaingrou")
                                {
                                    Sales_Target_Setup.FillComboFromQuery();
                                }
                                else if (pVal.ItemUID == "Item_3")
                                {
                                    Sales_Target_Setup.FillYearCombo();
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
                return BubbleEvent;
            }
            catch (Exception Ex)
            {
                return false;
            }
        }
        #endregion

        #region fillgrid

        public static void fillgrid()
        {
            SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_39").Specific;
            string project = oForm.Items.Item("cMaingrou").Specific.Value.ToString();
            string year = oForm.Items.Item("Item_3").Specific.Value.ToString();

            oGrid.DataTable.Rows.Clear();

            //string query = "Select c.\"SlpName\" as \"SO Name\"," +
            //                "b.\"U_stJan\" as \"JAN\"," +
            //                "b.\"U_stFeb\" as \"FEB\"," +
            //                "b.\"U_stMar\" as \"MAR\"," +
            //                "b.\"U_stApr\" as \"APR\"," +
            //                "b.\"U_stMay\" as \"MAY\"," +
            //                "b.\"U_stJun\" as \"JUN\"," +
            //                "b.\"U_stJul\" as \"JUL\"," +
            //                "b.\"U_stAug\" as \"AUG\"," +
            //                "b.\"U_stSep\" as \"SEP\"," +
            //                "b.\"U_stNov\" as \"NOV\"," +
            //                "b.\"U_stDec\" as \"DEC\" " +
            //                "From \"@STSA\" a " +
            //                "inner join \"@TSA1\" b on a.\"DocEntry\" = b.\"DocEntry\" " +
            //                "left join OSLP c on a.\"U_sCode\" = c.\"SlpCode\" " +
            //                "Where b.\"U_stPrj\" = '" + project + "' and a.\"U_sYear\" = '" + year + "'";

            string query = "SELECT " +
                     "c.\"SlpName\" AS \"SO Name\"," +
                     "SUM(b.\"U_stJan\") AS \"JAN\"," +
                     "SUM(b.\"U_stFeb\") AS \"FEB\"," +
                     "SUM(b.\"U_stMar\") AS \"MAR\"," +
                     "SUM(b.\"U_stApr\") AS \"APR\"," +
                     "SUM(b.\"U_stMay\") AS \"MAY\"," +
                     "SUM(b.\"U_stJun\") AS \"JUN\"," +
                     "SUM(b.\"U_stJul\") AS \"JUL\"," +
                     "SUM(b.\"U_stAug\") AS \"AUG\"," +
                     "SUM(b.\"U_stSep\") AS \"SEP\"," +
                     "SUM(b.\"U_stNov\") AS \"NOV\"," +
                     "SUM(b.\"U_stDec\") AS \"DEC\" " +
                     "FROM " +
                     "\"@STSA\" a " +
                     "INNER JOIN \"@TSA1\" b ON a.\"DocEntry\" = b.\"DocEntry\" " +
                     "LEFT JOIN OSLP c ON a.\"U_sCode\" = c.\"SlpCode\" " +
                     "WHERE " +
                     "b.\"U_stPrj\" = '"+project+"' AND a.\"U_sYear\" = '"+year+"' " +
                     "GROUP BY " +
                     "c.\"SlpName\"";


            string query1 = "SELECT \"ASG Name\", " +
                            "SUM(\"U_stJan\") AS \"JAN\", " +
                            "SUM(\"U_stFeb\") AS \"FEB\", " +
                            "SUM(\"U_stMar\") AS \"MAR\", " +
                            "SUM(\"U_stApr\") AS \"APR\", " +
                            "SUM(\"U_stMay\") AS \"MAY\", " +
                            "SUM(\"U_stJun\") AS \"JUN\", " +
                            "SUM(\"U_stJul\") AS \"JUL\", " +
                            "SUM(\"U_stAug\") AS \"AUG\", " +
                            "SUM(\"U_stSep\") AS \"SEP\", " +
                            "SUM(\"U_stOct\") AS \"OCT\", " +
                            "SUM(\"U_stNov\") AS \"NOV\", " +
                            "SUM(\"U_stDec\") AS \"DEC\" " +
                            "FROM (SELECT \"U_sName\" AS \"ASG Name\", " +
                            "\"U_stJan\", " +
                            "\"U_stFeb\", " +
                            "\"U_stMar\", " +
                            "\"U_stApr\", " +
                            "\"U_stMay\", " +
                            "\"U_stJun\", " +
                            "\"U_stJul\", " +
                            "\"U_stAug\", " +
                            "\"U_stSep\", " +
                            "\"U_stOct\", " +
                            "\"U_stNov\", " +
                            "\"U_stDec\" " +
                            "FROM \"@STSA\" a " +
                            "INNER JOIN \"@TSA1\" b ON a.\"DocEntry\" = b.\"DocEntry\" " +
                            "LEFT JOIN OSLP c ON a.\"U_sCode\" = c.\"SlpName\" " +
                            "WHERE b.\"U_stPrj\" = '" + project + "' AND a.\"U_sYear\" = '" + year + "') " +
                            "GROUP BY \"ASG Name\"";






            FillGridSales("Item_39", query);
            FillGridSales("Item_40", query1);
        }

        public static void FillGridSales(string gridName, string query)
        {
            SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(gridName).Specific;

            oGrid.DataTable.Rows.Clear();
            oGrid.DataTable.ExecuteQuery(query);
        }

    }

}
    #endregion


