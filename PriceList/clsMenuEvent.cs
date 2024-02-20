using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;


namespace PriceList
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;

        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "PriceList.Form1":
                        Pricelist_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                    case "PriceList.Form2":
                        objform.EnableMenu("1282", false);
                        Pricelist_MenuEvent(ref pVal, ref BubbleEvent);
                        
                        break;
                }
            }
            catch (Exception ex)
            {

            }
        }
            private void Pricelist_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
            {
                try
                {
                    //SAPbobsCOM.Recordset objRs;
                    //SAPbouiCOM.DBDataSource DBSource;
                    SAPbouiCOM.Matrix Matrix0;
                    objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                    //DBSource = objform.DataSources.DBDataSources.Item("@AT_PCQTESEL");
                    Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("Item_11").Specific;
                    if (pval.BeforeAction == true)
                    {
                        switch (pval.MenuUID)
                        {
                            case "1284": //Cancel
                                if (clsModule.objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                                {
                                    BubbleEvent = false;
                                }

                                break;
                            case "1286":
                                {
                                    //clsModule.objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    //BubbleEvent = false;
                                    //return;
                                    break;
                                }
                            case "1293":
                                if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                                break;
                        }
                    }
                    else
                    {
                        switch (pval.MenuUID)
                        {
                            case "1281": // Find Mode                           

                                //objform.Items.Item("mprodmod").Enabled = false;
                                //objform.EnableMenu("1282", true);
                                objform.ActiveItem = "etdno";
                                break;

                            case "1282"://Add Mode          

                            clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Citno", "#");

                            try
                            {
                                string getDocNum = @"Select IfNull(Max(""DocNum""),0) + 1 from ""@PRICELIST""";
                                SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oRsGetDocNum.DoQuery(getDocNum);
                                string tex = oRsGetDocNum.Fields.Item(0).Value.ToString();
                                string curdat = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyyMMdd");
                                ((SAPbouiCOM.EditText)objform.Items.Item("etddat").Specific).Value = curdat;
                                ((SAPbouiCOM.EditText)objform.Items.Item("etdno").Specific).Value = tex;
                           
                                objform.Items.Item("etccod").Click();

                            }
                            catch (Exception Ex)
                            {
                                //Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }

                            //objform.EnableMenu("1282", false);
                                break;
                            case "1292"://Add Row
                                        // if (((SAPbouiCOM.Folder)objform.Items.Item("fprodmod").Specific).Selected == true) 
                              //clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mprodmod").Specific, "modcode", "#");
                                break;

                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }
        
    }
}
