using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Diagnostics;
using Application = SAPbouiCOM.Framework.Application;
using SAPbouiCOM;
using System.Globalization;

namespace PriceList
{
    [FormAttribute("PriceList.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public static SAPbouiCOM.Form objform, ocompany;
        private bool update;
        private bool Findmode;
        private string Returnfilename = "";
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsContent, odbdsAttachment, odbdsBoqItem, odbdsBoqLabour;
        public Form1()
        {
            //try
            //{
            //    objform.AutoManaged = false;
            //}
            //catch (Exception ex)
            //{
            //    clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //}
        }

        public Form1(bool update)
        {
            try
            {
                Findmode = false;
                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                objform.Items.Item("1").Click();
                Findmode = true;
                objform.EnableMenu("1282", false);
                this.update = update;
                objform.Title = "Price List Display";
                Matrix0.Columns.Item("Citno").Editable = false;
                Matrix0.Columns.Item("citna").Editable = false;
                Matrix0.Columns.Item("Cpru").Editable = false;
                Matrix0.Columns.Item("CBpr").Editable = false;
                Matrix0.Columns.Item("Cfri").Editable = false;
                Matrix0.Columns.Item("CPrc").Editable = false;
                Matrix0.Columns.Item("Ccur").Editable = false;
                Matrix0.Columns.Item("Cfdat").Editable = false;
                Matrix0.Columns.Item("Ctdat").Editable = false;
                EditText3.Item.Enabled = false;



            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("stccod").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("stcus").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("stdno").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("stddat").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("etccod").Specific));
            this.EditText0.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText0_LostFocusAfter);
            //  this.EditText0.KeyDownBefore += new SAPbouiCOM._IEditTextEvents_KeyDownBeforeEventHandler(this.EditText0_KeyDownBefore);
            //  this.EditText0.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.EditText0_ValidateAfter);
            //  this.EditText0.GotFocusAfter += new SAPbouiCOM._IEditTextEvents_GotFocusAfterEventHandler(this.EditText0_GotFocusAfter);
            //  this.EditText0.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText0_LostFocusAfter);
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            //                            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("etcus").Specific));
            //                            this.EditText1.KeyDownBefore += new SAPbouiCOM._IEditTextEvents_KeyDownBeforeEventHandler(this.EditText1_KeyDownBefore);
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("etdno").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("etddat").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("tabco").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_11").Specific));
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.Matrix0_LostFocusAfter);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_14").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            //                        this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            //                     this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("statt").Specific));
            //                     this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("etatt").Specific));
            //                     this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_21").Specific));
            //                     this.Button4.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button4_ClickBefore);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_0").Specific));
            //                   objform.DataBrowser.BrowseBy = "etdno";
            //                     this.oActiveForm.DataBrowser.BrowseBy = "Item_16";
            //                this.Button8.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button8_ClickBefore);
            //                this.Button8.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button8_ClickAfter);
            //                this.Matrix2.PressedAfter += new SAPbouiCOM._IMatrixEvents_PressedAfterEventHandler(this.Matrix2_PressedAfter);
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_1").Specific));
            this.Matrix3.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix3_ClickAfter);
            this.Matrix3.PressedAfter += new SAPbouiCOM._IMatrixEvents_PressedAfterEventHandler(this.Matrix3_PressedAfter);
            this.Button11 = ((SAPbouiCOM.Button)(this.GetItem("btnbr").Specific));
            this.Button11.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button11_ClickBefore);
            this.Button11.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button11_ClickAfter);
            this.Button12 = ((SAPbouiCOM.Button)(this.GetItem("btndi").Specific));
            this.Button12.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button12_ClickAfter);
            this.Button13 = ((SAPbouiCOM.Button)(this.GetItem("btndel").Specific));
            this.Button13.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button13_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            //objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
            //Matrix0.AddRow(1);
         
            odbdsHeader = objform.DataSources.DBDataSources.Item("@PRICELIST");
            odbdsContent = objform.DataSources.DBDataSources.Item("@PRICELISTR");//Content
            odbdsAttachment = objform.DataSources.DBDataSources.Item("@PRICELISTA");
            EditText2.Value = "";
            try
            {
                IntialLoad();
            }
            catch (Exception Ex)
            {
               Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void IntialLoad()
        {
            clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Citno", "#");
            string getDocNum = @"Select IfNull(Max(""DocNum""),0) + 1 from ""@PRICELIST""";
            SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRsGetDocNum.DoQuery(getDocNum);
            EditText2.Value = oRsGetDocNum.Fields.Item(0).Value.ToString();
            string curdat = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyyMMdd");
            EditText3.Value = curdat;
            objform.Items.Item("etccod").Click();
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            // throw new System.NotImplementedException();
            objform = clsModule.objaddon.objapplication.Forms.GetForm("PriceList.Form1", pVal.FormTypeCount);

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        //private SAPbouiCOM.BoFormMode Mode;

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;

                SAPbouiCOM.ISBOChooseFromListEventArg CFL_0 = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = CFL_0.ChooseFromListUID;
                SAPbouiCOM.DataTable dt = CFL_0.SelectedObjects;
                EditText1.Value = dt.GetValue("CardName", 0).ToString();
                EditText0.Value = dt.GetValue("CardCode", 0).ToString();
            }
            catch (Exception Ex)
            {

            }

        }

        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (!pVal.InnerEvent) return;
            //if (pVal.ActionSuccess == false) return;
            switch (pVal.ColUID)
            {
                case "Citno":
                    Matrix0.FlushToDataSource();
                    SAPbouiCOM.ISBOChooseFromListEventArg CFL_1 = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                    SAPbouiCOM.DataTable dt = CFL_1.SelectedObjects;

                       
                        string val1 = dt.GetValue("ItemCode", 0).ToString();
                        string val2 = dt.GetValue("ItemName", 0).ToString();
                        
                           ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row).Specific).Value = val1;
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("citna").Cells.Item(pVal.Row).Specific).Value = val2;

                        
                        //((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row).Specific).Value = dt.GetValue("ItemCode", 0).ToString();
                       // ((SAPbouiCOM.EditText)Matrix0.Columns.Item("citna").Cells.Item(pVal.Row).Specific).Value = dt.GetValue("ItemName", 0).ToString();


                   
                    break;
                    
            }

            //Matrix0.FlushToDataSource();
            //Matrix0.AddRow(1,1);
            }

        private void Matrix0_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.ColUID=="citna")
            {
                try
                {
                    string itno = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row).Specific).Value;
                    string getUOM = @"SELECT ""SalUnitMsr""  from ""OITM"" where ""ItemCode"" = '" + itno + "'";
                    SAPbobsCOM.Recordset oRsGetUOM = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetUOM.DoQuery(getUOM);

                    SAPbouiCOM.EditText oEditText3 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Cpru").Cells.Item(pVal.Row).Specific;
                    oEditText3.Value = oRsGetUOM.Fields.Item(0).Value.ToString();

                }
                catch(Exception Ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }


            }
            if (pVal.ColUID == "CPrc")
            {
                SAPbouiCOM.EditText oEditText8 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Ccur").Cells.Item(pVal.Row).Specific;
                oEditText8.Value = "INR";
            }

                if (pVal.ColUID == "Citno")
            {
                    //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Citno", "#");

            }
            if (pVal.ColUID == "Cfdat")

            {
                //int selr = Matrix0.GetNextSelectedRow(0, BoOrderType.ot_SelectionOrder);
                int rcount = Matrix0.RowCount - 1;
                string inew = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row).Specific).Value;
                for (int j = rcount; j > 0; j--)
                {
                    string ino = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(j).Specific).Value;
                    string fdat = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Ctdat").Cells.Item(j).Specific).Value;
                    if (inew==ino)
                    {
                        string dat = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Cfdat").Cells.Item(pVal.Row).Specific).Value;
                        DateTime Tdat=DateTime.ParseExact(dat, "yyyyMMdd", CultureInfo.InvariantCulture);
                        //DateTime Tdat = DateTime.Parse(dat);
                        DateTime yday = Tdat.AddDays(-1);
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Ctdat").Cells.Item(j).Specific).Value = yday.ToString("yyyyMMdd");
                        break;
                    }
                    //arrino = new string[] { ino };

                }

                //for (int i = 0; i <= arrino.Length; i++)
                //{
                // if (inew == arrino[i])
                //{
                //  Application.SBO_Application.SetStatusBarMessage("Need to update", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                //}
                //}



            }
            if(pVal.ColUID=="Ctdat")
            {
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Citno", "#");
            }
            if (pVal.ColUID == "Cfri")
            {
                SAPbouiCOM.Column oColumn;
                objform.Freeze(true);
                try
                {

                    oColumn = Matrix0.Columns.Item("CPrc");
                    oColumn.Editable = true;

                    SAPbouiCOM.EditText oEditText1 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CBpr").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText oEditText2 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Cfri").Cells.Item(pVal.Row).Specific;

                    double value1 = Convert.ToDouble(oEditText1.Value);
                    double value2 = Convert.ToDouble(oEditText2.Value);
                    double value3 = value1 + value2;
                    SAPbouiCOM.EditText oEditTex3 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CPrc").Cells.Item(pVal.Row).Specific;
                    oEditTex3.Value = value3.ToString();
                    Matrix0.Columns.Item("Ccur").Cells.Item(pVal.Row).Click();
                    oColumn = Matrix0.Columns.Item("CPrc");
                    oColumn.Editable = false;
                }
                catch (Exception Ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                finally
                {
                    objform.Freeze(false);
                }
               //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Citno", "#");
            }

        }
        void AutogenDocNum()
        {

            EditText2.Value = "";
            try
            {

                string getDocNum = @"Select IfNull(Max(""DocNum""),0) + 1 from ""@PRICELIST""";
                SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRsGetDocNum.DoQuery(getDocNum);
                EditText2.Value = oRsGetDocNum.Fields.Item(0).Value.ToString();
                string curdat = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyyMMdd");
                EditText3.Value = curdat;
                objform.Items.Item("etccod").Click();

            }
            catch (Exception Ex)
            {
                Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        
        
        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
            {
            if (pVal.ItemUID == "Item_11" && pVal.CharPressed == 38)//up arrow
            {
              if(pVal.Row!=0)
                {
                    Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Double);

                }
            }
            if (pVal.ItemUID == "Item_11" && pVal.CharPressed == 40)//down arrow
            {
                if (pVal.Row != Matrix0.RowCount)
                {
                    Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Double);
                }
            }

            //throw new System.NotImplementedException();


        }



        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ColUID == "citna" && pVal.CharPressed == 9)
            {
                SAPbouiCOM.Column oColumn;
                objform.Freeze(true);
                try
                {

                    oColumn = Matrix0.Columns.Item("Clpr");
                    oColumn.Editable = true;
                    oColumn = Matrix0.Columns.Item("Clefr");
                    oColumn.Editable = true;
                    oColumn = Matrix0.Columns.Item("CUeto");
                    oColumn.Editable = true;
                    oColumn = Matrix0.Columns.Item("Cleto");
                    oColumn.Editable = true;

                    string Val1 = EditText0.Value;
                    string Val2 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row).Specific).Value;
                    string getprice = @"SELECT Top 1 IFNULL(T0.""U_prc"", 0)  from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + @"' order by T0.""DocEntry"" desc";
                    SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetDocNum.DoQuery(getprice);
                    SAPbouiCOM.EditText oEditText3 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Clpr").Cells.Item(pVal.Row).Specific;                    
                    oEditText3.Value = oRsGetDocNum.Fields.Item(0).Value.ToString();
                    

                    string getefdate = @"SELECT Top 1 IFNULL(T0.""U_fdat"",'') from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + @"' order by T0.""DocEntry"" desc";
                    SAPbobsCOM.Recordset oRsGetfdat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetfdat.DoQuery(getefdate);
                    SAPbouiCOM.EditText oEditText4 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Clefr").Cells.Item(pVal.Row).Specific;
                    string value = DateTime.Parse(oRsGetfdat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    if(value=="18991230")
                    {
                        oEditText4.Value = "";
                    }
                    else
                    oEditText4.Value = value;
                    


                    string getetdate = @"SELECT Top 1 IFNULL(T0.""U_tdat"",'') from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + @"' order by T0.""DocEntry"" desc";
                    SAPbobsCOM.Recordset oRsGettdat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGettdat.DoQuery(getetdate);

                    
                    SAPbouiCOM.EditText oEditText5 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Cleto").Cells.Item(pVal.Row).Specific;
                    string value1 = DateTime.Parse(oRsGettdat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    if (value1 == "18991230")
                    {
                        oEditText5.Value = "";
                    }
                    else
                         oEditText5.Value = value1;
                    


                    string getutdate = @"SELECT Top 1 IFNULL(T1.""UpdateDate"",'') from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + @"' order by T0.""DocEntry"" desc";
                    SAPbobsCOM.Recordset oRsGetudat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetudat.DoQuery(getutdate);
                    SAPbouiCOM.EditText oEditText6 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CUeto").Cells.Item(pVal.Row).Specific;
                    string value2 = DateTime.Parse(oRsGetudat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    if (value2 == "18991230")
                    {
                        oEditText6.Value = "";
                    }
                    else
                        oEditText6.Value = value2;

                    Matrix0.Columns.Item("Cpru").Cells.Item(pVal.Row).Click();
                    oColumn = Matrix0.Columns.Item("Clpr");
                    oColumn.Editable = false;
                    oColumn = Matrix0.Columns.Item("Clefr");
                    oColumn.Editable = false;
                    oColumn = Matrix0.Columns.Item("CUeto");
                    oColumn.Editable = false;
                    oColumn = Matrix0.Columns.Item("Cleto");
                    oColumn.Editable = false;
                    
                }
                catch (Exception Ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                finally
                {
                    objform.Freeze(false);
                }

            }
        }

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                if(EditText1.Value == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Select the customer", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                if (EditText3.Value == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter DocDate", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string itno = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(1).Specific).Value;
                if (itno=="")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Select item", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string prun = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Cpru").Cells.Item(1).Specific).Value;
                if (prun == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter Pricing Unit", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string BP = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CBpr").Cells.Item(1).Specific).Value;
                double t1 = Convert.ToDouble(BP);
                if (t1== 0.0)
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter BasePrice", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string BP1 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Cfri").Cells.Item(1).Specific).Value;
                double t2 = Convert.ToDouble(BP1);
                if (t2 == 0.0)
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter Freight", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string cur = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Ccur").Cells.Item(1).Specific).Value;
                if (cur == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter Currency", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string fdat = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Cfdat").Cells.Item(1).Specific).Value;
                if (fdat == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string tdat = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Ctdat").Cells.Item(1).Specific).Value;
                if (tdat == "")
                {
                    BubbleEvent = false;
                    Application.SBO_Application.SetStatusBarMessage("Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                string val2 = EditText0.Value;
                RemoveLastrow(Matrix0, "Citno");

                for (int j = 1; j <=Matrix0.RowCount; j++)
                {
                    string it= ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Citno").Cells.Item(j).Specific).Value;
                    string frdat=((SAPbouiCOM.EditText)Matrix0.Columns.Item("Cfdat").Cells.Item(j).Specific).Value;
                    string getdetail = @"SELECT Top 1 T0.""DocEntry"",T0.""U_fdat"",T0.""U_tdat"" from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry""=T1.""DocEntry"" where T0.""U_Ino"" = '" + it + @"' and T1.""U_Ccode"" = '" + val2 + @"' order by T0.""DocEntry"" desc";
                    SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetDocNum.DoQuery(getdetail);
                    string docen = oRsGetDocNum.Fields.Item("DocEntry").Value.ToString();
                    string fromdat = DateTime.Parse(oRsGetDocNum.Fields.Item("U_fdat").Value.ToString()).ToString("yyyyMMdd");
                    string todat = DateTime.Parse(oRsGetDocNum.Fields.Item("U_tdat").Value.ToString()).ToString("yyyyMMdd");
                    if (Convert.ToInt32(frdat)< Convert.ToInt32(fromdat))
                    {
                        Application.SBO_Application.SetStatusBarMessage("Entries added already", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                    if (Convert.ToInt32(frdat) > Convert.ToInt32(fromdat) && Convert.ToInt32(frdat) < Convert.ToInt32(todat))
                    {
                        string getset = @"Update Top 1 T0 set T0.""U_tdat""= ADD_DAYS(To_date('" + frdat+ @"'), -1)  from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1  ON T0.""DocEntry""=T1.""DocEntry"" where T0.""U_Ino"" = '" + it + @"' and T1.""U_Ccode"" = '" + val2 + @"' and T0.""DocEntry""='"+docen+"'";
                        SAPbobsCOM.Recordset oRsGetset= (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRsGetset.DoQuery(getset);
                    }

                }
                
                
            }
            //throw new System.NotImplementedException();
            RemoveLastrow(Matrix0, "Citno");
            RemoveLastrow(Matrix3, "trgtpath");
        }
        private void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
        {
            try
            {
                if (omatrix.VisualRowCount == 0)
                    return;
                if (string.IsNullOrEmpty(Columname_check.ToString()))
                    return;
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific).String == "")
                {
                    omatrix.DeleteRow(omatrix.VisualRowCount);
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void EditText1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            
        }

        private Matrix Matrix1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        
        private void Matrix1_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                Matrix1.SelectRow(pVal.Row, true, false);
                if (Matrix1.IsRowSelected(pVal.Row) == true)
                {
                    objform.Items.Item("btndi").Enabled = true;
                    objform.Items.Item("btndel").Enabled = true;
                }
            }
            catch (Exception ex)
            {

            }


        }

        private void Button0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.SetAttachMentFile(objform, odbdsHeader, Matrix1, odbdsAttachment);
                if (Matrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) == -1)
                {
                    objform.Items.Item("btndi").Enabled = false;
                    objform.Items.Item("btndel").Enabled = false;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //throw new System.NotImplementedException();

        }//Browse attach

        private SAPbouiCOM.Button Button4;

        private void Button5_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix1, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            //throw new System.NotImplementedException();

        }//Display Attach

        private SAPbouiCOM.Button Button5;

        private void Button6_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.DeleteRowAttachment(objform, Matrix1, odbdsAttachment, Matrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //throw new System.NotImplementedException();

        }//Delete Attachment

        private SAPbouiCOM.Button Button6;

        private void Matrix2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            
            //throw new System.NotImplementedException();

        }

        private void Matrix3_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "Item_1";
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

            //throw new System.NotImplementedException();

        }

        private void Matrix3_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                Matrix3.SelectRow(pVal.Row, true, false);
                if (Matrix3.IsRowSelected(pVal.Row) == true)
                {
                    objform.Items.Item("btndi").Enabled = true;
                    objform.Items.Item("btndel").Enabled = true;
                }
            }
            catch (Exception ex)
            {

            }

            //throw new System.NotImplementedException();

        }

        private void Button11_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.SetAttachMentFile(objform, odbdsHeader, Matrix3, odbdsAttachment);
                if (Matrix3.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) == -1)
                {
                    objform.Items.Item("btndi").Enabled = false;
                    objform.Items.Item("btndel").Enabled = false;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            //throw new System.NotImplementedException();

        }

        private void Button12_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix3, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            //throw new System.NotImplementedException();

        }

        private void Button13_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.DeleteRowAttachment(objform, Matrix3, odbdsAttachment, Matrix3.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //throw new System.NotImplementedException();

        }

        private void Button11_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            RemoveLastrow(Matrix3, "trgtpath");
            //throw new System.NotImplementedException();

        }

        private void Button2_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            //objform.Refresh();
            //throw new System.NotImplementedException();

        }

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) return;
                if (!pVal.InnerEvent && !Findmode)
                {
                    IntialLoad();
                }
                Findmode = false;
            }
            catch (Exception Ex)
            {
                Application.SBO_Application.SetStatusBarMessage(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }

        private void Form_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();

        }

        private void EditText0_KeyDownAfter(object sboObject, SBOItemEventArg pVal)
        {
            string code = EditText0.Value;
            string getdocno = @"SELECT Top 1 T1.""DocNum"" from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T1.""U_Ccode"" ='" + code + @"' order by T0.""DocEntry"" desc";
            SAPbobsCOM.Recordset oRsGetdocno = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRsGetdocno.DoQuery(getdocno);
            string value = oRsGetdocno.Fields.Item(0).Value.ToString();
            int count = oRsGetdocno.RecordCount;
            Findmode = false;
            if (count != 0)
            {
                objform.Freeze(true);
                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                EditText2.Value = value;
                objform.Items.Item("1").Click();
                objform.Freeze(false);
                Findmode = true;
            }
           


            //throw new System.NotImplementedException();

        }

        private void EditText0_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            
            //throw new System.NotImplementedException();

        }

        private void Form_DataLoadAfter(ref BusinessObjectInfo pVal)
        {
           if(update==true)
            {
                EditText3.Item.Enabled = false;
            }
            
            //throw new System.NotImplementedException();

        }

        private void Matrix0_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
              {

                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item("CFL_1");
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond;
                SAPbouiCOM.Conditions oEmptyConds = null;
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();
                //oConds = oCFL.GetConditions();
                oCond = oConds.Add();
                oCond.Alias = "ItmsGrpCod";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "110";
                oCFL.SetConditions(oConds);

                    }

                    catch (Exception ex)

                    {



                    }

            //throw new System.NotImplementedException();

        }

        private SAPbouiCOM.Button Button7;
        private Matrix Matrix3;
        private SAPbouiCOM.Button Button11;
        private SAPbouiCOM.Button Button12;
        private SAPbouiCOM.Button Button13;
        //private Matrix Matrix2;
    }
}