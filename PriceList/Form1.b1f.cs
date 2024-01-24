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

namespace PriceList
{
    [FormAttribute("PriceList.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public static SAPbouiCOM.Form objform, ocompany;
        private string Returnfilename = "";
        public Form1()
        {
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
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            //       this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("etcus").Specific));
            //       this.EditText1.KeyDownBefore += new SAPbouiCOM._IEditTextEvents_KeyDownBeforeEventHandler(this.EditText1_KeyDownBefore);
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("etdno").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("etddat").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("tabco").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_11").Specific));
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.Matrix0_LostFocusAfter);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_14").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            //   this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("statt").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("etatt").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_21").Specific));
            this.Button4.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button4_ClickBefore);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_0").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
            

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            //objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
            Matrix0.AddRow(1);
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
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        //private SAPbouiCOM.BoFormMode Mode;

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.ISBOChooseFromListEventArg CFL_0 = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
            string Uid = CFL_0.ChooseFromListUID;
            SAPbouiCOM.DataTable dt = CFL_0.SelectedObjects;
            EditText1.Value = dt.GetValue("CardName",0).ToString();
            EditText0.Value = dt.GetValue("CardCode", 0).ToString();

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
            if (pVal.ColUID == "Ctdat")

            {

                Matrix0.AddRow(1,-1);
                Matrix0.ClearRowData(Matrix0.RowCount);
                Matrix0.Columns.Item("Citno").Cells.Item(pVal.Row + 1).Click();

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

            }

        }

        public string FindFile()
        {
            System.Threading.Thread ShowFolderBrowserThread;

            try
            {
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }

                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                    // ShowFolderBrowserThread.Sleep(100)
                    Thread.Sleep(100);
                }

                if (Returnfilename != "")
                    return Returnfilename;


            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.MessageBox("File Find  Method Failed : " + ex.Message);
            }
            return "";
        }
        public void ShowFolderBrowser()
        {
            System.Diagnostics.Process[] MyProcs;
            OpenFileDialog OpenFile = new OpenFileDialog();

            try
            {
                OpenFile.Multiselect = false;
                OpenFile.Filter = "All files(*.)|*.*"; // "|*.*"
                int filterindex = 0;
                try
                {
                    filterindex = 0;
                }
                catch (Exception ex)
                {
                }
                OpenFile.FilterIndex = filterindex;
                OpenFile.InitialDirectory = clsModule.objaddon.objcompany.AttachMentPath; // "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\"
                MyProcs = Process.GetProcessesByName("SAP Business One");

                if (MyProcs.Length >= 1)
                {
                    for (int i = 0; i <= MyProcs.Length - 1; i++)
                    {
                        string[] comname = MyProcs[i].MainWindowTitle.ToString().Split(Convert.ToChar(@"-"));
                        if (comname[0] == "")
                            continue;
                        string com = clsModule.objaddon.objcompany.CompanyName.ToUpper();
                        if (comname[0].ToString().Trim().ToUpper() == com)
                        {
                            WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                            System.Windows.Forms.DialogResult ret = OpenFile.ShowDialog(MyWindow);
                            if (ret == System.Windows.Forms.DialogResult.OK)
                                Returnfilename = OpenFile.FileName;
                            else
                                System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message);
                Returnfilename = "";
            }
            finally
            {
                OpenFile.Dispose();
            }

        }
        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private IntPtr _hwnd;

            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }

            public System.IntPtr Handle
            {
                get
                {
                    return _hwnd;
                }
            }
        }
        public void OpenFile(string Path)
        {

            try
            {
                if (string.IsNullOrEmpty(Path)) return;
                System.Diagnostics.Process oProcess = new System.Diagnostics.Process();
                try
                {
                    oProcess.StartInfo.FileName = Path;
                    oProcess.Start();
                }
                catch (Exception ex1)
                {
                }
                finally
                {
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }


        

        private void Button4_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            
            BubbleEvent = true;
            string strFileName = FindFile();
            this.EditText4.Value = strFileName;
            //throw new System.NotImplementedException();

        }

        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
           
            
            }



        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ColUID == "Ccur" && pVal.CharPressed == 9)
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
                    string getprice = @"SELECT Top 1 IFNULL(T0.""U_prc"", 0)  from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + "'";
                    SAPbobsCOM.Recordset oRsGetDocNum = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetDocNum.DoQuery(getprice);

                    SAPbouiCOM.EditText oEditText3 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Clpr").Cells.Item(pVal.Row).Specific;                    
                    oEditText3.Value = oRsGetDocNum.Fields.Item(0).Value.ToString();
                    

                    string getefdate = @"SELECT Top 1 CAST(T0.""U_fdat"" AS DATE) from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + "'";
                    SAPbobsCOM.Recordset oRsGetfdat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetfdat.DoQuery(getefdate);

                  
                    SAPbouiCOM.EditText oEditText4 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Clefr").Cells.Item(pVal.Row).Specific;
                    string value = DateTime.Parse(oRsGetfdat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    oEditText4.Value = value;
                    


                    string getetdate = @"SELECT Top 1 CAST(T0.""U_tdat"" AS DATE) from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T0.""U_Ino"" = '" + Val2 + @"' and T1.""U_Ccode"" ='" + Val1 + "'";
                    SAPbobsCOM.Recordset oRsGettdat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGettdat.DoQuery(getetdate);

                    
                    SAPbouiCOM.EditText oEditText5 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Cleto").Cells.Item(pVal.Row).Specific;
                    string value1 = DateTime.Parse(oRsGettdat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    oEditText5.Value = value1;
                    


                    string getutdate = @"SELECT Top 1 CAST(T1.""UpdateDate"" AS DATE) from ""@PRICELISTR"" T0 INNER JOIN ""@PRICELIST"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where T1.""U_Ccode"" ='" + Val1 + "'";
                    SAPbobsCOM.Recordset oRsGetudat = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsGetudat.DoQuery(getutdate);

                    
                    SAPbouiCOM.EditText oEditText6 = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CUeto").Cells.Item(pVal.Row).Specific;
                    string value2 = DateTime.Parse(oRsGetudat.Fields.Item(0).Value.ToString()).ToString("yyyyMMdd");
                    oEditText6.Value = value2;

                    Matrix0.Columns.Item("Ccur").Cells.Item(pVal.Row).Click();
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



            }
            //throw new System.NotImplementedException();

        }

        private void EditText1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

            
        }

        
    }
}