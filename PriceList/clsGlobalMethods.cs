using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using System.Globalization;

namespace PriceList
{
    class clsGlobalMethods
    {
        string strsql, BankFileName;
        SAPbobsCOM.Recordset objrs;
        


        public string GetNextDocNum_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocNum\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string GetNextDocEntry_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocEntry\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
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

        public void SetAutomanagedattribute_Editable(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
        }

        public void SetAutomanagedattribute_Visible(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
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
                // OpenFile.RestoreDirectory = True
                OpenFile.InitialDirectory = clsModule.objaddon.objcompany.AttachMentPath; // "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\"
                MyProcs = Process.GetProcessesByName("SAP Business One");

               if (MyProcs.Length >= 1)
                {
                    for (int i = 0; i <= MyProcs.Length - 1; i++)
                    {
                        string[] comname = MyProcs[i].MainWindowTitle.ToString().Split(Convert.ToChar(@"-"));
                        if (comname[0] == "")
                            continue;
                        // Open dialog only for the company where the button is clicked
                        string com = clsModule.objaddon.objcompany.CompanyName.ToUpper();
                        if (comname[0].ToString().Trim().ToUpper() == com)
                        {
                            WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                            System.Windows.Forms.DialogResult ret = OpenFile.ShowDialog(MyWindow);
                            if (ret == System.Windows.Forms.DialogResult.OK)
                                BankFileName = OpenFile.FileName;
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
                BankFileName = "";
            }
            finally
            {
                OpenFile.Dispose();
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

                if (BankFileName != "")
                    return BankFileName;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.MessageBox("File Find  Method Failed : " + ex.Message);
            }
            return "";
        }
        
        public string GetServerDate()
        {
            try
            {
                SAPbobsCOM.SBObob rsetBob = (SAPbobsCOM.SBObob)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset rsetServerDate = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rsetServerDate = rsetBob.Format_StringToDate(clsModule.objaddon.objapplication.Company.ServerDate);
                DateTime DocDate = Convert.ToDateTime(rsetServerDate.Fields.Item(0).Value);

                return DocDate.ToString("yyyyMMdd");// Convert.ToString(rsetServerDate.Fields.Item(0).Value); //Convert.ToString(rsetServerDate.Fields.Item(0).Value);//.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                return "";
            }
            finally
            {
            }
        }
       
        public void OpenAttachment(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch, int PvalRow)
        {
            try
            {
                if (PvalRow <= oMatrix.VisualRowCount & PvalRow != 0)
                {
                    int RowIndex = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1;
                    string strServerPath, strClientPath;

                    strServerPath = oDBDSAttch.GetValue("U_TrgtPath", RowIndex) + @"\" + oDBDSAttch.GetValue("U_FileName", RowIndex) + "." + oDBDSAttch.GetValue("U_FileExt", RowIndex);
                    strClientPath = oDBDSAttch.GetValue("U_SrcPath", RowIndex) + @"\" + oDBDSAttch.GetValue("U_FileName", RowIndex) + "." + oDBDSAttch.GetValue("U_FileExt", RowIndex);
                    // Open Attachment File
                    OpenFile(strServerPath, strClientPath);
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("OpenAttachment Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        public bool SetAttachMentFile(SAPbouiCOM.Form oForm, SAPbouiCOM.DBDataSource oDBDSHeader, SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch)
        {
            try
            {
                if (clsModule.objaddon.objcompany.AttachMentPath.Length <= 0)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]");
                    return false;
                }

                string strFileName = FindFile();
                if (strFileName.Equals("") == false)
                {
                    string[] FileExist = strFileName.Split(Convert.ToChar(@"\"));
                    string FileDestPath = clsModule.objaddon.objcompany.AttachMentPath + FileExist[FileExist.Length - 1];

                    if (File.Exists(FileDestPath))
                    {
                        long LngRetVal = clsModule.objaddon.objapplication.MessageBox("A file with this name already exists,would you like to replace this?  " + FileDestPath + " will be replaced.", 1, "Yes", "No");
                        if (LngRetVal != 1)
                            return false;
                    }
                    string[] fileNameExt = FileExist[FileExist.Length - 1].Split(Convert.ToChar("."));
                    string ScrPath = clsModule.objaddon.objcompany.AttachMentPath;
                    ScrPath = ScrPath.Substring(0, ScrPath.Length - 1);
                    string TrgtPath = strFileName.Substring(0, strFileName.LastIndexOf(@"\"));

                    oMatrix.AddRow();
                    oMatrix.FlushToDataSource();
                    oDBDSAttch.Offset = oDBDSAttch.Size - 1;
                    oDBDSAttch.SetValue("LineId", oDBDSAttch.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath);
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, TrgtPath);
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt[0]);
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt[1]);
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, GetServerDate());
                    oMatrix.SetLineData(oDBDSAttch.Size);
                    oMatrix.FlushToDataSource();
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                return true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Set AttachMent File Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return false;
            }
            finally
            {
            }
        }

        public void DeleteRowAttachment(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch, int SelectedRowID)
        {
            try
            {
                oDBDSAttch.RemoveRecord(SelectedRowID - 1);
                oMatrix.DeleteRow(SelectedRowID);
                oMatrix.FlushToDataSource();

                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    oMatrix.GetLineData(i);
                    oDBDSAttch.Offset = i - 1;

                    oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, Convert.ToString(i));
                    oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("TrgtPath").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Path").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("FileName").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("FileExt").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Date").Cells.Item(i).Specific).Value);
                    oMatrix.SetLineData(i);
                    oMatrix.FlushToDataSource();
                }
                // oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
                oMatrix.LoadFromDataSource();

                oForm.Items.Item("btndi").Enabled = false;
                oForm.Items.Item("btndel").Enabled = false;

                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }

        public void OpenFile(string ServerPath, string ClientPath)
        {
            try
            {
                System.Diagnostics.Process oProcess = new System.Diagnostics.Process();
                try
                {
                    oProcess.StartInfo.FileName = ServerPath;
                    oProcess.Start();
                }
                catch (Exception ex1)
                {
                    try
                    {
                        oProcess.StartInfo.FileName = ClientPath;
                        oProcess.Start();
                    }
                    catch (Exception ex2)
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("" + ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    finally
                    {
                    }
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

        public void Matrix_Addrow(SAPbouiCOM.Matrix omatrix, string colname = "", string rowno_name = "", bool Error_Needed = false)
        {
            try
            {
                bool addrow = false;

                if (omatrix.VisualRowCount == 0)
                {
                    addrow = true;
                    goto addrow;
                }
                if (string.IsNullOrEmpty(colname))
                {
                    addrow = true;
                    goto addrow;
                }
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific).String != "")
                {
                    addrow = true;
                    goto addrow;
                }

            addrow:
                ;

                if (addrow == true)
                {
                   
                    omatrix.AddRow(1);
                    omatrix.ClearRowData(omatrix.VisualRowCount);
                    if (!string.IsNullOrEmpty(rowno_name))
                        ((SAPbouiCOM.EditText)omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific).String = Convert.ToString(omatrix.VisualRowCount);
                }
                else if (Error_Needed == true)
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception ex)
            {

            }
        }

    }
}
