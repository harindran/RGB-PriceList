using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PriceList
{
    class clsTable
    {
        Dictionary<string, string> keyvaltbl = new Dictionary<string, string>();
        public void FieldCreation()
        {


            AddFields("@PRICELIST", "docdat", "Doc Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@PRICELIST", "Ccode", "Cus Code", SAPbobsCOM.BoFieldTypes.db_Alpha,50);
            AddFields("@PRICELIST", "CNam", "Cus Name", SAPbobsCOM.BoFieldTypes.db_Alpha,200);
            AddFields("@PRICELIST", "att", "Attach", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
            AddFields("@PRICELISTR", "Ino", "Item No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@PRICELISTR", "INam", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
            AddFields("@PRICELISTR", "puni", "Pricing Unit", SAPbobsCOM.BoFieldTypes.db_Alpha,40);
            AddFields("@PRICELISTR", "Bprc", "Base Price", SAPbobsCOM.BoFieldTypes.db_Float,40, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@PRICELISTR", "Frei", "Freight", SAPbobsCOM.BoFieldTypes.db_Float, 40, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@PRICELISTR", "prc", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 40, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@PRICELISTR", "Cur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 40);
            AddFields("@PRICELISTR", "fdat", "From Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@PRICELISTR", "tdat", "To Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@PRICELISTR", "lprc", "Last Price", SAPbobsCOM.BoFieldTypes.db_Float, 40, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@PRICELISTR", "lef", "Last Eff From", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@PRICELISTR", "let", "Last Eff To", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@PRICELISTR", "uet", "Update Eff To", SAPbobsCOM.BoFieldTypes.db_Date);


        }
        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool canlog = false, bool Manageseries = false)
        {

            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;


                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;


                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);

                            oUserObjectMD.FormColumns.FormColumnDescription = sFind[i].Replace("U_", "");
                            if (sFind[i].StartsWith("U_"))
                                oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                            oUserObjectMD.FormColumns.FormColumnAlias = sFind[i];
                            oUserObjectMD.FormColumns.Add();
                            oUserObjectMD.FormColumns.SetCurrentLine(tablecount);

                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
                    {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }



                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule.objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        private void AddFieldsNew(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10,
     SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "",
     Dictionary<string, string> keyVal = null, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum linkob = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone,
     string setlinktable = null, string setLinkUDO = null)
        {

            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {

                if (!IsColumnExists(strTab, strCol))
                {
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;

                    foreach (var item in keyvaltbl)
                    {
                        oUserFieldMD1.ValidValues.Value = item.Key;
                        oUserFieldMD1.ValidValues.Description = item.Value;
                        oUserFieldMD1.ValidValues.Add();
                    }

                    if (setlinktable != null)
                    {
                        oUserFieldMD1.LinkedTable = setlinktable;

                    }
                    else if (setLinkUDO != null)
                    {
                        SAPbobsCOM.UserObjectsMD oUserObjectMD1 = (SAPbobsCOM.UserObjectsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                        oUserObjectMD1.GetByKey(setLinkUDO);
                        oUserFieldMD1.LinkedUDO = oUserObjectMD1.Code;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD1);
                        oUserObjectMD1 = null;

                        GC.Collect();

                    }
                    else if (linkob != SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone)
                    {
                        oUserFieldMD1.LinkedSystemObject = linkob;
                    }
                    int val;
                    val = oUserFieldMD1.Add();

                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule.objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
                keyvaltbl.Clear();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedSysObject = 0, SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {
                // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                // If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                // strTab = "@" + strTab
                // End If
                if (!IsColumnExists(strTab, strCol))
                {
                    // If Not oUserFieldMD1 Is Nothing Then
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    // End If
                    // oUserFieldMD1 = Nothing
                    // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;

                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    if (LinkedSysObject != 0)
                        oUserFieldMD1.LinkedSystemObject = LinkedSysObject;// SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulInvoices ;

                    string[] split_char;
                    if (Validvalues != null)
                    {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule.objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            string strSQL;
            try
            {
                if (clsModule.objaddon.HANA)
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                }
                else
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" + Table + "' AND AliasID = '" + Column + "'";
                }

                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
    }
}
