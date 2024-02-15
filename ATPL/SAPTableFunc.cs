using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATPL
{
    class SAPTableFunc
    {
        private bool IsColumnExists(SAPbobsCOM.Company company,string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            string strSQL;
            try
            {

                strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";

                oRecordSet = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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


        private void AddUDO(SAPbobsCOM.Company company,string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType,string strTable, out string ErrorMsg, string[] childTable =null, string[] sFind=null, bool canlog = true, bool Manageseries = true)
        {

            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            ErrorMsg = "";
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                oUserObjectMD.GetByKey(strUDO);

                if (!oUserObjectMD.GetByKey(strUDO))
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
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
                    oUserObjectMD.ExtensionName = "";

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
                        ErrorMsg= company.GetLastErrorDescription();
                    }
                }

                else
                {
                    tablecount = 0;
                    if (childTable.Length != oUserObjectMD.ChildTables.Count)
                    {
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
                                if (tablecount > 0)
                                {
                                    oUserObjectMD.Update();
                                }
                            }
                        }
                    }

                }
            }

            catch (Exception ex)
            {
                ErrorMsg = ex.ToString();
            }
            finally
            {
                if (oUserObjectMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }

        }

        private void AddFields(SAPbobsCOM.Company company, string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, out string ErrorMsg, int nEditSize = 10,
       SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "",
       string[] keyVal = null, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum linkob = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone,
       string setlinktable = null)
        {
            string[] split_char;
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            ErrorMsg = "";
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {

                if (!IsColumnExists(company,strTab, strCol))
                {
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;

                    if (keyVal != null)
                    {
                        if (keyVal.Length > 0)
                        {
                            for (int i = 0, loopTo1 = keyVal.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(keyVal[i]))
                                    continue;

                                split_char = keyVal[i].Split(Convert.ToChar(","));

                                if (split_char.Length != 2)
                                    continue;                          
                                if (string.IsNullOrEmpty(keyVal[i]))
                                    continue;

                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }

                    if (setlinktable != null)
                    {
                        oUserFieldMD1.LinkedTable = setlinktable;
                    }
                    else if (linkob != SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone)
                    {
                        oUserFieldMD1.LinkedSystemObject = linkob;
                    }
                    int val;
                    val = oUserFieldMD1.Add();

                    if (val != 0)
                    {
                        ErrorMsg = company.GetLastErrorDescription();
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
           
            }
            catch (Exception ex)
            {
                ErrorMsg= ex.ToString();
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
    }
}
