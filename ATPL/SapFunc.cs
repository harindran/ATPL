using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATPL
{
    
    public  class SapFunc
    {
        private STDFunc STDFunc = new STDFunc();

        public void AddMatrixcol(SAPbouiCOM.Matrix matrix, SAPbouiCOM.Form form, string UniqName,  string Title,
                                  string Table, string TableCol,
                                  int width = 50,
                                  bool Edit = true, SAPbouiCOM.BoFormItemTypes Types=SAPbouiCOM.BoFormItemTypes.it_EDIT,
                                  SAPbouiCOM.BoLinkedObject link = SAPbouiCOM.BoLinkedObject.lf_None,
                                  SAPbouiCOM.ChooseFromList cfl=null,
                                    bool visible=true)
        {
            SAPbouiCOM.Columns oColumns = matrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Add(UniqName, Types);
            oColumn.TitleObject.Caption = Title;
            oColumn.Width = width;
            oColumn.Editable = Edit;
            oColumn.Visible = visible;
            if (link != SAPbouiCOM.BoLinkedObject.lf_None)
            {
                SAPbouiCOM.LinkedButton oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                oLinkButton.LinkedObject =link;
            }
          
            if (!string.IsNullOrEmpty(Table))
            {
                SAPbouiCOM.DBDataSource dbEdit = form.DataSources.DBDataSources.Add(Table);
                oColumn.DataBind.SetBound(true, Table, TableCol);

                if (cfl != null)
                {
                    oColumn.ChooseFromListUID = cfl.UniqueID;             
                }

            }


        }

       

        public void EditMatrixcol(SAPbouiCOM.Matrix matrix, SAPbouiCOM.Form form, string UniqName, string Title,
                              string Table, string TableCol,
                              int width = 50, bool Edit = true, 
                              SAPbouiCOM.BoFormItemTypes Types = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                              SAPbouiCOM.BoLinkedObject link = SAPbouiCOM.BoLinkedObject.lf_None,
                              bool visible=true)
        {
            SAPbouiCOM.Columns oColumns = matrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Item(UniqName);
            oColumn.TitleObject.Caption = Title;
            oColumn.Width = width;
            oColumn.Editable = Edit;
            oColumn.Visible = visible;
            if (link != SAPbouiCOM.BoLinkedObject.lf_None)
            {
                SAPbouiCOM.LinkedButton oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                oLinkButton.LinkedObject = link;
            }

            if (!string.IsNullOrEmpty(Table))
            {
                SAPbouiCOM.DBDataSource dbEdit = form.DataSources.DBDataSources.Add(Table);
                oColumn.DataBind.SetBound(true, Table, TableCol);
            }

        }

        public SAPbobsCOM.Recordset GetmultipleRS(SAPbobsCOM.Recordset rset, string StrSQL)
        {
            try
            {
                rset.DoQuery(StrSQL);
                return rset;
            }
            catch (Exception ex)
            {
                return rset;
            }
        }

        public string GetSingleValue(SAPbobsCOM.Recordset rset, string StrSQL)
        {
            try
            {         
                  rset.DoQuery(StrSQL);
                return STDFunc.ObjtoStr((rset.RecordCount) > 0 ? rset.Fields.Item(0).Value.ToString() : "");
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public void Load_Combo(SAPbobsCOM.Recordset rset,string FormUID, SAPbouiCOM.ComboBox comboBox, string Query, string[] Validvalues = null)
        {          
                string[] split_char;
                if (comboBox.ValidValues.Count != 0) return;
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

                            comboBox.ValidValues.Add(split_char[0], split_char[1]);
                        }

                    }
                }
            if (!string.IsNullOrEmpty(Query))
            {
                rset.DoQuery(Query);
                if (rset.RecordCount == 0) return;
                for (int i = 0; i < rset.RecordCount; i++)
                {
                    comboBox.ValidValues.Add(rset.Fields.Item(0).Value.ToString(), rset.Fields.Item(1).Value.ToString());
                    rset.MoveNext();
                }
            }

                if (Validvalues != null)
                {
                    comboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }                                  
        }
        public void Load_ComboMatrix(SAPbobsCOM.Recordset rset, string FormUID, SAPbouiCOM.Column comboBox, string Query, string[] Validvalues = null)
        {
            string[] split_char;
            if (comboBox.ValidValues.Count != 0) return;
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

                        comboBox.ValidValues.Add(split_char[0], split_char[1]);
                    }

                }
            }
            if (!string.IsNullOrEmpty(Query))
            {
                rset.DoQuery(Query);
                if (rset.RecordCount == 0) return;
                for (int i = 0; i < rset.RecordCount; i++)
                {
                    comboBox.ValidValues.Add(rset.Fields.Item(0).Value.ToString(), rset.Fields.Item(1).Value.ToString());
                    rset.MoveNext();
                }
            }
            
        }
        public bool CheckFormOpen(SAPbouiCOM.Application Forms, string FormID)
        {
            bool FormExistRet = false;
            try
            {
                FormExistRet = false;
                foreach (SAPbouiCOM.Form uid in Forms.Forms)
                {
                    if (uid.TypeEx == FormID)
                    {
                        FormExistRet = true;
                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                return FormExistRet ;
            }

            return FormExistRet;

        }

    }
}
