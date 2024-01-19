using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATPL
{
    public  class SapFunc
    {
        public void AddMatrixcol(SAPbouiCOM.Matrix matrix, SAPbouiCOM.Form form, string UniqName,  string Title,
                                  string Table, string TableCol,
                                  int width = 50, bool Edit = true, SAPbouiCOM.BoFormItemTypes Types=SAPbouiCOM.BoFormItemTypes.it_EDIT,SAPbouiCOM.BoLinkedObject link = SAPbouiCOM.BoLinkedObject.lf_None)
        {
            SAPbouiCOM.Columns oColumns = matrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Add(UniqName, Types);
            oColumn.TitleObject.Caption = Title;
            oColumn.Width = width;
            oColumn.Editable = Edit;

            if (link != SAPbouiCOM.BoLinkedObject.lf_None)
            {
                SAPbouiCOM.LinkedButton oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                oLinkButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Employee;
            }

            if (!string.IsNullOrEmpty(Table))
            {
                SAPbouiCOM.DBDataSource dbEdit = form.DataSources.DBDataSources.Add(Table);
                oColumn.DataBind.SetBound(true, Table, TableCol);
            }

        }

        public void EditMatrixcol(SAPbouiCOM.Matrix matrix, SAPbouiCOM.Form form, string UniqName, string Title,
                              string Table, string TableCol,
                              int width = 50, bool Edit = true, SAPbouiCOM.BoFormItemTypes Types = SAPbouiCOM.BoFormItemTypes.it_EDIT, SAPbouiCOM.BoLinkedObject link = SAPbouiCOM.BoLinkedObject.lf_None)
        {
            SAPbouiCOM.Columns oColumns = matrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Item(UniqName);
            oColumn.TitleObject.Caption = Title;
            oColumn.Width = width;
            oColumn.Editable = Edit;

            if (link != SAPbouiCOM.BoLinkedObject.lf_None)
            {
                SAPbouiCOM.LinkedButton oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                oLinkButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Employee;
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
    }
}
