using BusinessPartnerSpecialPrice;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BusinessPartnerSpecialPrice.Classes
{
    public class UDO
    {
        public void Create()
        {
            //BP Special Price Form Header
            CreateUDT("BPCP", "BP Catalog Price ", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateUDF("@BPCP", "BPCode", "BPCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP", "BPName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP", "BPType", "BP Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);

            //BP Special Price Form Lines
            CreateUDT("BPCP_LINES", "Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            CreateUDF("@BPCP_LINES", "GridRowNo", "Row Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "BPCCode", "BP Catalog Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "Discount", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 100);
            CreateUDF("@BPCP_LINES", "ItmPrce", "Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100);
            CreateUDF("@BPCP_LINES", "ValidFrom", "Valid From", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "ValidTo", "Effective Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "QoutRef", "Quotation Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "DeciMaker", "Decision Maker", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "ItemType", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDF("@BPCP_LINES", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
            CreateUDO_SPBP();
        }

        void CreateUDF(string udtName, string udfName, string udfDesc, SAPbobsCOM.BoFieldTypes udfType, SAPbobsCOM.BoFldSubTypes udfSubType, int udfEditSize, [Optional] string udfLinkTable, [Optional] IList<ValidValues> validvalues, [Optional] string defaultval)
        {
            SAPbobsCOM.UserFieldsMD oUDFMD = null;
            try
            {
                oUDFMD = (SAPbobsCOM.UserFieldsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                oUDFMD.TableName = udtName;
                oUDFMD.Name = udfName;
                oUDFMD.Description = udfDesc;
                oUDFMD.Type = udfType;
                oUDFMD.SubType = udfSubType;
                oUDFMD.EditSize = udfEditSize;
                oUDFMD.LinkedTable = udfLinkTable;

                if (validvalues != null)
                {
                    foreach (var item in validvalues)
                    {

                        oUDFMD.ValidValues.Add();
                        oUDFMD.ValidValues.Value = item.Value;
                        oUDFMD.ValidValues.Description = item.Description;
                    }
                }
                oUDFMD.DefaultValue = defaultval;

                Program.lRetCode = oUDFMD.Add();

                //string a = Program.oCompany.GetLastErrorDescription();

                GC.Collect();

                Program.oApplication.StatusBar.SetText(String.Format("UDF {0} Creation in UDT {1}", udfName, udtName), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Program.oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oUDFMD != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oUDFMD);
                    oUDFMD = null;
                    GC.Collect();
                }
            }
        }

        void CreateUDT(string udtName, string udtDesc, SAPbobsCOM.BoUTBTableType udtType)
        {
            SAPbobsCOM.UserTablesMD oUDTMD = null;
            try
            {
                oUDTMD = (SAPbobsCOM.UserTablesMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                oUDTMD.TableName = udtName;
                oUDTMD.TableDescription = udtDesc;
                oUDTMD.TableType = udtType;

                Program.lRetCode = oUDTMD.Add();

                Program.oApplication.StatusBar.SetText(String.Format("UDT {0} Creation", udtName), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Program.oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oUDTMD != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oUDTMD);
                    oUDTMD = null;
                    GC.Collect();
                }
            }
        }

        void CreateUDO(string udoCode, string udoName, SAPbobsCOM.BoUDOObjType udoType, string mainTable, IList<string> childTableList, int fatherMenuID, string menuID, string menuCaption, int position, [Optional] IList<string> enchancedFormColumnList, [Optional] IList<string> findColumnList, [Optional] SAPbobsCOM.BoYesNoEnum canDefForm, SAPbobsCOM.BoYesNoEnum enhancedForm, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.BoYesNoEnum haveMenuItem, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canFind, SAPbobsCOM.BoYesNoEnum canNewForm)
        {
            SAPbobsCOM.IUserObjectsMD oUDOMD = null;

            try
            {
                oUDOMD = (SAPbobsCOM.IUserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                oUDOMD.Code = udoCode;
                oUDOMD.Name = udoName;
                oUDOMD.ObjectType = udoType;
                oUDOMD.TableName = mainTable;

                bool firstLine = true;
                foreach (var childTable in childTableList)
                {
                    if (firstLine)
                    {
                        firstLine = false;
                    }
                    else
                    {
                        oUDOMD.ChildTables.Add();
                    }
                    oUDOMD.ChildTables.TableName = childTable;
                }

                oUDOMD.CanCreateDefaultForm = canDefForm;
                oUDOMD.FatherMenuID = fatherMenuID;
                oUDOMD.MenuItem = haveMenuItem;
                oUDOMD.MenuCaption = menuCaption;
                oUDOMD.Position = position;
                oUDOMD.EnableEnhancedForm = enhancedForm;

                if (enchancedFormColumnList != null)
                {
                    firstLine = true;

                    SAPbobsCOM.IUserObjectMD_EnhancedFormColumns formColumns = oUDOMD.EnhancedFormColumns;

                    foreach (var enhanceFormColumn in enchancedFormColumnList)
                    {
                        if (firstLine)
                        {
                            firstLine = false;
                        }
                        else
                        {
                            formColumns.Add();
                        }
                        formColumns.ColumnAlias = enhanceFormColumn;
                        formColumns.ColumnDescription = enhanceFormColumn;
                    }

                }

                if (canFind == SAPbobsCOM.BoYesNoEnum.tYES && findColumnList != null && findColumnList.Count > 0 && firstLine == true)
                {
                    SAPbobsCOM.IUserObjectMD_FindColumns findColumns = oUDOMD.FindColumns;

                    foreach (var findColumn in findColumnList)
                    {
                        if (firstLine)
                        {
                            firstLine = false;
                        }
                        else
                        {
                            findColumns.Add();
                        }

                        findColumns.ColumnAlias = findColumn;
                        findColumns.ColumnDescription = findColumn;
                    }
                }

                oUDOMD.ManageSeries = manageSeries;

                oUDOMD.CanDelete = canDelete;
                oUDOMD.CanCancel = canCancel;
                oUDOMD.CanClose = canClose;
                oUDOMD.CanFind = canFind;
                oUDOMD.CanCreateDefaultForm = canNewForm;

                Program.lRetCode = oUDOMD.Add();

                if (Program.lRetCode != 0)
                {
                    Program.oApplication.StatusBar.SetText(Program.oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    Program.oApplication.StatusBar.SetText(String.Format("UDO {0}, {1} Registration", udoCode, udoName), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

            }
            catch (Exception ex)
            {
                Program.oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oUDOMD != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oUDOMD);
                    oUDOMD = null;
                    GC.Collect();
                }
            }
        }

        void CreateUDO_SPBP()
        {
            //TAS
            SAPbobsCOM.UserObjectsMD oudtMD = null;
            oudtMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            if (oudtMD.GetByKey("BPCP"))
            {
                return;
            }
            else
            {
                IList<string> enhancedFormColList = new List<string>();
                IList<string> findColumnList = new List<string>();

                CreateUDO("BPCP",
                          "BPCP",
                          SAPbobsCOM.BoUDOObjType.boud_Document,
                          "BPCP",
                           new List<string>()
                           ,
                           47619,
                         "BPCP",
                         "BPCP",
                         1,
                         enhancedFormColList,
                         findColumnList,
                         SAPbobsCOM.BoYesNoEnum.tYES,
                         SAPbobsCOM.BoYesNoEnum.tYES,
                         SAPbobsCOM.BoYesNoEnum.tYES,
                         SAPbobsCOM.BoYesNoEnum.tYES,
                         SAPbobsCOM.BoYesNoEnum.tNO,
                         SAPbobsCOM.BoYesNoEnum.tNO,
                         SAPbobsCOM.BoYesNoEnum.tNO,
                         SAPbobsCOM.BoYesNoEnum.tYES,
                         SAPbobsCOM.BoYesNoEnum.tYES);
            }

        }
    }
}
