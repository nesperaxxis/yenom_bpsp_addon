using BusinessPartnerSpecialPrice.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BusinessPartnerSpecialPrice
{
    static class Program
    {

        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application oApplication = null;
        public static SAPbobsCOM.Recordset oRecSet;
        public static List<BPCatalogItem> oBPItems = new List<BPCatalogItem>();
        public static string BPCode = "";
        public static string BPName = "";
        public static string BPType = "";
        public static string ValidFrom = "";
        public static string ValidTo = "";
        public static string sErrMsg = null;
        public static int lErrCode = 0;
        public static int lRetCode = 0;
        public static bool btnInvAdjClicked = false;
        public static bool btnIsExists = false;
        public static string SelectedItemCodeSO = "";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Connect.ConnectUI();
                Connect.ConnectDI();
                UDO _udo = new UDO();
                _udo.Create();

                oApplication.StatusBar.SetText("Connected Successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(OApplication_AppEvent);
                oApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(OApplication_ItemEvent);

                Application.Run();
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
        }

        static void OApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

        }

        static void OApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && (pVal.FormTypeEx == "668"))
            {
                try
                {
                    if (!btnIsExists)
                    {
                        var form = (SAPbouiCOM.Form)oApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        var newbtn = form.Items.Add("btnBPSP", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        var oItem = form.Items.Item("2");
                        newbtn.Width = oItem.Width + 60;
                        newbtn.Left = oItem.Left + 80;
                        newbtn.Top = oItem.Top;
                        newbtn.LinkTo = "48";
                        SAPbouiCOM.Button oBtn = (SAPbouiCOM.Button)newbtn.Specific;
                        oBtn.Caption = "BP Catalog Prices";
                        btnIsExists = true;
                    }
                }
                catch (Exception ex)
                {
                    oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.FormTypeEx == "668")
            {
                btnIsExists = false;
            }
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.FormTypeEx == "139" && !String.IsNullOrEmpty(SelectedItemCodeSO) && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Matrix fmsForm = (SAPbouiCOM.Matrix)oApplication.Forms.ActiveForm.Items.Item("38").Specific;
                int selrow1 = fmsForm.GetCellFocus().rowIndex;
                SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)fmsForm.Columns.Item("1").Cells.Item(selrow1).Specific;
                _itemCodeCol.Value = SelectedItemCodeSO;
                SelectedItemCodeSO = "";
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.FormTypeEx == "668")
            {
                try
                {
                    if (btnInvAdjClicked)
                    {
                        SAPbouiCOM.Matrix s_priceMtx = (SAPbouiCOM.Matrix)oApplication.Forms.ActiveForm.Items.Item("14").Specific;

                        int totalItems = oBPItems.Count;
                        int cnt = 1; 
                        foreach(BPCatalogItem item in oBPItems)
                        {

                            oApplication.StatusBar.SetText($"Updating Special Prices {cnt} of {totalItems}, Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                            var isExists = false;
                            var _priceList = 0;
                            int i = 1;
                            
                            while(!isExists && s_priceMtx.RowCount >= i)
                            {
                                SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("1").Cells.Item(i).Specific;
                                SAPbouiCOM.EditText _itmType = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_ItemType").Cells.Item(i).Specific;

                                var checkBPCode = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                checkBPCode.DoQuery("SELECT s.Substitute, i.U_ITEM_TYPE " + 
                                                    "FROM OSCN s JOIN OITM i ON i.ItemCode = s.ItemCode " + 
                                                    "WHERE COALESCE(s.ItemCode,'') <> '' AND s.CardCode ='" + BPCode + "' " + 
                                                    "AND COALESCE(s.ItemCode,'') = '" + _itemCodeCol +  "' AND i.frozenfor = 'N' AND i.validfor = 'Y'");
                                var bpCatCode = Convert.ToString(checkBPCode.Fields.Item(0).Value);
                                var itmType = Convert.ToString(checkBPCode.Fields.Item(1).Value);


                                if (_itemCodeCol.Value?.ToString().ToLower() == item.ItemCode.ToLower() || (bpCatCode == item.BPCatalogNo && itmType == item.ItmType))
                                {
                                    SAPbouiCOM.EditText _price = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("5").Cells.Item(i).Specific;
                                    SAPbouiCOM.EditText _validTo = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_Valid_Until").Cells.Item(i).Specific;
                                    SAPbouiCOM.EditText _ref = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_QUOTE_REF").Cells.Item(i).Specific;
                                    SAPbouiCOM.EditText _dMaker = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_DeciMaker").Cells.Item(i).Specific;
                                    SAPbouiCOM.ComboBox _cbxPL = (SAPbouiCOM.ComboBox)s_priceMtx.Columns.Item("3").Cells.Item(i).Specific;
                                    try { _price.Value = $"{item.Currency} {item.Price}"; } catch { _price.Value = ""; }
                                    try { _validTo.Value = item.ValidTo; } catch { _validTo.Value = ""; }
                                    try { _ref.Value = item.ItmQRef; } catch { _ref.Value = ""; }
                                    try { _dMaker.Value = item.ItmDMaker; } catch { _dMaker.Value = ""; }
                                    try { _itmType.Value = item.ItmType; } catch { _itmType.Value = ""; }
                                    int.TryParse(_cbxPL?.Selected.Value, out _priceList);
                                    isExists = true;                                  
                                }
                                i++;
                            }

                            if (!isExists)
                            {                                
                                int rowId = s_priceMtx.RowCount;
                                SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("1").Cells.Item(rowId).Specific;
                                //SAPbouiCOM.EditText _itemNameCol = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("2").Cells.Item(rowId).Specific;
                                SAPbouiCOM.EditText _price = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("5").Cells.Item(rowId).Specific;
                                SAPbouiCOM.EditText _validTo = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_Valid_Until").Cells.Item(rowId).Specific;
                                SAPbouiCOM.EditText _ref = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_QUOTE_REF").Cells.Item(rowId).Specific;
                                SAPbouiCOM.EditText _dMaker = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_DeciMaker").Cells.Item(rowId).Specific;
                                SAPbouiCOM.EditText _itmType = (SAPbouiCOM.EditText)s_priceMtx.Columns.Item("U_ItemType").Cells.Item(rowId).Specific;
                                SAPbouiCOM.ComboBox _cbxPL = (SAPbouiCOM.ComboBox)s_priceMtx.Columns.Item("3").Cells.Item(rowId).Specific;
                                _itemCodeCol.Value = item.ItemCode;
                                try { _price.Value = $"{item.Currency} {item.Price}"; } catch { _price.Value = ""; }
                                try { _validTo.Value = item.ValidTo; } catch { _validTo.Value = ""; }
                                try { _ref.Value = item.ItmQRef; } catch { _ref.Value = ""; }
                                try { _dMaker.Value = item.ItmDMaker; } catch { _dMaker.Value = ""; }
                                try { _itmType.Value = item.ItmType; } catch { _itmType.Value = ""; }
                                int.TryParse(_cbxPL?.Selected.Value, out _priceList);
                                

                            }

                            var oSpp = (SAPbobsCOM.SpecialPrices)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices);
                            var isExistsOSPP = oSpp.GetByKey(item.ItemCode, BPCode);
                            if (!isExistsOSPP)
                            {
                                oSpp = (SAPbobsCOM.SpecialPrices)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices);
                                oSpp.CardCode = BPCode;
                                oSpp.ItemCode = item.ItemCode;
                            }
                            oSpp.Price = double.Parse(item.Price);
                            oSpp.PriceListNum = _priceList;
                            oSpp.Currency = item.Currency;
                            oSpp.DiscountPercent = 0;
                            oSpp.UserFields.Fields.Item("U_QUOTE_REF").Value = item.ItmQRef;
                            oSpp.UserFields.Fields.Item("U_Valid_Until").Value = $"{item.ValidTo.Substring(4, 2)}/{item.ValidTo.Substring(6, 2)}/{item.ValidTo.Substring(0, 4)}";
                            oSpp.UserFields.Fields.Item("U_DeciMaker").Value = item.ItmDMaker;
                            oSpp.UserFields.Fields.Item("U_ItemType").Value = item.ItmType;
                            var isExistsSPP1 = false;
                            var _recSetItems = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            _recSetItems.DoQuery("SELECT  1 [IsExists] FROM OSCN s WHERE COALESCE(s.ItemCode,'') = '" + item.ItemCode + "' AND s.CardCode ='" + BPCode + "'");
                            isExistsSPP1 = _recSetItems.RecordCount > 0;
                            if (!isExistsSPP1)
                            {
                                oSpp.SpecialPricesDataAreas.Add();
                            }

                            oSpp.SpecialPricesDataAreas.PriceListNo = _priceList;
                            oSpp.SpecialPricesDataAreas.DateFrom = DateTime.Parse($"{item.ValidTo.Substring(4, 2)}/{item.ValidTo.Substring(6, 2)}/{item.ValidTo.Substring(0, 4)}");
                            oSpp.SpecialPricesDataAreas.SpecialPrice = double.Parse(item.Price);
                            oSpp.SpecialPricesDataAreas.PriceCurrency = item.Currency;
                            if (!isExistsOSPP)
                            {
                                oSpp.Add();
                            }
                            else
                            {
                                oSpp.Update();
                            }
                            cnt++;
                        }
                        var form = oApplication.Forms.ActiveForm;
                        form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        oBPItems = new List<BPCatalogItem>();
                        oApplication.StatusBar.SetText($"Entered Special Prices successfully saved...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
                catch (Exception ex)
                {
                    //oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && (new string[] { "668" }).Contains(pVal.FormTypeEx))
            {
                try
                {

                    if (btnInvAdjClicked)
                    {
                        btnInvAdjClicked = false;
                        btnIsExists = false;
                    }
                }
                catch (Exception ex)
                {
                    oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }

            if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) && (new string[] { "668"}).Contains(pVal.FormTypeEx))
            {
                try
                {
                    SAPbouiCOM.DBDataSource dbDataSource;
                    switch (pVal.FormTypeEx)
                    {
                        case "668":
                            dbDataSource = oApplication.Forms.ActiveForm.DataSources.DBDataSources.Item("OSPP");
                            BPCode = dbDataSource.GetValue("CardCode", 0).ToString();
                            BPName = dbDataSource.GetValue("CardName", 0).ToString();
                            BPType = dbDataSource.GetValue("CardType", 0).ToString();
                            break;                       
                    }

                }
                catch (Exception ex)
                {

                }

            }
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormTypeEx == "2000015" && (pVal.ItemUID == "1" || pVal.ItemUID == "4") && BubbleEvent == true)
            {
                try
                {
                    SAPbouiCOM.Matrix fmsForm = (SAPbouiCOM.Matrix)oApplication.Forms.ActiveForm.Items.Item("4").Specific;
                    int selrow1 = fmsForm.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)fmsForm.Columns.Item("COL5").Cells.Item(selrow1).Specific;
                    SelectedItemCodeSO = _itemCodeCol.Value.ToString();
                }
                catch { SelectedItemCodeSO = ""; }                
            }
            
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormTypeEx == "UDO_FT_BPCP" && pVal.ItemUID == "1" && BubbleEvent == true)
            {
                BubbleEvent = false;
                oBPItems = new List<BPCatalogItem>();
                SAPbouiCOM.Matrix bpCatalogMtx = (SAPbouiCOM.Matrix)oApplication.Forms.ActiveForm.Items.Item("0_U_G").Specific;
                int totalItems = bpCatalogMtx.RowCount;
                for (int i = 1; i <= totalItems; i++)
                {
                    try
                    {                        
                        SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_0").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _itemNameCol = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_1").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _prce = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("C_0_3").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _validTo = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("C_0_5").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _itemType = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_3").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _currency = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_2").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _deciMaker = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_4").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText _qoutRef = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_5").Cells.Item(i).Specific;

                        if (!String.IsNullOrEmpty(_prce.Value?.ToString() ?? "") && (_prce.Value?.ToString() ?? "0") != "0.0")
                        {
                            oBPItems.Add(new BPCatalogItem()
                            {
                                ItemCode = _itemCodeCol.Value?.ToString() ?? "",
                                ItemName = _itemNameCol.Value?.ToString() ?? "",
                                Price = _prce.Value?.ToString() ?? "",
                                ValidTo = _validTo.Value?.ToString() ?? "",
                                ItmType = _itemType.Value?.ToString() ?? "",
                                Currency = _currency.Value?.ToString() ?? "",
                                ItmDMaker = _deciMaker.Value?.ToString() ?? "",
                                ItmQRef = _qoutRef.Value?.ToString() ?? ""                                
                            });
                        }
                        oApplication.StatusBar.SetText($"Saving entered prices {i} of {totalItems}, Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    }
                    catch (Exception ex)
                    {
                        oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }                   
                }
                var form = oApplication.Forms.ActiveForm;
                form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                form.Close();


            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormTypeEx == "668" && pVal.ItemUID == "btnBPSP")
            {
                try
                {
                    BubbleEvent = true;
                    //var _function = new Common();
                    //var _neededItems = new List<NeededItem>();
                    //for (int i = 1; i <= batchesMatrix.RowCount; i++)
                    //{
                    //    var _itemCodeCol = ((SAPbouiCOM.EditText)batchesMatrix.Columns.Item("1").Cells.Item(i).Specific).Value?.ToString() ?? "";
                    //    var _itemNameCol = ((SAPbouiCOM.EditText)batchesMatrix.Columns.Item("54").Cells.Item(i).Specific).Value?.ToString() ?? "";
                    //    var _itemQty = ((SAPbouiCOM.EditText)batchesMatrix.Columns.Item("234000021").Cells.Item(i).Specific).Value?.ToString() ?? "";
                    //    var _whouse = ((SAPbouiCOM.EditText)batchesMatrix.Columns.Item("3").Cells.Item(i).Specific).Value?.ToString() ?? "";
                    //    var _price = "0";

                    //    if (!String.IsNullOrEmpty(_itemCodeCol) && !String.IsNullOrEmpty(_itemNameCol) && !String.IsNullOrEmpty(_itemQty))
                    //    {
                    //        _neededItems.Add(new NeededItem()
                    //        {
                    //            ItemCode = _itemCodeCol,
                    //            ItemName = _itemNameCol,
                    //            Quantity = _itemQty,
                    //            WHouse = _whouse,
                    //            Price = _price
                    //        });
                    //    }
                    //}
                    oBPItems = new List<BPCatalogItem>();
                    SAPbouiCOM.EditText _code = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("5").Specific;
                    SAPbouiCOM.EditText _name = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("6").Specific;
                    SAPbouiCOM.EditText _type = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("10").Specific;
                    //SAPbouiCOM.EditText _validFrom = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("234000003").Specific;
                    //SAPbouiCOM.EditText _validTo = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("234000005").Specific;
                    BPCode = _code.Value?.ToString() ?? "";
                    BPName = _name.Value?.ToString() ?? "";
                    BPType = _type.Value?.ToString() ?? "";                    
                    //ValidFrom = _validFrom.Value?.ToString() ?? "";
                    //ValidTo = _validTo.Value?.ToString() ?? "";                    

                    if (btnInvAdjClicked)
                    {

                        if (String.IsNullOrEmpty(BPCode))
                        {
                            oApplication.StatusBar.SetText("Please select BP Code", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }                      

                        var form = oApplication.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BPCP", "");
                        //= oApplication.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsReceipt, "", "");
                        SAPbouiCOM.EditText _bpCode = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("20_U_E").Specific;
                        SAPbouiCOM.EditText _bpName = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("21_U_E").Specific;
                        SAPbouiCOM.EditText _bpType = (SAPbouiCOM.EditText)oApplication.Forms.ActiveForm.Items.Item("22_U_E").Specific;
                        _bpCode.Value = BPCode;
                        _bpName.Value = BPName;
                        _bpType.Value = BPType;
                       
                        var _recSetItems = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        _recSetItems.DoQuery("SELECT s.Substitute, s.ItemCode, s.U_CUST_NAME, i.ItmsGrpCod, i.U_ITEM_TYPE FROM OSCN s JOIN OITM i ON i.ItemCode = s.ItemCode WHERE COALESCE(s.ItemCode,'') <> '' AND s.CardCode ='" + BPCode + "' AND i.frozenfor = 'N' AND i.validfor = 'Y'");

                        var _bpCNoItems = new List<BPCatalogItem>();
                        while (!_recSetItems.EoF)
                        {
                            _bpCNoItems.Add(new BPCatalogItem()
                            {
                                BPCatalogNo = _recSetItems.Fields.Item("Substitute").Value.ToString(),
                                ItemCode = _recSetItems.Fields.Item("ItemCode").Value.ToString(),
                                ItemName = _recSetItems.Fields.Item("U_CUST_NAME").Value.ToString(),
                                ItmGrpCode = _recSetItems.Fields.Item("ItmsGrpCod").Value.ToString(),
                                ItmType = _recSetItems.Fields.Item("U_ITEM_TYPE").Value.ToString()
                            });
                            _recSetItems.MoveNext();
                        }

                        _bpCNoItems = _bpCNoItems.OrderBy(x => x.BPCatalogNo).ToList();

                        int cnt = 1;
                        int totalItems = _bpCNoItems.Count();
                        SAPbouiCOM.Matrix bpCatalogMtx = (SAPbouiCOM.Matrix)oApplication.Forms.ActiveForm.Items.Item("0_U_G").Specific;
                        List<GridItemsInfo> tempItems = new List<GridItemsInfo>();
                        foreach (BPCatalogItem _item in _bpCNoItems)
                        {                              
                            try
                            {
                                
                                if (cnt > 1)
                                {
                                    bpCatalogMtx.AddRow();
                                }

                                var checkItem = new GridItemsInfo()
                                {
                                    code = _item.BPCatalogNo,
                                    name = _item.ItemName,
                                    type = _item.ItmType
                                };

                                if (tempItems.Where( x => x.code == checkItem.code && x.type == checkItem.type).Count() > 0) 
                                {
                                    continue;
                                }
                                
                                SAPbouiCOM.EditText _rowNo = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("#").Cells.Item(cnt).Specific;
                                SAPbouiCOM.EditText _bpCNo = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("C_0_1").Cells.Item(cnt).Specific;
                                SAPbouiCOM.EditText _itemCodeCol = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_0").Cells.Item(cnt).Specific;
                                SAPbouiCOM.EditText _itemNameCol = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_1").Cells.Item(cnt).Specific;
                                SAPbouiCOM.EditText _itemTypeCol = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("Col_3").Cells.Item(cnt).Specific;
                                //SAPbouiCOM.EditText _validFromBPSP = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("C_0_4").Cells.Item(cnt).Specific;
                                //SAPbouiCOM.EditText _validToBPSP = (SAPbouiCOM.EditText)bpCatalogMtx.Columns.Item("C_0_5").Cells.Item(cnt).Specific;
                                _bpCNo.Value = _item.BPCatalogNo;
                                _itemCodeCol.Value = _item.ItemCode;
                                _itemNameCol.Value = _item.ItemName;
                                _itemTypeCol.Value = _item.ItmType;
                                _rowNo.Value = cnt.ToString();
                                //_validFromBPSP.Value = ValidFrom;
                                //_validToBPSP.Value = ValidTo;

                                tempItems.Add(checkItem);
                                oApplication.StatusBar.SetText($"Loading BP Catalog Items {cnt} of {totalItems}. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                cnt++;                                
                                
                            }
                            catch (Exception ex)
                            {
                                oApplication.StatusBar.SetText(ex.InnerException?.Message ?? ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            }
                        }
                        oApplication.StatusBar.SetText($"All items loaded successfully.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        //
                        //for (int i = 1; i <= bpCatalogMtx.RowCount; i++)
                        //{

                        //}
                        //_ref2No.Value = Ref2No;
                    }
                    btnInvAdjClicked = true;
                }
                catch (Exception ex)
                {

                }
            }

        }

        public class BPCatalog
        {
            public string BPCatalogNo { get; set; }
            public List<BPCatalogItem> Items { get; set; }
        }

        public class BPCatalogItem
        {
            public string BPCatalogNo { get; set; }
            public string ItemCode { get; set; }
            public string ItemName { get; set; }
            public string ValidFrom { get; set; }
            public string ValidTo { get; set; }
            public string Discount { get; set; }
            public string Price { get; set; }
            public string PriceUnit { get; set; }
            public string ItmGrpCode { get; set; }
            public string ItmQRef { get; set; }
            public string ItmDMaker { get; set; }
            public string ItmType { get; set; }
            public string Currency { get; set; }
        }
        public static void SetMatrixValue(SAPbouiCOM.Form oForm, string matrixUID, string newValue, string itemUID, [Optional] int rowIndex)
        {
            try
            {
                SAPbouiCOM.Matrix oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;

                oMat.SetCellWithoutValidation(rowIndex, itemUID, newValue);
            }
            catch (Exception ex)
            {
                Program.oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public class GridItemsInfo
        {
            public string code { get; set; }
            public string name { get; set; }
            public string type { get; set; }
        }
    }
}
