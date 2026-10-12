using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.ApplicationBlocks.Data;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
//using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Linq;
using System.Text;
using System.Configuration;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
using System.Net.Mail;
using System.Threading;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Timers;

namespace DeliveryCreation
{
    class ClsDeliveryCreation : Base
    {
        #region Variable Declaration
        private SAPbouiCOM.Matrix _Matrix;
        private SAPbobsCOM.Recordset _RS = null, _RSDel = null,_RSDelNo = null;
        string SQL = string.Empty;
        private SAPbouiCOM.EditText _EditText;
        int _LastErrorCode = 0;
        string _LastErrorMsg = " ";
        public static string _PONum = string.Empty;
        public static string _PODocEntry = string.Empty;
        public static string _AfterLandingCode = string.Empty;
        SAPbouiCOM.Condition oCon;
        SAPbouiCOM.Conditions oCons;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private int RowToDelete;
        private SAPbouiCOM.CheckBox _CheckBox;
        private SAPbouiCOM.ComboBox _ComboBox;
        double _LineTotal = 0, SumLineTotal = 0, _TaxAmount = 0, SumTaxAmount = 0;
        public bool _Checked = false;
        double TaxAmount = 0, LineTotal = 0;
        string _DocNum = string.Empty;
        SAPbouiCOM.EditTextColumn _EditTextColumn;
        string _PostingDate = string.Empty,_DocumentDate = string.Empty,_DeliveryDate = string.Empty;

        string DeliveryEntry = string.Empty;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        static string ScanCode = string.Empty;
        static double Quantity = 0;
        static double CurrentQuantityBulk = 0;
        static double CurrentQuantity = 0;
        static bool _Status =false;
        private StringBuilder barcodeBuffer = new StringBuilder();
        private bool isCapturingBarcode = false;
        private System.Timers.Timer barcodeTimer;
        private const int BarcodeProcessingDelay = 200;
        private static SAPbouiCOM.EditText oScanTxt;
        private string _scanBuffer = "";
        private static SAPbouiCOM.Form oScanForm;
        #endregion

        public ClsDeliveryCreation()
        {
            // Initialize the timer
            barcodeTimer = new System.Timers.Timer(BarcodeProcessingDelay);
            barcodeTimer.Elapsed += OnBarcodeTimerElapsed;
            barcodeTimer.AutoReset = false; 
        }

        #region Item Events

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                   //case SAPbouiCOM.BoEventTypes.et_CLICK:
                   //     if (pVal.ColUID == "V_14")
                   //     {
                   //         if(Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific).Value) > Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_15").Cells.Item(pVal.Row).Specific).Value))
                   //         {
                   //             oBubbleEvent = false;
                   //             return;
                   //         }
                   //     }
                   //     break;

                    #region et_ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            int selectedRowCount = 0;
                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                string _QueryValidation = "select \"U_DocNum\",\"U_CardName\" from \"@DELCH\" where \"U_Status\" = 'Open' AND  \"U_CardCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString() + "'";
                                Utilities.ExecuteSQL(ref _RS, _QueryValidation);
                                if (_RS.RecordCount > 0)
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Document No. '" + _RS.Fields.Item("U_DocNum").Value + "' is already open for :'" + _RS.Fields.Item("U_CardName").Value + "'", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                            }
                            _Form = Utilities.Application.SBO_Application.Forms.Item(FormUID);
                            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("14").Specific;
                            _Matrix.FlushToDataSource();
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                if (((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(i).Specific).Checked == true)
                                {
                                    selectedRowCount = selectedRowCount + 1;
                                }
                                
                            }
                            if (selectedRowCount <= 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Without row selection can't Add or Update Data", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                if (((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(i).Specific).Checked == false)
                                {
                                    _Matrix.DeleteRow(i);
                                    i = 0;
                                }
                            }
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = Utilities.getMaxColumnValueNum("@DELCH", "U_DocNum");
                        }
                        if(pVal.ItemUID=="1" && _Form.Mode ==SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value) > Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Sum of ScanQuantity and ManualQuantity not matched with Require quantity in Line No:" + i + " ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value) >0  && ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_55").Cells.Item(i).Specific).Value == "")
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Please add Reason on why manual quantity in Line No:" + i + " ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        if (pVal.ItemUID == "34")//&& (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            _PostingDate = ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString();
                            TaxAmount = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value.ToString());
                            LineTotal = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value.ToString());
                            _DocNum = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                            DeliveryEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("1000004").Specific).Value.ToString();
                            _DocumentDate = ((SAPbouiCOM.EditText)_Form.Items.Item("38").Specific).Value.ToString();
                            _DeliveryDate = ((SAPbouiCOM.EditText)_Form.Items.Item("36").Specific).Value.ToString();

                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value) > Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_15").Cells.Item(i).Specific).Value))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Deliver quantity is greater than stock quantity in Line No:" + i + " ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                                
                            }
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value) > Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value) + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value))
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Sum of ScanQuantity and ManualQuantity not matched with Require quantity in Line No:" + i + " ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        break;
                    #endregion

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "4")
                        {
                            this.FilterBP("");
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "41" && pVal.ActionSuccess == true)  // "41" is the ItemUID of the barcode field
                        {
                            // Get the active form and the input from the barcode field
                            SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                            string scannedBarcode = barcodeField.Value.Trim().ToUpper();  // Get the complete input

                            if (!string.IsNullOrEmpty(scannedBarcode))
                            {
                                // Process the barcode
                                //ProcessBarcode(scannedBarcode);
                            }

                            // Optionally clear the input field after processing
                            barcodeField.Value = "";
                        }
                        break;
                }
            }
            #endregion

            #region After Action
            else
            {

                switch (pVal.EventType)
                {
                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1000003")
                        {
                           
                            string _BPCode = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                            FilterData(_BPCode);
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            _Form.Items.Item("10").Enabled = false;
                            Utilities.SerializedMartix(ref _Form, ref _Matrix);

                        }
                        if (pVal.ItemUID == "31")
                        {
                            Utilities.Application.SBO_Application.StatusBar.SetText("Please wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {

                                ((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(i).Specific).Checked = true;
                            }
                            Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed sucessfully", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        }
                        if (pVal.ItemUID == "32")
                        {
                            Utilities.Application.SBO_Application.StatusBar.SetText("Please wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                            {
                                ((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(i).Specific).Checked = false;
                                ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = "0";
                                ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = "0";
                            }
                            Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed sucessfully", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        }
                        if (pVal.ItemUID == "34")// && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) && pVal.Action_Success == true
                        {
                            if (DeliveryEntry == "")
                            {
                                if (Utilities.Application.SBO_Application.MessageBox("Are you sure want to post the delivery ?", 1, "OK", "Cancel") == 1)
                                {

                                    Utilities.Application.SBO_Application.StatusBar.SetText("Please wait Delivery Creation is in process", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    DeliveryCreation();

                                    this.Form.Select();
                                    if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                                    {
                                        Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                                    }
                                }
                            }
                        }
                        if (pVal.ItemUID == "1000006")
                        {
                            string _QueryQtyRefresh = "select T3.\"OnHand\",T2.\"OpenQty\" from \"@DELCH\" T0 INNER JOIN \"@DELCL\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" "+
                                                      " INNER JOIN RDR1 T2 ON T2.\"DocEntry\" = T1.\"U_DocEntry\" AND T2.\"LineNum\" = T1.\"U_LineNo\" and T2.\"ItemCode\" = T1.\"U_ItemCode\" "+
                                                      " INNER JOIN OITW T3 ON T3.\"ItemCode\" = T1.\"U_ItemCode\" and T3.\"WhsCode\" = T1.\"U_WhsCode\" "+
                                                      " where T0.\"U_DocNum\" = '"+ ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() +"'";

                            Utilities.ExecuteSQL(ref _RS, _QueryQtyRefresh);
                            if(_RS.RecordCount >0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please wait Data refresh is in process", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                _Matrix.Columns.Item("V_15").Editable = true;
                                for(int i = 1;i<= _RS.RecordCount;i++)
                                {
                                   ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_15").Cells.Item(i).Specific).Value =  _RS.Fields.Item("OnHand").Value.ToString();
                                   ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value = _RS.Fields.Item("OpenQty").Value.ToString();
                                   ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value = "N";
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "NO";
                                    _RS.MoveNext();
                                }
                                Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed sucessfully", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                _Matrix.Columns.Item("V_15").Editable = false;
                            }
                        }
                        break;
                    #endregion

                    #region et_CLICK

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ColUID == "V_14" && pVal.Row > 0)
                        {
                            if (((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific).Checked == true)
                            {
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value.ToString()) > 0)
                                {
                                    _TaxAmount = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value);
                                    SumTaxAmount = _TaxAmount + SumTaxAmount;
                                    ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = SumTaxAmount.ToString();
                                }
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific).Value.ToString()) > 0)
                                {
                                    _LineTotal = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific).Value);
                                    SumLineTotal = _LineTotal + SumLineTotal;
                                    ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = SumLineTotal.ToString();
                                }
                            }
                            else if (((SAPbouiCOM.CheckBox)_Matrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific).Checked == false)
                            {
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value.ToString()) > 0)
                                {
                                    _TaxAmount = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value);
                                    SumTaxAmount = SumTaxAmount - _TaxAmount;
                                    ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = SumTaxAmount.ToString();
                                }
                                if (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific).Value.ToString()) > 0)
                                {
                                    _LineTotal = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific).Value);
                                    SumLineTotal = SumLineTotal - _LineTotal;
                                    ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = SumLineTotal.ToString();
                                }
                            }
                        }

                        break;
                    #endregion

                    #region Set CFL Values
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "4" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_CardCode", 0, _SysCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString().Trim());
                            _MasterDBDataSource.SetValue("U_CardName", 0, _SysCFLEvent.SelectedObjects.GetValue("CardName", 0).ToString().Trim());
                        }

                        break;
                    #endregion

                    //case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                    //    if (pVal.ItemUID == "41")
                    //    {
                    //        // Get the active form
                    //        _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("14").Specific;
                    //        SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                    //        string scannedBarcode = barcodeField.Value.Trim().ToUpper();  // Convert to uppercase


                    //        if (scannedBarcode != "")
                    //        {
                    //            string[] parts = scannedBarcode.Split('/');
                    //            if (parts.Length == 2)
                    //            {
                    //                ScanCode = parts[0];
                    //                Quantity = Convert.ToDouble(parts[1]);
                    //            }
                    //            else
                    //            {
                    //                ScanCode = parts[0];
                    //                Quantity = 0;
                    //            }
                    //            CheckBarcodeInMatrix(ScanCode, Quantity);
                    //            ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value = null;
                    //            //_Form.Items.Item("41").Click();
                    //        }

                    //    }
                    //    break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        if(_Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.ColUID!="V_25" && pVal.ColUID != "V_55" && pVal.ColUID != "V_65" && pVal.ColUID != "V_1")
                        {
                            _Form.Items.Item("41").Click();
                            //_Matrix.Item.Refresh();
                        }
                        break;


                    #region KEY DOWN
                    //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    //    if (pVal.ItemUID == "41" && !pVal.BeforeAction && (pVal.CharPressed == 13 || pVal.CharPressed == 9))
                    //    //if (pVal.ItemUID == "41")
                    //    {
                    //        try
                    //        {
                    //            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("14").Specific;
                    //            SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                    //            string scannedBarcode = barcodeField.Value.Trim().ToUpper();

                    //            if (!string.IsNullOrEmpty(scannedBarcode))
                    //            {
                    //                _Form.Freeze(true);
                    //                // Split into code and quantity if present (e.g., ITEM001/5)
                    //                string[] parts = scannedBarcode.Split('/');
                    //                string ScanCode = parts[0].Trim();
                    //                double Quantity = (parts.Length == 2) ? Convert.ToDouble(parts[1]) : 0;

                    //                // Call your existing matrix update logic
                    //                CheckBarcodeInMatrix(ScanCode, Quantity);

                    //                // Clear the field for the next scan
                    //                barcodeField.Value = "";

                    //                // set focus back to the barcode field for the next scan
                    //                //_Form.Items.Item("41").Click();
                    //            }
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            Utilities.Application.SBO_Application.MessageBox("Error while processing scan: " + ex.Message);
                    //        }
                    //        finally
                    //        {
                    //            _Form.Freeze(false);
                    //            _Form.Items.Item("41").Click();
                    //        }
                    //    }
                    //    break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "41" && !pVal.BeforeAction && (pVal.CharPressed == 13 || pVal.CharPressed == 9))
                        {
                            try
                            {
                                _Form.Freeze(true);
                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("14").Specific;
                                SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                                string rawValue = barcodeField.Value.Trim();

                                if (!string.IsNullOrEmpty(rawValue) && rawValue.Length >= 6)
                                {
                                    // Position 6 = index 5 (0-based)
                                    if (rawValue[5] != '-' && !rawValue.Contains("-"))
                                    {
                                        // Insert '-' between 5th and 6th character
                                        rawValue = rawValue.Insert(5, "-");
                                    }
                                }

                                if (!string.IsNullOrEmpty(rawValue))
                                {
                                    string[] parts = rawValue.Split('/');
                                    string scanCode = parts[0].Trim();
                                    double quantity = (parts.Length == 2) ? Convert.ToDouble(parts[1]) : 0;

                                    CheckBarcodeInMatrix(scanCode, quantity);

                                    barcodeField.Value = ""; // Clear for next scan
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.Application.SBO_Application.SetStatusBarMessage("Error: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);
                                _Form.Items.Item("41").Click(); // Return focus
                            }
                        }
                        break;
                        #endregion

                        //case SAPbouiCOM.BoEventTypes.:
                        //    if (pVal.ItemUID == "41" && pVal.ActionSuccess == true)  // "41" is the ItemUID of the barcode field
                        //    {
                        //        // Get the active form and the input from the barcode field
                        //        SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                        //        string scannedBarcode = barcodeField.Value.Trim().ToUpper();  // Get the complete input

                        //        if (!string.IsNullOrEmpty(scannedBarcode))
                        //        {
                        //            // Process the barcode
                        //            //ProcessBarcode(scannedBarcode);
                        //        }

                        //        // Optionally clear the input field after processing
                        //        barcodeField.Value = "";
                        //    }
                        //    break;
                        //if (pVal.ItemUID == "41" && pVal.ActionSuccess==true)  // "41" is the ItemUID of the barcode field
                        //{
                        //    // Get the active form and the input from the barcode field
                        //    SAPbouiCOM.EditText barcodeField = (SAPbouiCOM.EditText)_Form.Items.Item("41").Specific;
                        //    string scannedBarcode = barcodeField.Value.Trim().ToUpper();  // Get the complete input

                        //    if (!string.IsNullOrEmpty(scannedBarcode))
                        //    {
                        //        // Process the barcode
                        //        //ProcessBarcode(scannedBarcode);
                        //    }

                        //    // Optionally clear the input field after processing
                        //    barcodeField.Value = "";
                        //}
                        //break;

                        //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: 
                        //if (isCapturingBarcode && !pVal.BeforeAction)
                        // {
                        //    if (pVal.CharPressed != 0 && pVal.CharPressed != 13)
                        //    {
                        //        isCapturingBarcode = true;
                        //        barcodeBuffer.Append((char)pVal.CharPressed);
                        //        //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("14").Specific;
                        //        //string barcode = barcodeBuffer.ToString().Trim();
                        //        barcodeTimer.Stop();
                        //        barcodeTimer.Start();
                        //    }
                        // }

                        // break;

                }



            }
            #endregion
        }

        #endregion

        #region Menu Event

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            _Form.Freeze(true);
            try
            {
                #region Before Action

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_Cancelled:
                        //    {
                        //        string StrSQL = "UPDATE \"@DELCH\" set \"U_Status\" = 'Cancelled' where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "'";
                        //        Utilities.ExecuteSQL(ref _RS, StrSQL);
                        //    }
                        //    break;
                    }
                }

                #endregion

                #region After Action

                else
                {
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                            {
                                _Form.Items.Item("10").Enabled = true;
                                _Form.Items.Item("1").Enabled = true;
                                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            }
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            {
                                AddModeFormValue();
                                _Form.Items.Item("1").Enabled = true;
                                _Form.Items.Item("34").Enabled = true;
                                _Form.Items.Item("31").Enabled = true;
                                _Form.Items.Item("32").Enabled = true;
                               
                                _Form.Items.Item("1000003").Enabled = true;
                                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            {
                                _Form.Items.Item("10").Enabled = false;
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("29").Specific).Value.ToString() != "" || ((SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific).Selected.Value.ToString() == "Cancelled")
                                {
                                    _Form.Items.Item("1").Enabled = false;
                                    _Form.Items.Item("34").Enabled = false;
                                    _Form.Items.Item("31").Enabled = false;
                                    _Form.Items.Item("32").Enabled = false;
                                  
                                    _Form.Items.Item("1000003").Enabled = false;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                  

                                }
                                else
                                {
                                    _Form.Items.Item("1").Enabled = true;
                                    _Form.Items.Item("34").Enabled = true;
                                    _Form.Items.Item("31").Enabled = true;
                                    _Form.Items.Item("32").Enabled = true;
                                
                                    _Form.Items.Item("1000003").Enabled = true;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                    //_Form.Items.Item("41").Click();
                                }
                            }
                            break;
                        case Constants.System_Menus.mnu_LAST:
                            {
                                _Form.Items.Item("10").Enabled = false;
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("29").Specific).Value.ToString() != "" || ((SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific).Selected.Value.ToString() == "Cancelled")
                                {
                                    _Form.Items.Item("1").Enabled = false;
                                    _Form.Items.Item("34").Enabled = false;
                                    _Form.Items.Item("31").Enabled = false;
                                    _Form.Items.Item("32").Enabled = false;
                                
                                    _Form.Items.Item("1000003").Enabled = false;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                }
                                else
                                {
                                    _Form.Items.Item("1").Enabled = true;
                                    _Form.Items.Item("34").Enabled = true;
                                    _Form.Items.Item("31").Enabled = true;
                                    _Form.Items.Item("32").Enabled = true;
                                 
                                    _Form.Items.Item("1000003").Enabled = true;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                }
                            }
                            break;
                        case Constants.System_Menus.mnu_NEXT:
                            {
                                _Form.Items.Item("10").Enabled = false;
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("29").Specific).Value.ToString() != "" || ((SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific).Selected.Value.ToString() == "Cancelled")
                                {
                                    _Form.Items.Item("1").Enabled = false;
                                    _Form.Items.Item("34").Enabled = false;
                                    _Form.Items.Item("31").Enabled = false;
                                    _Form.Items.Item("32").Enabled = false;
                                 
                                    _Form.Items.Item("1000003").Enabled = false;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                }
                                else
                                {
                                    _Form.Items.Item("1").Enabled = true;
                                    _Form.Items.Item("34").Enabled = true;
                                    _Form.Items.Item("31").Enabled = true;
                                    _Form.Items.Item("32").Enabled = true;
                                 
                                    _Form.Items.Item("1000003").Enabled = true;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                    //_Form.Items.Item("41").Click();
                                }
                            }
                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            {
                                _Form.Items.Item("10").Enabled = false;
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("29").Specific).Value.ToString() != "" || ((SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific).Selected.Value.ToString() == "Cancelled")
                                {
                                    _Form.Items.Item("1").Enabled = false;
                                    _Form.Items.Item("34").Enabled = false;
                                    _Form.Items.Item("31").Enabled = false;
                                    _Form.Items.Item("32").Enabled = false;
                                
                                    _Form.Items.Item("1000003").Enabled = false;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                }
                                else
                                {
                                    _Form.Items.Item("1").Enabled = true;
                                    _Form.Items.Item("34").Enabled = true;
                                    _Form.Items.Item("31").Enabled = true;
                                    _Form.Items.Item("32").Enabled = true;
                                 
                                    _Form.Items.Item("1000003").Enabled = true;
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                    //_Form.Items.Item("41").Click();
                                }
                            }
                            break;
                        case Constants.System_Menus.mnu_Cancelled:
                            {
                                string StrSQL = "UPDATE \"@DELCH\" set \"U_Status\" = 'Cancelled' where \"U_DocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "'";
                                Utilities.ExecuteSQL(ref _RS, StrSQL);
                            }
                            break;

                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

        #region From Default
        internal void FormDefault()
        {
            try
            {
                _Form.Freeze(true);
                _Form.EnableMenu("1292", true);
                _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@DELCH");
                _ChildDBDataSource = Form.DataSources.DBDataSources.Item("@DELCL");
                _Form.DataBrowser.BrowseBy = "10";

                ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("36").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("38").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                //((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Active = true;
                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("13").Specific;
                _ComboBox.ValidValues.Add("Primary", "Primary");
                _ComboBox.ValidValues.Add("Manual", "Manual");
                _ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific;
                _ComboBox.ValidValues.Add("Open", "Open");
                _ComboBox.ValidValues.Add("Closed", "Closed");
                _ComboBox.ValidValues.Add("Cancelled", "Cancelled");
                _ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("14").Specific;
                ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = Utilities.getMaxColumnValueNum("@DELCH", "U_DocNum");
                _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("10").Enabled = false;
                _Matrix.Columns.Item("V_13").Editable = false;
                _Matrix.Columns.Item("V_12").Editable = false;
                _Matrix.Columns.Item("V_11").Editable = false;
                _Matrix.Columns.Item("V_9").Editable = false;
                _Matrix.Columns.Item("V_8").Editable = false;
                _Matrix.Columns.Item("V_7").Editable = false;
                _Matrix.Columns.Item("V_6").Editable = false;
                _Matrix.Columns.Item("V_5").Editable = false;
                _Matrix.Columns.Item("V_4").Editable = false;
                _Matrix.Columns.Item("V_3").Editable = false;
                _Matrix.Columns.Item("V_2").Editable = false;
                _Matrix.Columns.Item("V_10").Editable = false;
                _Matrix.Columns.Item("V_16").Editable = false;
                _Matrix.Columns.Item("V_15").Editable = false;
                _Matrix.Columns.Item("V_0").Editable = false;
                _Matrix.Columns.Item("V_20").Editable = false;
                _Matrix.Columns.Item("V_22").Editable = false;
                _Matrix.Columns.Item("V_23").Editable = false;
                _Form.Items.Item("6").Enabled = false;
                _Form.Items.Item("1000002").Enabled = false;
                _Form.Items.Item("29").Enabled = false;
                _Form.Items.Item("1000004").Enabled = false;
                _Form.Items.Item("27").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("36").Enabled = false;
                _Form.Items.Item("38").Enabled = false;
                _Form.Items.Item("22").Enabled = false;
                _Form.Items.Item("24").Enabled = false;
           
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
              
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }
        #endregion

        #region AddModeFormValue
        internal void AddModeFormValue()
        {
            try
            {
                _Form.Freeze(true);
                _Form.EnableMenu("1292", true);
                _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@DELCH");
                _ChildDBDataSource = Form.DataSources.DBDataSources.Item("@DELCL");
                _Form.DataBrowser.BrowseBy = "10";

                ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("36").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("38").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");

                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("13").Specific;
                _ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("27").Specific;
                _ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("14").Specific;
                ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = Utilities.getMaxColumnValueNum("@DELCH", "U_DocNum");
                _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("10").Enabled = false;
                _Matrix.Columns.Item("V_13").Editable = false;
                _Matrix.Columns.Item("V_12").Editable = false;
                _Matrix.Columns.Item("V_11").Editable = false;
                _Matrix.Columns.Item("V_9").Editable = false;
                _Matrix.Columns.Item("V_8").Editable = false;
                _Matrix.Columns.Item("V_7").Editable = false;
                _Matrix.Columns.Item("V_6").Editable = false;
                _Matrix.Columns.Item("V_5").Editable = false;
                _Matrix.Columns.Item("V_4").Editable = false;
                _Matrix.Columns.Item("V_3").Editable = false;
                _Matrix.Columns.Item("V_2").Editable = false;
                _Matrix.Columns.Item("V_10").Editable = false;
                _Matrix.Columns.Item("V_16").Editable = false;
                _Matrix.Columns.Item("V_15").Editable = false;
                _Matrix.Columns.Item("V_0").Editable = false;
                _Matrix.Columns.Item("V_20").Editable = false;
                _Matrix.Columns.Item("V_22").Editable = false;
                _Matrix.Columns.Item("V_23").Editable = false;
                _Form.Items.Item("6").Enabled = false;
                _Form.Items.Item("1000002").Enabled = false;
                _Form.Items.Item("29").Enabled = false;
                _Form.Items.Item("1000004").Enabled = false;
                _Form.Items.Item("27").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("36").Enabled = false;
                _Form.Items.Item("38").Enabled = false;
                _Form.Items.Item("22").Enabled = false;
                _Form.Items.Item("24").Enabled = false;
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }
        #endregion

        #region Filter Data
        internal void FilterData(string _BPCode)
        {
            try
            {
                _Form.Freeze(true);
                string QueryFillData = string.Empty;
                
                _Matrix.Clear();
                Utilities.Application.SBO_Application.StatusBar.SetText("Please wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("14").Specific;
                
                    QueryFillData = "select T0.\"DocType\",T0.\"CardCode\",T0.\"CardName\",T0.\"NumAtCard\",T0.\"DocCur\",T0.\"DocDate\",T0.\"DocDueDate\",T0.\"TaxDate\", " +
                                    "T0.\"DocTotal\",T0.\"DiscSum\",T0.\"VatPaid\",T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"Price\", " +
                                    "T1.\"LineTotal\",T1.\"DiscPrcnt\",T1.\"Currency\",T1.\"LineNum\",T1.\"DocEntry\",T0.\"DocNum\",T1.\"LineVat\", " +
                                    "T3.\"DfltWH\" as \"WhsCode\",T1.\"OpenCreQty\",(T1.\"Quantity\" - T1.\"OpenCreQty\") as \"DelQty\",T2.\"OnHand\",'Y' as \"Status\",'NO' as \"ScanStatus\"  from ORDR T0 " +
                                    "INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" and T1.\"LineStatus\" = 'O' " +
                                    "INNER JOIN OITM T3 ON T1.\"ItemCode\" = T3.\"ItemCode\" AND T3.\"ManSerNum\" = 'N' AND T3.\"ManBtchNum\" = 'N' " +
                                    "INNER JOIN OITW T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" and T2.\"WhsCode\" = T3.\"DfltWH\" " +
                                    "where T2.\"OnHand\" > 0 AND T0.\"DocStatus\" = 'O' and T0.\"CardCode\" = '"+ _BPCode +"'";
             
                Utilities.ExecuteSQL(ref _RS, QueryFillData);

                Utilities.Application.SBO_Application.StatusBar.SetText("Please wait " + _RS.RecordCount + " lines data loading", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                
                for (int i = 0; i < _RS.RecordCount; i++)
                {
                    _Matrix.AddRow(1, -1);
                    _Matrix.FlushToDataSource();
                    _ChildDBDataSource.SetValue("U_ItemCode", i, _RS.Fields.Item("ItemCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemName", i, _RS.Fields.Item("Dscription").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Qty", i, _RS.Fields.Item("Quantity").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Price", i, _RS.Fields.Item("Price").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LineTot", i, _RS.Fields.Item("LineTotal").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LineNo", i, _RS.Fields.Item("LineNum").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocNum", i, _RS.Fields.Item("DocNum").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DocEntry", i, _RS.Fields.Item("DocEntry").Value.ToString());
                    _ChildDBDataSource.SetValue("U_TaxAmt", i, _RS.Fields.Item("LineVat").Value.ToString());
                    _ChildDBDataSource.SetValue("U_OrdDt", i, DateTime.Parse(_RS.Fields.Item("DocDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_DelDt", i, DateTime.Parse(_RS.Fields.Item("DocDueDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource.SetValue("U_OpnQty", i, _RS.Fields.Item("OpenCreQty").Value.ToString());
                    _ChildDBDataSource.SetValue("U_StockQty", i, _RS.Fields.Item("OnHand").Value.ToString());
                    _ChildDBDataSource.SetValue("U_DelQty", i, _RS.Fields.Item("DelQty").Value.ToString());
                    _ChildDBDataSource.SetValue("U_WhsCode", i, _RS.Fields.Item("WhsCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LnStatus", i, _RS.Fields.Item("Status").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ScanStatus", i, _RS.Fields.Item("ScanStatus").Value.ToString());
                    _Matrix.LoadFromDataSource();
                    _RS.MoveNext();
                }

                _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("10").Enabled = false;
                
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                Utilities.Message("Operation completed sucessfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }
        #endregion

        #region right click event

        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
            try
            {
                this.RowToDelete = oEventInfo.Row;
            }
            catch (Exception)
            {

                throw;
            }
        }
        #endregion right click event

        #region FILTER BP Type
        internal void FilterBP(string CFLId)
        {
            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "CardType";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "C";
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("CFL_1").SetConditions(_Cons);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region ItemInvoiceCreation

        public void DeliveryCreation()
        {
            string DelQuery = "select T1.\"U_DocNum\",T1.\"U_DocEntry\",T0.\"U_DocDate\",T0.\"U_PstDt\",T0.\"U_DelDt\",T0.\"U_CardCode\",T0.\"U_CardName\", " +
                                    "T0.\"U_BPRefNo\",T0.\"U_Series\",T0.\"U_TaxAmt\",T0.\"U_DocTotal\",T1.\"U_ItemCode\",T1.\"U_ItemName\",T1.\"U_Qty\",T1.\"U_Price\"," +
                                    "T1.\"U_LineTot\",T1.\"U_LineNo\",T1.\"U_TaxAmt\",T1.\"U_DelQty\",T1.\"U_WhsCode\",T1.\"U_Chk\",T1.\"U_StockQty\",T1.\"U_OpnQty\",T1.\"U_ScanStatus\",T1.\"U_ScanQty\",T1.\"U_ManualQty\" from \"@DELCH\" T0 INNER JOIN \"@DELCL\" T1 ON " +
                                    "T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"U_DocNum\" = '" + _DocNum + "' order by \"U_DocEntry\",\"U_LineNo\" ";
            Utilities.ExecuteSQL(ref _RSDel, DelQuery);
            SAPbobsCOM.Documents _Delivery = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);

            for (int i = 0; i < _RSDel.RecordCount; i++)
            {
                if (_RSDel.Fields.Item("U_Chk").Value.ToString() == "Y")
                {
                    if (i == 0)
                    {
                        _Delivery.CardCode = _RSDel.Fields.Item("U_CardCode").Value.ToString().Trim();
                        _Delivery.CardName = _RSDel.Fields.Item("U_CardName").Value.ToString().Trim();
                        _Delivery.DocDate = Utilities.ToDate(_DocumentDate);
                        
                        _Delivery.TaxDate = Utilities.ToDate(_PostingDate);
                        _Delivery.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                    }
                    if (i > 0)
                    _Delivery.Lines.Add();
                    _Delivery.Lines.ItemCode = _RSDel.Fields.Item("U_ItemCode").Value.ToString();
                    _Delivery.Lines.ItemDescription = _RSDel.Fields.Item("U_ItemName").Value.ToString();
                    _Delivery.Lines.Quantity = double.Parse(_RSDel.Fields.Item("U_OpnQty").Value.ToString());
                    _Delivery.Lines.UnitPrice = double.Parse(_RSDel.Fields.Item("U_Price").Value.ToString());
                    //_Delivery.Lines.LineTotal = double.Parse(_RSDel.Fields.Item("U_LineTot").Value.ToString());
                    _Delivery.Lines.WarehouseCode = _RSDel.Fields.Item("U_WhsCode").Value.ToString();
                    _Delivery.Lines.UserFields.Fields.Item("U_ScanStatus").Value= _RSDel.Fields.Item("U_ScanStatus").Value.ToString();
                    _Delivery.Lines.UserFields.Fields.Item("U_ScanQty").Value = _RSDel.Fields.Item("U_ScanQty").Value.ToString();
                    _Delivery.Lines.UserFields.Fields.Item("U_ManualQty").Value = _RSDel.Fields.Item("U_ManualQty").Value.ToString();
                    _Delivery.Lines.BaseType = 17;
                    _Delivery.Lines.BaseEntry = Convert.ToInt32(_RSDel.Fields.Item("U_DocEntry").Value.ToString());
                    _Delivery.Lines.BaseLine = Convert.ToInt32(_RSDel.Fields.Item("U_LineNo").Value.ToString());
                }
                _RSDel.MoveNext();
            }

            int lRetCode = _Delivery.Add();
            if (lRetCode == 0)
            {
                string objCode = "";
                Utilities.Application.Company.GetNewObjectCode(out objCode);

                string StrSQLDel = "select \"DocNum\" from ODLN where \"DocEntry\" = '"+ objCode +"'";
                Utilities.ExecuteSQL(ref _RSDelNo, StrSQLDel);

                string StrSQL = "UPDATE \"@DELCH\" set \"U_DelNum\" = '" + _RSDelNo.Fields.Item("DocNum").Value.ToString().Trim() + "', \"U_DelEntry\" = '" + objCode.ToString() + "',\"U_Status\" = 'Closed' where \"U_DocNum\" = '" + _DocNum + "'";
                Utilities.ExecuteSQL(ref _RS, StrSQL);
                Utilities.WriteLogFile("Delivery Created Sucessfully,Delivery No:" + _RSDelNo.Fields.Item("DocNum").Value.ToString().Trim() + ": For Customer Code : " + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "    :" + DateTime.Now);
                Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            else
            {
                Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
              //  Utilities.WriteLogFileForInterest("Invoice Creation Error is :" + _LastErrorMsg + ":For Customer:" + _Grid.DataTable.Columns.Item("Customer Code").Cells.Item(i).Value.ToString().Trim() + ":" + DateTime.Now);
                Utilities.WriteLogFile("Delivery Creation Error is:" + _LastErrorMsg + ": For Customer Code : " + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "   :" + DateTime.Now);
                Utilities.Application.SBO_Application.StatusBar.SetText(_LastErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                               
            }
        }

        #endregion

        private void CheckBarcodeInMatrix(string scannedBarcode, double qtyToScan)
        {
            bool itemFoundInMatrix = false;
            double remainingQty = (qtyToScan > 0) ? qtyToScan : 1; // If 0 is passed, we treat it as a single scan (1)
            bool updateHappened = false;

            // 1. Loop through the matrix once
            for (int i = 1; i <= _Matrix.RowCount; i++)
            {
                // Get values from the current row
                string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_13").Cells.Item(i).Specific).Value.Trim();

                // Only process if the item code matches
                if (scannedBarcode.Equals(itemCode, StringComparison.OrdinalIgnoreCase))
                {
                    itemFoundInMatrix = true;

                    double reqQTY = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value);
                    double currentQty = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value);
                    double manualQty = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value);

                    double alreadyFilled = currentQty + manualQty;
                    double spaceInThisRow = reqQTY - alreadyFilled;

                    if (spaceInThisRow > 0)
                    {
                        // Calculate how much we can add to THIS specific row
                        double amountToAdd = Math.Min(remainingQty, spaceInThisRow);

                        // Update the Matrix Cells
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "YES";
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value = (currentQty + amountToAdd).ToString();
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_22").Cells.Item(i).Specific).Value = scannedBarcode;

                        remainingQty -= amountToAdd;
                        updateHappened = true;

                        // If we have fulfilled the scanned quantity, exit the loop
                        if (remainingQty <= 0)
                            break;
                    }
                }
            }

            // 2. Error Handling based on the results of the loop
            if (!itemFoundInMatrix)
            {
                Utilities.Application.SBO_Application.MessageBox($"Scanned item {scannedBarcode} is not found in the list.");
            }
            else if (!updateHappened)
            {
                // This means the item was found, but all rows were already full
                Utilities.Application.SBO_Application.MessageBox($"Quantity for item {scannedBarcode} has already been fulfilled in all rows.");
            }
            else if (remainingQty > 0)
            {
                // This means we updated some rows, but we still have leftover quantity because the matrix is full
                Utilities.Application.SBO_Application.StatusBar.SetText($"Partial update: {remainingQty} could not be added as all rows are full.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }











        //private void CheckBarcodeInMatrix(string scannedBarcode, double Qty)
        //{
        //    bool allRowsMatched = true; // Track if all rows with same item code have required quantity
        //    bool itemFound = false;     // Track if the item was found in any row

        //    for (int i = 1; i <= _Matrix.RowCount; i++)
        //    {
        //        string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_13").Cells.Item(i).Specific).Value.Trim();
        //        double reqQTY = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value);
        //        double currentQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value);
        //        double ManualQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value);

        //        if (scannedBarcode.ToUpper() == itemCode.ToUpper())
        //        {
        //            itemFound = true; // Item found at least in one row

        //            // Check if the current quantity is already equal to or greater than the required quantity
        //            if (currentQuantity + ManualQuantity < reqQTY)
        //            {
        //                allRowsMatched = false; // There's at least one row where quantity is less than required
        //                break; // No need to continue if one row doesn't meet the condition
        //            }
        //        }
        //    }

        //    if (!itemFound)
        //    {
        //        Utilities.Application.SBO_Application.MessageBox("Scanned item " + scannedBarcode + " is not found.");
        //        return;
        //    }

        //    // If all matching rows already have the required quantity
        //    if (itemFound && allRowsMatched)
        //    {
        //        //Console.Beep(1000, 500);
        //        Utilities.Application.SBO_Application.MessageBox("Quantity for the item " + scannedBarcode + " has already reached the required quantity in all rows.");
        //        return; // Exit the function, no further processing needed
        //    }

        //    // Proceed with the existing logic for updating the matrix if quantity has not been fulfilled in all rows
        //    for (int i = 1; i <= _Matrix.RowCount; i++)
        //    {
        //        string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_13").Cells.Item(i).Specific).Value.Trim();
        //        double reqQTY = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value);
        //        double currentQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value);
        //        double currentQuantityBulk = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value);
        //        double ManualQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_25").Cells.Item(i).Specific).Value);
        //        if (scannedBarcode.ToUpper() == itemCode.ToUpper())
        //        {
        //            currentQuantity++; // Increment the quantity for this row

        //            if (Qty > 0)
        //            {
        //                if (reqQTY >= (currentQuantityBulk+ManualQuantity + Qty))
        //                {
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "YES";
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value = (currentQuantityBulk + Qty).ToString();
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_22").Cells.Item(i).Specific).Value = scannedBarcode.ToString();

        //                    if (scannedBarcode.ToUpper() == itemCode.ToUpper())
        //                    {
        //                        currentQuantity++; // Increment the quantity for this row

        //                        if (Qty > 0)
        //                        {
        //                            if (reqQTY >= (currentQuantityBulk + ManualQuantity + Qty))
        //                            {
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "YES";
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value = (currentQuantityBulk + Qty).ToString();
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_22").Cells.Item(i).Specific).Value = scannedBarcode.ToString();
        //                                _Status = true;
        //                                //_Matrix.LoadFromDataSource();
        //                            }
        //                            else
        //                            {
        //                                _Status = false;
        //                            }
        //                        }
        //                        else
        //                        {
        //                            if (reqQTY >= (currentQuantity + ManualQuantity))
        //                            {
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "YES";
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value = currentQuantity.ToString();
        //                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_22").Cells.Item(i).Specific).Value = scannedBarcode.ToString();

        //                               // _Matrix.LoadFromDataSource();
        //                                _Status = true;
        //                            }
        //                            else
        //                            {
        //                                _Status = false;
        //                            }
        //                        }
        //                    }
        //                    _Status = true;

        //                }
        //                else
        //                {
        //                    _Status = false;
        //                }
        //            }
        //            else
        //            {
        //                if (reqQTY >= (currentQuantity+ManualQuantity))
        //                {
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_20").Cells.Item(i).Specific).Value = "YES";
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_23").Cells.Item(i).Specific).Value = currentQuantity.ToString();
        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_22").Cells.Item(i).Specific).Value = scannedBarcode.ToString();

        //                    //_Matrix.LoadFromDataSource();
        //                    _Status = true;
        //                }
        //                else
        //                {
        //                    _Status = false;
        //                }
        //            }
        //        }

        //        if (i == _Matrix.RowCount && _Status == false)
        //        {
        //            Utilities.Application.SBO_Application.MessageBox("Quantity exceeds for item: " + scannedBarcode + " with given quantity:" + reqQTY + " in Autodelivery");
        //            Utilities.Application.SBO_Application.SetStatusBarMessage("Quantity exceeds the limit!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //    }
        //}



        private void OnBarcodeTimerElapsed(object sender, ElapsedEventArgs e)
        {
            barcodeTimer.Stop();
            string barcode = barcodeBuffer.ToString().Trim();
            barcodeBuffer.Clear();
        }


    } 

}
        



    

