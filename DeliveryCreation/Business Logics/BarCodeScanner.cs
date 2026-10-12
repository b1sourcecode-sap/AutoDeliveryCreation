using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeliveryCreation
{
    public class BarcodeScanner
    {
        private StringBuilder barcodeBuffer = new StringBuilder();
        private SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.Form _Form;
        private string ScanCode;
        private double Quantity;
        private SAPbouiCOM.Application _Application;

        private void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                    // When the form is activated, get the form reference
                    if (!pVal.BeforeAction)
                    {
                        _Form = Utilities.Application.SBO_Application.Forms.Item("60004");
                    }
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    if (!pVal.BeforeAction)
                    {
                        // Accumulate characters in barcodeBuffer
                        if (pVal.CharPressed != 0 && pVal.CharPressed != 13) // 13 = Enter key
                        {
                            barcodeBuffer.Append((char)pVal.CharPressed);
                        }
                        else if (pVal.CharPressed == 13) // Enter key pressed, assume barcode complete
                        {
                            ProcessScannedBarcode();
                        }
                    }
                    break;
            }
        }

        private void ProcessScannedBarcode()
        {
            string scannedBarcode = barcodeBuffer.ToString().Trim().ToUpper();  // Convert to uppercase

            if (!string.IsNullOrEmpty(scannedBarcode))
            {
                string[] parts = scannedBarcode.Split('/');
                if (parts.Length == 2)
                {
                    ScanCode = parts[0];
                    Quantity = Convert.ToDouble(parts[1]);
                }
                else
                {
                    ScanCode = parts[0];
                    Quantity = 0;
                }

                CheckBarcodeInMatrix(ScanCode, Quantity);

                // Clear the buffer after processing
                barcodeBuffer.Clear();

                // Clear the value in the field if needed
                ((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Value = string.Empty;
            }
        }

        private void CheckBarcodeInMatrix(string scanCode, double quantity)
        {
            // Your logic for checking the barcode in the matrix
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("14").Specific;

            // Example: Find and update the matrix row based on the scanned barcode
            for (int i = 1; i <= _Matrix.RowCount; i++)
            {
                SAPbouiCOM.EditText cell = (SAPbouiCOM.EditText)_Matrix.Columns.Item("YourColumnID").Cells.Item(i).Specific;
                if (cell.Value == scanCode)
                {
                    // Update the quantity or other fields as needed
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("QuantityColumnID").Cells.Item(i).Specific).Value = quantity.ToString();
                    break;
                }
            }
        }
    }
}
