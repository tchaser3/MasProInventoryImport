/* Title:           Import Maspro Inventory
 * Date:            5-13-19
 * Author:          Terrance Holmes
 * 
 * Description:     This program will import masspro from a spreadsheet */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using InventoryDLL;
using NewEventLogDLL;
using NewEmployeeDLL;
using Excel = Microsoft.Office.Interop.Excel;
using NewPartNumbersDLL;

namespace ImportMassProInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        InventoryClass TheInventoryClass = new InventoryClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        PartNumberClass ThePartNumberClass = new PartNumberClass();

        //setting up the data
        FindPartsWarehousesDataSet TheFindPartsWarehousesDataSet = new FindPartsWarehousesDataSet();
        FindWarehouseInventoryPartDataSet TheFindWarehouseInventoryPartDataSet = new FindWarehouseInventoryPartDataSet();
        FindPartByPartIDDataSet TheFindPartByPartIDDataSet = new FindPartByPartIDDataSet();
        ImportInventoryDataSet TheImportInventoryDataSet = new ImportInventoryDataSet();
        ItemsNotFoundDataSet TheItemsNotFoundDataSet = new ItemsNotFoundDataSet();
        FindWarehouseInventoryDataSet TheFindWarehouseInventoryDataSet = new FindWarehouseInventoryDataSet();

        int gintWarehouseID;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;

            TheFindPartsWarehousesDataSet = TheEmployeeClass.FindPartsWarehouses();

            intNumberOfRecords = TheFindPartsWarehousesDataSet.FindPartsWarehouses.Rows.Count - 1;
            cboSelectWarehouse.Items.Add("Select Warehouse");

            for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                cboSelectWarehouse.Items.Add(TheFindPartsWarehousesDataSet.FindPartsWarehouses[intCounter].FirstName);
            }

            cboSelectWarehouse.SelectedIndex = 0;
            btnImportExcel.IsEnabled = false;
            btnProcess.IsEnabled = false;
        }

        private void CboSelectWarehouse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;
            string strCurrentWarehouse = "The Current Warehouse is :  ";

            intSelectedIndex = cboSelectWarehouse.SelectedIndex - 1;

            if(intSelectedIndex > -1)
            {
                gintWarehouseID = TheFindPartsWarehousesDataSet.FindPartsWarehouses[intSelectedIndex].EmployeeID;
                strCurrentWarehouse += TheFindPartsWarehousesDataSet.FindPartsWarehouses[intSelectedIndex].FirstName;
                lblCurrentWarehouse.Content = strCurrentWarehouse;

                btnImportExcel.IsEnabled = true;
            }
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strPartID;
            int intPartID;
            string strQuantity;
            int intQuantity;
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            string strPartNumber;
            string strJDEPartNumber;
            string strPartDescription;
            string strJDEPartNumberFromSpreadSheet;
            int intCurrentQuantity;

            try
            {
                TheImportInventoryDataSet.importinventory.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count + 1;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter < intNumberOfRecords; intCounter++)
                {
                    strPartID = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);
                    intPartID = Convert.ToInt32(strPartID);
                    strQuantity = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2);
                    intQuantity = Convert.ToInt32(strQuantity);
                    strJDEPartNumberFromSpreadSheet = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2);

                    TheFindPartByPartIDDataSet = ThePartNumberClass.FindPartByPartID(intPartID);

                    intRecordsReturned = TheFindPartByPartIDDataSet.FindPartByPartID.Rows.Count;

                    if (intRecordsReturned == 0)
                    {

                        TheMessagesClass.ErrorMessage("Part ID " + strPartID + " Was Not Found");
                        return;
                    }

                    strPartNumber = TheFindPartByPartIDDataSet.FindPartByPartID[0].PartNumber;
                    strJDEPartNumber = TheFindPartByPartIDDataSet.FindPartByPartID[0].JDEPartNumber;
                    strPartDescription = TheFindPartByPartIDDataSet.FindPartByPartID[0].PartDescription;

                    if(strJDEPartNumber != strJDEPartNumberFromSpreadSheet)
                    {
                        TheMessagesClass.ErrorMessage("JDE Part Numbers " + strJDEPartNumber + "and " + strJDEPartNumberFromSpreadSheet + " Do Not Match");
                        return;
                    }

                    TheFindWarehouseInventoryPartDataSet = TheInventoryClass.FindWarehouseInventoryPart(intPartID, gintWarehouseID);

                    intRecordsReturned = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart.Rows.Count;

                    if(intRecordsReturned == 0)
                    {
                        intCurrentQuantity = 0;
                    }
                    else
                    {
                        intCurrentQuantity = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart[0].Quantity;
                    }

                    ImportInventoryDataSet.importinventoryRow NewPartRow = TheImportInventoryDataSet.importinventory.NewimportinventoryRow();

                    NewPartRow.CurrentQuantity = intCurrentQuantity;
                    NewPartRow.NewQuantity = intQuantity;
                    NewPartRow.JDEPartNumber = strJDEPartNumber;
                    NewPartRow.PartDescription = strPartDescription;
                    NewPartRow.PartID = intPartID;
                    NewPartRow.PartNumber = strPartNumber;

                    TheImportInventoryDataSet.importinventory.Rows.Add(NewPartRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportInventoryDataSet.importinventory;
                btnProcess.IsEnabled = true;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import MasPro Inventory // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intPartID;
            int intRecordsReturned;
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            int intQuantity;
            int intTransactionID;

            try
            {
                CompareParts();

                intNumberOfRecords = TheImportInventoryDataSet.importinventory.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intPartID = TheImportInventoryDataSet.importinventory[intCounter].PartID;
                    intQuantity = TheImportInventoryDataSet.importinventory[intCounter].NewQuantity;

                    TheFindWarehouseInventoryPartDataSet = TheInventoryClass.FindWarehouseInventoryPart(intPartID, gintWarehouseID);

                    intRecordsReturned = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart.Rows.Count;

                    if (intRecordsReturned == 0)
                    {
                        blnFatalError = TheInventoryClass.InsertInventoryPart(intPartID, intQuantity, gintWarehouseID);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                    else if (intRecordsReturned == 1)
                    {
                        intTransactionID = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart[0].TransactionID;

                        blnFatalError = TheInventoryClass.UpdateInventoryPart(intTransactionID, intQuantity);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                    else if (intRecordsReturned > 1)
                    {
                        TheMessagesClass.ErrorMessage("fuck you");
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Maspro Inventory // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void CompareParts()
        {
            //setting local variables
            int intFirstCounter;
            int intSecondCounter;
            int intFirstNumberOfRecords;
            int intSecondNumberOfRecords;
            bool blnItemFound;
            int intPartID;
            int intWarehouseID = 0;
            bool blnFatalError = false;

            try
            {
                TheItemsNotFoundDataSet.itemsnotfound.Rows.Clear();

                TheFindWarehouseInventoryDataSet = TheInventoryClass.FindWarehouseInventory(gintWarehouseID);

                intFirstNumberOfRecords = TheFindWarehouseInventoryDataSet.FindWarehouseInventory.Rows.Count - 1;
                intSecondNumberOfRecords = TheImportInventoryDataSet.importinventory.Rows.Count - 1;

                for(intFirstCounter = 0; intFirstCounter <= intFirstNumberOfRecords; intFirstCounter++)
                {
                    blnItemFound = false;
                    intPartID = TheFindWarehouseInventoryDataSet.FindWarehouseInventory[intFirstCounter].PartID;

                    for(intSecondCounter = 0; intSecondCounter <= intSecondNumberOfRecords; intSecondCounter++)
                    {
                        if(intPartID == TheImportInventoryDataSet.importinventory[intSecondCounter].PartID)
                        {
                            blnItemFound = true;
                        }
                    }

                    if(blnItemFound == false)
                    {
                        ItemsNotFoundDataSet.itemsnotfoundRow PartNotFound = TheItemsNotFoundDataSet.itemsnotfound.NewitemsnotfoundRow();

                        PartNotFound.PartID = intPartID;
                        PartNotFound.PartNumber = TheFindWarehouseInventoryDataSet.FindWarehouseInventory[intFirstCounter].PartNumber;
                        PartNotFound.PartDescription = TheFindWarehouseInventoryDataSet.FindWarehouseInventory[intFirstCounter].PartDescription;
                        PartNotFound.JDEPartNumber = TheFindWarehouseInventoryDataSet.FindWarehouseInventory[intFirstCounter].JDEPartNumber;
                        PartNotFound.Quantity = TheFindWarehouseInventoryDataSet.FindWarehouseInventory[intFirstCounter].Quantity;

                        TheItemsNotFoundDataSet.itemsnotfound.Rows.Add(PartNotFound);
                    }
                }

                intFirstNumberOfRecords = TheFindPartsWarehousesDataSet.FindPartsWarehouses.Rows.Count - 1;

                for(intFirstCounter = 0; intFirstCounter <= intFirstNumberOfRecords; intFirstCounter++)
                {
                    if(TheFindPartsWarehousesDataSet.FindPartsWarehouses[intFirstCounter].FirstName == "MASPRO-HOLDING")
                    {
                        intWarehouseID = TheFindPartsWarehousesDataSet.FindPartsWarehouses[intFirstCounter].EmployeeID;
                    }
                }

                intFirstNumberOfRecords = TheItemsNotFoundDataSet.itemsnotfound.Rows.Count - 1;

                if(intFirstNumberOfRecords > -1)
                {
                    for (intFirstCounter = 0; intFirstCounter <= intFirstNumberOfRecords; intFirstCounter++)
                    {
                        intPartID = TheItemsNotFoundDataSet.itemsnotfound[intFirstCounter].PartID;

                        blnFatalError = TheInventoryClass.MovePartToNewWarehouse(gintWarehouseID, intWarehouseID, intPartID);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                }

                
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Mass Pro Inventory // Compare Parts " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
