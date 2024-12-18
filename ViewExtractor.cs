using System;
using System.Runtime.InteropServices;
using Inventor;

namespace ViewExtraction
{
    [Guid("37B44353-D112-4260-B500-4A63FCAA2841")]
    [ComVisible(true)]
    public class ViewExtractor : ApplicationAddInServer
    {
        private Inventor.Application inventorApp;
        private ButtonDefinition myButton;
        private RibbonPanel customPanel;
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // called when application loads dll file in memory
            // has definitions of UI components
            inventorApp = addInSiteObject.Application;

            Ribbon partRibbon = inventorApp.UserInterfaceManager.Ribbons["Part"];
            RibbonTab assembleTab = partRibbon.RibbonTabs["id_TabTools"];
            customPanel = assembleTab.RibbonPanels.Add("Custom Tools", "CustomPanel", "");

            myButton = inventorApp.CommandManager.ControlDefinitions.AddButtonDefinition("Export Views", "Exports all othgonal views in DWG file", CommandTypesEnum.kNonShapeEditCmdType);

            customPanel.CommandControls.AddButton(myButton);

            myButton.OnExecute += MyButton_OnExecute;
        }

        private void MyButton_OnExecute(NameValueMap Context)
        {
            Document activeDoc = inventorApp.ActiveDocument;

            if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                try
                {

                    DrawingDocument drawDoc = (DrawingDocument)inventorApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, inventorApp.FileManager.
                                                GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject), true);
                    Sheet sheet = drawDoc.Sheets[1];

                    Point2d pt1 = inventorApp.TransientGeometry.CreatePoint2d(5, 15);
                    Point2d pt2 = inventorApp.TransientGeometry.CreatePoint2d(15, 15);
                    Point2d pt3 = inventorApp.TransientGeometry.CreatePoint2d(30, 15);

                    sheet.DrawingViews.AddBaseView((_Document)activeDoc, pt1, 20,
                        ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

                    sheet.DrawingViews.AddBaseView((_Document)activeDoc, pt2, 20,
                        ViewOrientationTypeEnum.kBottomViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

                    sheet.DrawingViews.AddBaseView((_Document)activeDoc, pt3, 20,
                        ViewOrientationTypeEnum.kLeftViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

                    drawDoc.SaveAs(@"D:\narayanWorkspace\Inventor_Practise\API_Practice\ViewExtractor\OutputFile.dwg", true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }
        
        public void Deactivate()
        {
            if (inventorApp != null)
            {
                RemoveRibbonButton();
            }

            inventorApp = null;
        }

        private void RemoveRibbonButton()
        {
            try
            {
                UserInterfaceManager uiManager = (UserInterfaceManager)inventorApp.UserInterfaceManager;
                Ribbon partRibbon = uiManager.Ribbons["Part"];
                RibbonTab assembleTab = partRibbon.RibbonTabs["id_TabTools"];

                foreach (RibbonPanel panel in assembleTab.RibbonPanels)
                {
                    if (panel.DisplayName == "Custom Tools")
                    {
                        panel.Delete();  //Delete method to remove the panel
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during cleanup: " + ex.Message);
            }
        }

        public void ExecuteCommand(int commandID)
        {
            throw new NotImplementedException();
        }

        public object Automation => null;
    }
}