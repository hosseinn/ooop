package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class ListenerOfDialogs implements  XDialogEventHandler, XTopWindowListener {


    private GuiOfDialogs    m_Gui               = null;
    private Controller      m_Controller        = null;


    ListenerOfDialogs(GuiOfDialogs gui, Controller controller){
        m_Gui = gui;
        m_Controller = controller;
    }
    
    public GuiOfDialogs getGui(){
        return m_Gui;
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    // XDialogEventHandler
    @Override
    public String[] getSupportedMethodNames() {

        String[] aMethods = new String[173];

        //DiagramGallery2 events **************************************
        aMethods[0] = "itemChangedInList";
        for(int i = 161; i < 161 + GuiOfDialogs.NUM_OF_DIAGRAMS; i++)
            aMethods[i] = "item" + (i - 161) + "Pressed";
        //*************************************************************

        //ControlDialog1, ControlDialog2 common events ****************
        aMethods[1] = "addShape";
        aMethods[2] = "removeShape";
        aMethods[3] = "convert";
        aMethods[4] = "setProperties";
        //aMethods[5] = "downUpAction";
        aMethods[5] = "export";
        aMethods[6] = "textFieldModified";
        //short cuts - hidden buttons:
        aMethods[7] = "allShape";
        aMethods[8] = "previous";
        aMethods[9] = "next";
        aMethods[10] = "edit";
        aMethods[11] = "size";
        //************************************************************

        //ControlDialog1 events ***************************************
        aMethods[12] = "showColorTable3";
        //short cuts - hidden buttons:
        aMethods[13] = "first";
        aMethods[14] = "mainShape";
        aMethods[15] = "textShape";
        //*************************************************************

        //ControlDialog2 events ***************************************
        aMethods[16] = "showColorTable2";
        aMethods[17] = "changedAddProp";
        //short cuts - hidden buttons:
        aMethods[18] = "up";
        aMethods[19] = "down";
        //*************************************************************

        //ColorTables common events ***********************************
        for(int i = 1; i <= 124; i++)
            aMethods[i+19] = "image" +i;
        //*************************************************************

        //ColorTable2 events ******************************************
        aMethods[144] = "setGradientMode";
        //*************************************************************

        //ColorTable3 events ******************************************
        aMethods[145] = "setSimpleColorModeOptionButton";
        aMethods[146] = "setBaseColorsModeOptionButton";
        aMethods[147] = "setBaseColorsGradientsModeOptionButton";
        aMethods[148] = "setColorThemeModeOptionButton";
        aMethods[149] = "setColorSchemeModeOptionButton";
        aMethods[150] = "setBaseColorsSettings";
        aMethods[151] = "colorTableHelp";
        aMethods[152] = "saveColorModeAndExit";
        //*************************************************************

        //GradientDialog events ***************************************
        aMethods[153] = "setColorMode";
        aMethods[154] = "setGradientSettings";
        aMethods[155] = "setStartColor";
        aMethods[156] = "setEndColor";
        //*************************************************************

        //ConvertDiaglog events ***************************************
        aMethods[157] = "converButton1Pressed";
        aMethods[158] = "converButton2Pressed";
        //************************************************************

        //BaseColorsSettingsDialog events *****************************
        aMethods[159] = "setDefaultBaseColorsSettings";
        aMethods[160] = "showColorTable";
        //*************************************************************
        
        return aMethods;
    }

    // XDialogEventHandler
    @Override
    public boolean callHandlerMethod(XDialog xDialog, Object eventObject, String methodName) throws WrappedTargetException {

        //DiagramGallery2 events *****************************************************************************
        //item changed in listBox
        if(methodName.equals("itemChangedInList")){
            if(((ItemEvent)eventObject).Selected == Controller.ORGANIGROUP){
                getController().setGroupType(Controller.ORGANIGROUP);
                getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
            }
            if(((ItemEvent)eventObject).Selected == Controller.RELATIONGROUP){
                getController().setGroupType(Controller.RELATIONGROUP);
                getController().setDiagramType(Controller.VENNDIAGRAM);
            }
            if(((ItemEvent)eventObject).Selected == Controller.PROCESSGROUP){
                getController().setGroupType(Controller.PROCESSGROUP);
                getController().setDiagramType(Controller.CONTINUOUSBLOCKPROCESS);
            }
            if(((ItemEvent)eventObject).Selected == Controller.LISTGROUP){
                getController().setGroupType(Controller.LISTGROUP);
                getController().setDiagramType(Controller.NOTDIAGRAM);
            }
            if(((ItemEvent)eventObject).Selected == 4){
                getController().setGroupType(Controller.MATRIXGROUP);
                getController().setDiagramType(Controller.NOTDIAGRAM);
            }
            getGui().setGalleryDialog2Images();
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 1. imageButton
        if(methodName.equals("item0Pressed") || methodName.equals("item4Pressed")|| methodName.equals("item8Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.VENNDIAGRAM);
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                getController().setDiagramType(Controller.CONTINUOUSBLOCKPROCESS);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 2. imageButton
        if(methodName.equals("item1Pressed") || methodName.equals("item5Pressed") || methodName.equals("item9Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.HORIZONTALORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.CYCLEDIAGRAM);
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                getController().setDiagramType(Controller.STAGGEREDPROCESS);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 3. imageButton
        if(methodName.equals("item2Pressed") || methodName.equals("item6Pressed") || methodName.equals("item10Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.TABLEHIERARCHYDIAGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.PYRAMIDDIAGRAM);
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                getController().setDiagramType(Controller.BENDINGPROCESS);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 4. imageButton
        if(methodName.equals("item3Pressed") || methodName.equals("item7Pressed") || methodName.equals("item11Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.ORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.TARGETDIAGRAM);
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                getController().setDiagramType(Controller.UPWARDARROWPROCESS);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }
        //****************************************************************************************************


        //ControlDialog1, ControlDialog2 common events *******************************************************
        //pressed Add shape button
        if(methodName.equals("addShape")){
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                        getController().getDiagram().addShape();
                        getController().getDiagram().refreshDiagram();
                        if(getController().getDiagram().isColorSchemeMode())
                            orgChart.refreshSchemesColors();
                    }
                }else{
                    getController().getDiagram().addShape();
                    getController().getDiagram().refreshDiagram();
                }
                getController().addSelectionListener();
                getController().setTextFieldOfControlDialog();
            }
            return true;
        }

        //pressed Remove shape button
        if(methodName.equals("removeShape")){
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                        getController().getDiagram().removeShape();
                        getController().getDiagram().refreshDiagram();
                        if(getController().getDiagram().isColorSchemeMode())
                            orgChart.refreshSchemesColors();
                    }
                }else{
                    getController().getDiagram().removeShape();
                    getController().getDiagram().refreshDiagram();
                }
                getController().addSelectionListener();
                getController().setTextFieldOfControlDialog();
            }
            return true;
        }

        //show ConvertDialog
        if(methodName.equals("convert")){
            if(!getController().isAppropriateShapeIsSelected()){
                String title = getGui().getDialogPropertyValue("Strings2", "Strings2.ConvertError.Title");
                String message = getGui().getDialogPropertyValue("Strings2", "Strings2.ConvertError.Message");
                getGui().showMessageBox(title, message);
                return true;
            }else{
                getController().removeSelectionListener();
                getController().convert();
                getController().addSelectionListener();
            }
            return true;
        }

        //pressed properties button
        if(methodName.equals("setProperties")){
            getController().getDiagram().showPropertyDialog();
            return true;
        }
        
        //pressed export button
        if(methodName.equals("export")){
            getController().export();
            return true;
        }
/*        
        //open-close textField (down-up)
        if(methodName.equals("downUpAction")){
            getGui().moveTextFieldOfControlDialog();
            return true;
        }
*/
        //set text of shape if textField has modified
        if(methodName.equals("textFieldModified")){
            getGui().modifiedTextFieldOfControlDialog();
            return true;
        }
        
        //a shortcut key
        if(methodName.equals("allShape")){
            if(getController().getDiagram() != null){
                getController().removeSelectionListener();
                getController().getDiagram().setFocusGroupShape();
                getController().addSelectionListener();
                getController().setTextFieldOfControlDialog();
            }
            return true;
        }

        //p shortcut key
        if(methodName.equals("previous")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isDad()){
                             OrganizationChartTreeItem previousItem = treeItem.getDad().getPreviousSibling(treeItem);
                            if(previousItem != null)
                                getController().setSelectedShape(previousItem.getRectangleShape());
                        }
                    }
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        XShape xShape = ((RelationDiagram)getController().getDiagram()).getPreviousShape();
                        if(xShape != null){
                            getController().setSelectedShape(xShape);
                            setNextColorOnControlDialog();
                        }
                    }
                    if(getController().getGroupType() == Controller.PROCESSGROUP){
                        XShape xShape = ((ProcessDiagram)getController().getDiagram()).getPreviousShape();
                        if(xShape != null){
                            getController().setSelectedShape(xShape);
                            setNextColorOnControlDialog();
                        }
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //n shortcut key
        if(methodName.equals("next")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isFirstSibling())
                            getController().setSelectedShape(treeItem.getFirstSibling().getRectangleShape());
                    }
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        XShape xShape = ((RelationDiagram)getController().getDiagram()).getNextShape();
                        if(xShape != null){
                            getController().setSelectedShape(xShape);
                            setNextColorOnControlDialog();
                        }
                    }
                    if(getController().getGroupType() == Controller.PROCESSGROUP){
                        XShape xShape = ((ProcessDiagram)getController().getDiagram()).getNextShape();
                        if(xShape != null){
                            getController().setSelectedShape(xShape);
                            setNextColorOnControlDialog();
                        }
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //e shortcut key
        if(methodName.equals("edit")){
//            if(!getGui().isShownTextFieldOfControlDialog())
//                getGui().moveTextFieldOfControlDialog();
            if(getController().isOnlySimpleItemIsSelected())
                getGui().setFocusTextFieldOfControlDialog();
            return true;
        }
        
        //s shortcut key
        if(methodName.equals("size")){
            if(getController().getDiagram() != null) {
                if(getController().getGroupType() != Controller.RELATIONGROUP){
                    getController().removeSelectionListener();
                    getController().setNewSize();
                    getController().addSelectionListener();
                }
            }  
            return true;
        }
        //****************************************************************************************************


        //ControlDialog1 events ******************************************************************************
        //click on colorImageControl
        if(methodName.equals("showColorTable3")){
            getGui().executeColorTable3();
            return true;
        }
        
        //f shortcut key
        if(methodName.equals("first")){
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = null;
                    if(!orgChart.isHiddenRootElementProp())
                        treeItem = orgChart.getDiagramTree().getRootItem();
                    else
                        treeItem = orgChart.getDiagramTree().getRootItem().getFirstChild();
                    if(treeItem != null)
                        getController().setSelectedShape(treeItem.getRectangleShape());
                }
                if(getController().getGroupType() == Controller.RELATIONGROUP){
                    RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getFirstItem();
                    if(item != null){
                        getController().setSelectedShape(item.getMainShape());
                        setNextColorOnControlDialog();
                    }
                }
                if(getController().getGroupType() == Controller.PROCESSGROUP){
                    ProcessDiagramItem item = ((ProcessDiagram)getController().getDiagram()).getFirstItem();
                    if(item != null){
                        if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                            getController().setSelectedShape(item.getSecondShape());
                        else
                            getController().setSelectedShape(item.getMainShape());
                        setNextColorOnControlDialog();
                    }
                }
                getController().addSelectionListener();
                getController().setTextFieldOfControlDialog();
            }
            return true;
        }

        //m shortcut key
        if(methodName.equals("mainShape")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                        if(item != null){
                            getController().setSelectedShape(item.getMainShape());
                            setNextColorOnControlDialog();
                        }
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //t shortcut key
        if(methodName.equals("textShape")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                        if(item != null){
                            XShape xShape = item.getTextShape();
                            if(xShape != null && getController().getDiagram().isInGruopShapes(xShape))
                                getController().setSelectedShape(xShape);
                        }
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }
        //****************************************************************************************************


        //ControlDialog2 events ******************************************************************************
        //in Organigroup
        //click on colorImageControl
        if(methodName.equals("showColorTable2")){
            getGui().enableControlDialogWindow(false);
            if(getController().getDiagram().isAnyGradientColorMode())
                getGui().executeGradientDialog();
            else
                getGui().executeColorTable2();
            getGui().enableAndSetFocusControlDialog();
            return true;
        }
        
        //change the add-mode property of shape (underling-associate)
        if(methodName.equals("changedAddProp")){
            ((OrganizationChart)getController().getDiagram()).setNewItemHType((short)((ItemEvent)eventObject).Selected);
            return true;
        }

        //u shortcut key
        if(methodName.equals("up")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isDad())
                            if(!orgChart.isHiddenRootElementProp() || !orgChart.getDiagramTree().getRootItem().equals(treeItem.getDad()))
                                getController().setSelectedShape(treeItem.getDad().getRectangleShape());
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //d shortcut key
        if(methodName.equals("down")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isFirstChild())
                            getController().setSelectedShape(treeItem.getFirstChild().getRectangleShape());
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }
        //****************************************************************************************************


        //ColorTable dialogs common event ********************************************************************
        //click on some image
        if(methodName.contains("image")){
            getGui().setResultColorTable(methodName);
            if((getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP) && !getGui().isShownPropsDialog() && !getGui().isShownBaseColorsSettingsDialog()){
                getGui().setColorModeBasedColorTable3();
                getGui().endExecuteColorTable3();
            } else {
                getGui().endExecuteColorTable();
            }
            if(!getGui().isShownPropsDialog() && !getGui().isShownGradientDialog() && getController().getGroupType() == Controller.ORGANIGROUP)
                getController().getDiagram().setColorModeProp(Diagram.SIMPLE_COLOR_MODE);
            return true;
        }
        //****************************************************************************************************


        //ColorTable2 event **********************************************************************************
        //pressed GradientMode button
        if(methodName.equals("setGradientMode")){
            getGui().endExecuteColorTable();
            getGui().enableControlDialogWindow(false);
            getGui().executeGradientDialog();
            getGui().enableAndSetFocusControlDialog();
            return true;
        }
        //****************************************************************************************************


        //ColorTable3 event **********************************************************************************
        if(methodName.equals("setSimpleColorModeOptionButton")){
            getGui().enableSetBaseColorsSettingsCommandButtonOnColorTable3(false);
            getGui().enableImageControlsOfColorTable(true);
            return true;
        }
        
        if(methodName.equals("setBaseColorsModeOptionButton")){
            getGui().enableSetBaseColorsSettingsCommandButtonOnColorTable3(true);
            getGui().enableImageControlsOfColorTable(true);
            return true;
        }
        
        if(methodName.equals("setBaseColorsGradientsModeOptionButton")){
            getGui().enableSetBaseColorsSettingsCommandButtonOnColorTable3(true);
            getGui().enableImageControlsOfColorTable(true);
            return true;
        }
        
        if(methodName.equals("setColorThemeModeOptionButton")){
            getGui().enableSetBaseColorsSettingsCommandButtonOnColorTable3(false);
            getGui().enableImageControlsOfColorTable(false);
            return true;
        }
        
        if(methodName.equals("setColorSchemeModeOptionButton")){
            getGui().enableSetBaseColorsSettingsCommandButtonOnColorTable3(false);
            getGui().enableImageControlsOfColorTable(false);
            return true;
        }
        
        if(methodName.equals("setBaseColorsSettings")){
            getGui().showBaseColorsSettingsDialog();
            return true;
        }
        
        if(methodName.equals("saveColorModeAndExit")){
            getGui().setColorModeBasedColorTable3();
            getGui().endExecuteColorTable3();
            //getGui().enableAndSetFocusControlDialog();
            return true;
        }
        
        
        if(methodName.equals("colorTableHelp")){
            getGui().showColorTableHelpDialog();
            return true;
        }
     
// ha a style értéke séma vagy téma és színt választanak -> style = uerdefined legyen
        //****************************************************************************************************

        //Gradients dialog events ****************************************************************************
        //selected ColorMode option button
        if(methodName.equals("setColorMode")){
            getGui().enableGradientDialogControls(false);
            return true;
        }

        //selected GradientMode option button
        if(methodName.equals("setGradientSettings")){
            getGui().enableGradientDialogControls(true);
            return true;
        }

        //click on StartColor image control
        if(methodName.equals("setStartColor")){
            XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, ((EventObject)eventObject).Source);
            getGui().m_xEventObjectControl = xControl;
            getGui().executeColorTable();
            return true;
        }

        //click on EndColor image controle
        if(methodName.equals("setEndColor")){
            XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, ((EventObject)eventObject).Source);
            getGui().m_xEventObjectControl = xControl;
            getGui().executeColorTable();
            return true;
        }
        //****************************************************************************************************

        //ConvertDialog events *******************************************************************************
        //Default setting: 1. (or 2. option if DiagramType == Conroller.SIMPLEORGANIGRAM) option is selected
        //2. option has selected
        if(methodName.equals("converButton1Pressed")){
            getGui().enableListBoxOfConvertDialog(false);
            return true;
        }

        //3. option has selected
        if(methodName.equals("converButton2Pressed")){
            getGui().enableListBoxOfConvertDialog(true);
            return true;
        }
        //****************************************************************************************************
        
        //BaseColorsSettingsDialog events ********************************************************************
        if(methodName.equals("setDefaultBaseColorsSettings")){
            getGui().setDefaultBaseColors();
            return true;
        }
        if(methodName.equals("showColorTable")){
            XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, ((EventObject)eventObject).Source);
            getGui().m_xEventObjectControl = xControl;
            getGui().executeColorTable();
            return true;
        }
        //****************************************************************************************************
        
        
        return false;
    }

        // XTopWindowListener
    @Override
    public void windowClosing(EventObject event) {
        if(event.Source.equals(getGui().m_xControlDialogTopWindow)){
            getGui().setVisibleControlDialog(false);
        }
    }

    // XTopWindowListener
    @Override
    public void windowOpened(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void windowClosed(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void windowMinimized(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void windowNormalized(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void windowActivated(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void windowDeactivated(EventObject arg0) { }

    // XTopWindowListener
    @Override
    public void disposing(EventObject arg0) { }
    
    public void setNextColorOnControlDialog(){
        if(getController().getDiagram().isBaseColorsMode() || getController().getDiagram().isBaseColorsWithGradientMode()){
            if(getController().getGroupType() == Controller.RELATIONGROUP){
                RelationDiagram diagram = ((RelationDiagram)getController().getDiagram());
                diagram.setColorProp(diagram.getNextColor());
            }
            if(getController().getGroupType() == Controller.PROCESSGROUP){
                ProcessDiagram process = ((ProcessDiagram)getController().getDiagram());
                process.setColorProp(process.getNextColor());
            }
        }                
    }

}
