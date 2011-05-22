package oxygenoffice.extensions.smart;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.EventObject;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XRadioButton;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.buttdiagram.TargetDiagram;


public class Listener implements  XDialogEventHandler, XTopWindowListener {


    private Gui             m_Gui               = null;
    private Controller      m_Controller        = null;
    private short           m_iLastHorLevel     = 2;


    Listener(Gui gui, Controller controller){
        m_Gui = gui;
        m_Controller = controller;
    }
    
    public Gui getGui(){
        return m_Gui;
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    // XDialogEventHandler
    @Override
    public String[] getSupportedMethodNames() {

        String[] aMethods = new String[163];

        //DiagramGallery1 events - deprecated *************************
        //aMethods[0] = "Organigram";
        //aMethods[1] = "VennDiagram";
        //aMethods[2] = "PyramidDiagram";
        //aMethods[3] = "CycleDiagram";
        //aMethods[4] = "OK";
        //*************************************************************

        //DiagramGallery2 events **************************************
        aMethods[0] = "itemChangedInList";
        aMethods[1] = "item0Pressed";
        aMethods[2] = "item1Pressed";
        aMethods[3] = "item2Pressed";
        aMethods[4] = "item3Pressed";
        aMethods[5] = "item4Pressed";
        aMethods[6] = "item5Pressed";
        aMethods[7] = "item6Pressed";
        aMethods[8] = "item7Pressed";
        aMethods[9] = "eventOK";
        //*************************************************************

        //ControlDialog1, ControlDialog2 common events ****************
        aMethods[10] = "addShape";
        aMethods[11] = "removeShape";
        aMethods[12] = "convert";
        aMethods[13] = "setProperties";
        aMethods[14] = "downUpAction";
        aMethods[15] = "textFieldModified";
        //short cuts - hidden buttons:
        aMethods[16] = "allShape";
        aMethods[17] = "previous";
        aMethods[18] = "next";
        aMethods[19] = "edit";
        //************************************************************

        //ControlDialog1 events ***************************************
        aMethods[20] = "showColorTable3";
        //"addShape" - common
        //"removeShape" - common
        //"convert" - common
        //"setProperties" - common
        //"downUpAction" - common
        //"textFieldModified" - common
        //short cuts - hidden buttons:
        //"allShape" - common
        aMethods[21] = "first";
        //"previous" - common
        //"next" - common
        aMethods[22] = "mainShape";
        aMethods[23] = "textShape";
        //"edit" - common
        //*************************************************************

        //ControlDialog2 events ***************************************
        aMethods[24] = "showColorTable2";
        aMethods[25] = "changedAddProp";
        //"addShape" - common
        //"removeShape" - common
        //"convert" - common
        //"setProperties" - common
        //"downUpAction" - common
        //"textFieldModified" - common
        //short cuts - hidden buttons:
        //"allShape" - common
//        aMethods[26] = "root";
        //"previous" - common
        //"next" - common
        aMethods[27] = "up";
        aMethods[28] = "down";
        //"edit" - common
        //*************************************************************

        //ColorTables common events ***********************************
        for(int i=1;i<=92;i++)
            aMethods[i+28] = "image" +i;
        //*************************************************************

        //ColorTable2 events ******************************************
        aMethods[121] = "setGradientMode";
        //*************************************************************

        //ColorTable3 events ******************************************
        //*************************************************************

        //common method of GradientDialog and PropsDialogs ************
        aMethods[122] = "showColorTable";
        aMethods[123] = "changedGradientMode";
        //*************************************************************

        //GradientDialog events ***************************************
        aMethods[124] = "setColorMode";
        aMethods[125] = "setGradientSettings";
        aMethods[126] = "setStartColor";
        aMethods[127] = "setEndColor";
        //"changedGradientMode" -common method w. OrganigramPropsDialog
        aMethods[128] = "gradientDialogOK";
        //*************************************************************

        //PropsDialogs common events***********************************
        aMethods[129] = "changedStyle";
        aMethods[130] = "setAllShapes";
        aMethods[131] = "setSelectedShapes";
        aMethods[132] = "modifyColors";
        aMethods[133] = "setColor";
        aMethods[134] = "setOutlineYes";
        aMethods[135] = "setOutlineNo";
        //*************************************************************

        //OrganigramPropsDialog events ********************************
        //"changedStyle" - common
        //"setAllShapes" - common
        //"setSelectedShapes"; - common
        aMethods[136] = "setSiblingsShapes";
        aMethods[137] = "setBranchShapes";
        //"modifyColors" - common
        //"setColor" - common
        //"showColorTable" - common (set simple color)
        aMethods[138] = "setGradients";
        //"changedGradientMode" - common with Gradient dialog
        //"showColorTable" - common (set start color)
        //"showColorTable" - common (set end color)
        aMethods[139] = "setNoRounded";
        aMethods[140] = "setMediumRounded";
        aMethods[141] = "setExtraRounded";
        aMethods[142] = "setShadowYes";
        aMethods[143] = "setShadowNo";
        //"setOutlineYes" - common
        //"setOutlineNo" - common
        aMethods[144] = "setOrganigramProps";
        //*************************************************************

        //VennDiagramPropsDialog events *******************************
        //"changedStyle" - common
        //"setAllShapes" - common
        //"setSelectedShapes" - common
        //"modifyColors" - common
        aMethods[145] = "setBaseColor";
        //"setColor" - common
        //"showColorTable" - common (set simple color)
        aMethods[146] = "setNoTransparency";
        aMethods[147] = "setMediumTransparency";
        aMethods[148] = "setExtraTransparency";
        //"setOutlineYes" - common
        //"setOutlineNo" - common
        aMethods[149] = "setFrameYes";
        aMethods[150] = "setFrameNo";
        aMethods[151] = "setFrameRoundedYes";
        aMethods[152] = "setFrameRoundedNo";
        aMethods[153] = "setVennDiagramProps";
        //*************************************************************

        //CycleDiagramPropsDialog events *******************************
        //"changedStyle" - common
        //"setAllShapes" - common
        //"setSelectedShapes" - common
        //"modifyColors" - common
        //"setBaseColor" - common with VennDiagramPropsDialog
        //"setColor" - common
        //"showColorTable" - common (set simple color)
        //"setShadowYes" - common with OrganigramPropsDialog
        //"setShadowNo" - common with OrganigramPropsDialog
        //"setOutlineYes" - common
        //"setOutlineNo" - common
        //"setFrameYes" - common with VennDiagramPropsDialog
        //"setFrameNo" - common with VennDiagramPropsDialog
        aMethods[154] = "setCycleDiagramProps";
        //*************************************************************

        //ConvertDiaglog events ***************************************
        aMethods[155] = "converButton1Pressed";
        aMethods[156] = "converButton2Pressed";
        aMethods[157] = "changedLastHorLevel";
        aMethods[158] = "convertAction";
        //*************************************************************


        aMethods[159] = "setTargetDiagramProps";
        aMethods[160] = "setPyramidDiagramProps";

        aMethods[161] = "saveBaseColorsCheckBoxSetting";
        aMethods[163] = "changedBaseColorsCheckBox";

        return aMethods;
    }

    // XDialogEventHandler
    @Override
    public boolean callHandlerMethod(XDialog xDialog, Object eventObject, String methodName) throws WrappedTargetException {

        //DiagramGallery2 events *****************************************************************************
        //item changed in listBox
        if(methodName.equals("itemChangedInList")){
            if(((ItemEvent)eventObject).Selected == 0){
                getController().setGroupType(Controller.ORGANIGROUP);
                getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
            }
            if(((ItemEvent)eventObject).Selected == 1){
                getController().setGroupType(Controller.RELATIONGROUP);
                getController().setDiagramType(Controller.VENNDIAGRAM);
            }
            if(((ItemEvent)eventObject).Selected == 2){
                getController().setGroupType(Controller.LISTGROUP);
                getController().setDiagramType(Controller.NOTDIAGRAM);
            }
            if(((ItemEvent)eventObject).Selected == 3){
                getController().setGroupType(Controller.PROCESSGROUP);
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
        if(methodName.equals("item0Pressed") || methodName.equals("item4Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.VENNDIAGRAM);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 2. imageButton
        if(methodName.equals("item1Pressed") || methodName.equals("item5Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.HORIZONTALORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.CYCLEDIAGRAM);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 3. imageButton
        if(methodName.equals("item2Pressed") || methodName.equals("item6Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.TABLEHIERARCHYDIAGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.PYRAMIDDIAGRAM);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed 4. imageButton
        if(methodName.equals("item3Pressed") || methodName.equals("item7Pressed")){
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().setDiagramType(Controller.ORGANIGRAM);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getController().setDiagramType(Controller.TARGETDIAGRAM);
            if(getController().getGroupType() == Controller.LISTGROUP)
                ;
            if(getController().getGroupType() == Controller.PROCESSGROUP)
                ;
            if(getController().getGroupType() == Controller.MATRIXGROUP)
                ;
            getGui().setGalleryDialogText();
            return true;
        }

        //pressed OK button ("OK" - pressed OK button of DiagramGallery1 dialog - deprecated)
        if(methodName.equals("OK") || methodName.equals("eventOK")){
            getGui().setShownTips();
            getGui().executeGalleryDialog(false);
            getController().instantiateDiagram();
            if(getController().getDiagram() != null){
                getGui().setVisibleControlDialog(true);
                getController().getDiagram().createDiagram();
                getController().getDiagram().initDiagram();
            }
            if(getGui().isShownTips()){
                String title = getGui().getDialogPropertyValue("Strings2", "Strings2.Tips.Name.Label");
                String message = "";
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    message = getGui().getDialogPropertyValue("Strings2", "Strings2.Organizationcharts.Tips.Label");
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    message = getGui().getDialogPropertyValue("Strings2", "Strings2.Relationdiagrams.Tips.Label");
                getGui().showMessageBox(title, message);
            }
     
            return true;
        }
        //****************************************************************************************************


        //ControlDialog1, ControlDialog2 common events *******************************************************
        //pressed Add shape button
        if(methodName.equals("addShape")){
            if(getController().getDiagram() != null) {
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                        getController().getDiagram().addShape();
                        getController().getDiagram().refreshDiagram();
                    }
                }else{
                    getController().getDiagram().addShape();
                    getController().getDiagram().refreshDiagram();
                }
            }
            return true;
        }

        //pressed Remove shape button
        if(methodName.equals("removeShape")){
            if(getController().getDiagram() != null) {
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                       getController().getDiagram().removeShape();
                       getController().getDiagram().refreshDiagram();
                    }
                }else{
                    getController().getDiagram().removeShape();
                    getController().getDiagram().refreshDiagram();
                }
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
            }
            if(getController().getGroupType() == Controller.ORGANIGROUP && getController().getDiagramType() != Controller.ORGANIGRAM)
                m_iLastHorLevel = 2;
            getGui().executeConvertDialog();
            return true;
        }

        //pressed properties button
        if(methodName.equals("setProperties")){
            getController().getDiagram().refreshShapeProperties();
            getController().getDiagram().refreshDiagram();
            return true;
        }
        
        //open-close textField (down-up)
        if(methodName.equals("downUpAction")){
            getGui().moveTextFieldOfControlDialog();
            return true;
        }

        //set text of shape if textField has modified
        if(methodName.equals("textFieldModified")){
            getGui().modifiedTextFieldOfControlDialog();
            return true;
        }
        
        //a shortcut key
        if(methodName.equals("allShape")){
            if(getController().getDiagram() != null){
                getController().getDiagram().setFocusGroupShape();
            }
            return true;
        }

        //p shortcut key
        if(methodName.equals("previous")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
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
                        if(xShape != null)
                            getController().setSelectedShape(xShape);
                    }
                }
            }
            return true;
        }

        //n shortcut key
        if(methodName.equals("next")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isFirstSibling())
                            getController().setSelectedShape(treeItem.getFirstSibling().getRectangleShape());
                    }
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        XShape xShape = ((RelationDiagram)getController().getDiagram()).getNextShape();
                        if(xShape != null)
                            getController().setSelectedShape(xShape);
                    }
                }
            }
            return true;
        }

        //e shortcut key
        if(methodName.equals("edit")){
            if(!getGui().isShownTextFieldOfControlDialog())
                getGui().moveTextFieldOfControlDialog();
            if(getController().isOnlySimpleItemIsSelected())
                getGui().setFocusTextFieldOfControlDialog();
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
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getRootItem();
                    if(treeItem != null)
                        getController().setSelectedShape(treeItem.getRectangleShape());
                }
                if(getController().getGroupType() == Controller.RELATIONGROUP){
                    RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getFirstItem();
                    if(item != null)
                        getController().setSelectedShape(item.getMainShape());
                }
            }
            return true;
        }

        //m shortcut key
        if(methodName.equals("mainShape")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                        if(item != null)
                            getController().setSelectedShape(item.getMainShape());
                    }
                }
            }
            return true;
        }

        //t shortcut key
        if(methodName.equals("textShape")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    if(getController().getGroupType() == Controller.RELATIONGROUP){
                        RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                        if(item != null){
                            XShape xShape = item.getTextShape();
                            if(xShape != null && getController().getDiagram().isInGruopShapes(xShape))
                                getController().setSelectedShape(xShape);
                        }
                    }
                }
            }
            return true;
        }
        //****************************************************************************************************


        //ControlDialog2 events ******************************************************************************
        //click on colorImageControl
        if(methodName.equals("showColorTable2")){
            if(getController().getDiagram().isGradientProps() || getController().getDiagram().isBlueGradientsProps() || getController().getDiagram().isRedGradientsProps())
                getGui().executeGradientDialog();
            else
                getGui().executeColorTable2();
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
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isDad())
                            getController().setSelectedShape(treeItem.getDad().getRectangleShape());
                    }
                }
            }
            return true;
        }

        //d shortcut key
        if(methodName.equals("down")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                        OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                        if(treeItem.isFirstChild())
                            getController().setSelectedShape(treeItem.getFirstChild().getRectangleShape());
                    }
                }
            }
            return true;
        }
        //****************************************************************************************************


        //ColorTable dialogs common event ********************************************************************
        //click on some image
        if(methodName.contains("image")){
            getGui().setResultColorTable(true, methodName);
            getGui().endExecuteColorTable();
            getGui().disposeColorTable();
            return true;
        }
        //****************************************************************************************************


        //ColorTable2 event **********************************************************************************
        //pressed GradientMode button
        if(methodName.equals("setGradientMode")){
            getController().getDiagram().setGradientProps(true);
            if(getController().getGroupType() == Controller.ORGANIGROUP){
                getController().getDiagram().setBlueGradientsProps(false);
                getController().getDiagram().setRedGradientsProps(false);
            }
            getGui().endExecuteColorTable();
            getGui().disposeColorTable();
            getGui().executeGradientDialog();
            return true;
        }
        //****************************************************************************************************


        //ColorTable3 event **********************************************************************************
        if(methodName.equals("changedBaseColorsCheckBox")){
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM || getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getDiagramType() == Controller.TARGETDIAGRAM){
                if(getController().getDiagram().isBlueGradientsProps() || getController().getDiagram().isRedGradientsProps()){
                    if(getGui().isBaseColorModeOfColorTable())
                        getGui().enableImageControlsOfColorTable(true);
                    else
                        getGui().enableImageControlsOfColorTable(false);
                }
            }
            return true;
        }
        if(methodName.equals("saveBaseColorsCheckBoxSetting")){
            boolean isGradients = false;
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM || getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getDiagramType() == Controller.TARGETDIAGRAM)
                if(getController().getDiagram().isBlueGradientsProps() || getController().getDiagram().isRedGradientsProps())
                   isGradients = true;
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getGui().setBaseColorsProps();
            if(isGradients && getController().getDiagram().isBaseColorsProps())
                getGui().setImageColorOfControlDialog(((RelationDiagram)getController().getDiagram()).getNextColor());
            getGui().enableControlDialogWindow(true);
            getGui().setFocusControlDialog();
            getGui().endExecuteColorTable();
            getGui().disposeColorTable();
            return true;
        }
        //****************************************************************************************************


        //GradientDialog and DiagramPropsDialogs common events ***********************************************
        //click on colorImageControl
        if(methodName.equals("showColorTable")){
            //set source when the DagramProps dialog's occured the call (3 options)
            if(((EventObject)eventObject).Source.equals(getGui().m_xColorImageControlOfPD))
                getGui().m_xEventObjectControl = getGui().m_xColorImageControlOfPD;
            if(((EventObject)eventObject).Source.equals(getGui().m_xStartColorImageControlOfPD))
                getGui().m_xEventObjectControl = getGui().m_xStartColorImageControlOfPD;
            if(((EventObject)eventObject).Source.equals(getGui().m_xEndColorImageControlOfPD))
                getGui().m_xEventObjectControl = getGui().m_xEndColorImageControlOfPD;
            getGui().executeColorTable();
            return true;
        }

        //changed gradientMode combo box
        if(methodName.contains("changedGradientMode")){
            getController().getDiagram().setGradientDirectionProps((short)((ItemEvent)eventObject).Selected);
            return true;
        }
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

        //click on StartColor image controle
        if(methodName.equals("setStartColor")){
            getGui().m_sImageType = "StartImage";
            getGui().executeColorTable();
            return true;
        }

        //click on EndColor image controle
        if(methodName.equals("setEndColor")){
            getGui().m_sImageType = "EndImage";
            getGui().executeColorTable();
            return true;
        }

        //pressed OK button
        if(methodName.equals("gradientDialogOK")){
            getGui().setGradientProps();
            if(getController().getDiagram().isGradientProps()){
                getController().getDiagram().setStartColorProps(getGui().getColorOfStartImageOfGradientDialog());
                getController().getDiagram().setEndColorProps(getGui().getColorOfEndImageOfGradientDialog());
            }
            getGui().disposeGradientDialog();
            getGui().setColorModeOfImageOfControlDialog();
            return true;
        }
        //****************************************************************************************************


        //DiagramProps dialogs common events *****************************************************************
        //item changed in Style ListBox
        if(methodName.equals("changedStyle")){
            if( getController().getDiagram() != null ){
                getController().getDiagram().setStyle((short)( (ItemEvent)eventObject ).Selected);
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    getGui().setOrganigramPropsDialog();
                if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                    getGui().setVennDiagramPropsDialog();
                if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                    getGui().setCycleDiagramPropsDialog();
                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                    getGui().setPyramidDiagramPropsDialog();
                if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                    getGui().setTargetDiagramPropsDialog();
            }
            return true;
        }

        //selected setAllShape option button
        if(methodName.contains("setAllShapes")){
            getController().getDiagram().setSelectedAllShapesProps(true);
//            if(getGui().m_xBaseColorRadioButton != null)
//                if(getGui().m_xBaseColorRadioButton.getState())
//                    getGui().m_xColorRadioButton.setState(true);
//            getGui().enableControl(getGui().m_xBaseColorControl, true);
//            getGui().enableControl(getGui().m_xYesOutlineOBControl, true);
//            getGui().enableControl(getGui().m_xNoOutlineOBControl, true);
//            getGui().enableControl(getGui().m_xYesFrameOBControl, true);
//            getGui().enableControl(getGui().m_xNoFrameOBControl, true);
//            getGui().enableControl(getGui().m_xNoFrameRoundedOBControl, true);
//            getGui().enableControl(getGui().m_xYesFrameRoundedOBControl, true);

            //if(getController().getGroupType() == Controller.ORGANIGROUP)
               
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().enableVennDiagramNotAllShapeControls(true);
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().enableTargetDiagramNotAllShapeControls(true);
            return true;
        }

        //selected setSelectedShapes option button
        if(methodName.contains("setSelectedShapes")){
            getController().getDiagram().setSelectedAllShapesProps(false);
//            if(getGui().m_xBaseColorRadioButton != null)
//                if(getGui().m_xBaseColorRadioButton.getState())
//                    getGui().m_xColorRadioButton.setState(true);
//            getGui().enableControl(getGui().m_xBaseColorControl, false);
//            getGui().enableControl(getGui().m_xYesOutlineOBControl, false);
//            getGui().enableControl(getGui().m_xNoOutlineOBControl, false);
//            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
//            getGui().enableControl(getGui().m_xYesFrameOBControl, false);
//            getGui().enableControl(getGui().m_xNoFrameOBControl, false);
//            getGui().enableControl(getGui().m_xNoFrameRoundedOBControl, false);
//            getGui().enableControl(getGui().m_xYesFrameRoundedOBControl, false);
            
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getController().getDiagram().setSelectedAreaProps(Diagram.SELECTED_SHAPES);
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().enableVennDiagramNotAllShapeControls(false);
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().enableTargetDiagramNotAllShapeControls(false);
            return true;
        }

        //changed modifyColors check box
        if(methodName.contains("modifyColors")){
            if(getController().getGroupType() == Controller.ORGANIGROUP){
                if(getGui().m_xModifyColorsCheckBoxOfPD.getState() == 0){
                    getController().getDiagram().setModifyColorsProps(false);
                    getGui().enableOrganigramColorControls(false);
                }else{
                    getController().getDiagram().setModifyColorsProps(true);
                    getGui().enableControl(getGui().m_xColorOBControlOfPD, true);
                    getGui().enableControl(getGui().m_xGradientsOBControlOfPD, true);
                    boolean state = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, getGui().m_xColorOBControlOfPD)).getState();
                    getGui().enableOrganigramControlsIsColorProps(state);
                }
            }
            if(getController().getGroupType() == Controller.RELATIONGROUP){
                if(getGui().m_xModifyColorsCheckBoxOfPD.getState() == 0){
                    getController().getDiagram().setModifyColorsProps(false);
                    getGui().enableControl(getGui().m_xBaseColorOBControlOfPD, false);
                    getGui().enableControl(getGui().m_xColorOBControlOfPD, false);
                    getGui().enableControl(getGui().m_xColorImageControlOfPD, false);
                }else{
                    getController().getDiagram().setModifyColorsProps(true);
                    getGui().enableControl(getGui().m_xBaseColorOBControlOfPD, true);
                    getGui().enableControl(getGui().m_xColorOBControlOfPD, true);
                    XRadioButton xRB = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, getGui().m_xColorOBControlOfPD);
                    getGui().enableControl(getGui().m_xColorImageControlOfPD, xRB.getState());
                }
            }
            return true;
        }

        //selected setColor option button
        if(methodName.contains("setColor")){
//            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
//            getGui().enableControl(getGui().m_xStartColorLabelControl, false);
//            getGui().enableControl(getGui().m_xEndColorLabelControl, false);
//            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, false);
//            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, false);
//            getGui().enableControl(getGui().m_xColorImageControl, true);

            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getGui().enableOrganigramControlsIsColorProps(true);
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getGui().enableControl(getGui().m_xColorImageControlOfPD, true);

            //getController().getDiagram().setGradientProps(false);
            //getController().getDiagram().setBaseColorsProps(false);
            return true;
        }

        //selected yesOutlin option button
        if(methodName.contains("setOutlineYes")){
            getController().getDiagram().setOutlineProps(true);
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                getGui().enableControl(getGui().m_xYesShadowOBOfPD, true);
                getGui().enableControl(getGui().m_xNoShadowOBOfPD, true);
            }
            return true;
        }

        //selected noOutline option button
        if(methodName.contains("setOutlineNo")){
            getController().getDiagram().setOutlineProps(false);
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                getController().getDiagram().setShadowProps(false);
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, getGui().m_xNoShadowOBOfPD)).setState(true);
                getGui().enableControl(getGui().m_xYesShadowOBOfPD, false);
                getGui().enableControl(getGui().m_xNoShadowOBOfPD, false);
            }
            return true;
        }
        //****************************************************************************************************


        //OrganigramPropsDialog events ***********************************************************************
        //selected selected shapes and its sibling shapes option button in selected area field
        if(methodName.equals("setSiblingsShapes")){
            getController().getDiagram().setSelectedAllShapesProps(false);
            getController().getDiagram().setSelectedAreaProps(Diagram.SIBLING_SHAPES);
            return true;
        }

        //selected selected shapes and its branches option button in selected area field
        if(methodName.equals("setBranchShapes")){
            getController().getDiagram().setSelectedAllShapesProps(false);
            getController().getDiagram().setSelectedAreaProps(Diagram.BRANCH_SHAPES);
            return true;
        }

        //selected setGradients option button
        if(methodName.contains("setGradients")){
//            getGui().enableVisibleControl(getGui().m_xColorImageControl, false);
//            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
//            getGui().enableControl(getGui().m_xStartColorLabelControl, true);
//            getGui().enableControl(getGui().m_xEndColorLabelControl, true);
//            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, true);
//            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, true);

            if(getController().getGroupType() == Controller.ORGANIGROUP)
                getGui().enableOrganigramControlsIsColorProps(false);

            //getController().getDiagram().setGradientProps(true);
//            getController().getDiagram().setBaseColorsProps(false);
            return true;
        }

        //selected setNoRounded option button
        if(methodName.contains("setNoRounded")){
            getController().getDiagram().setRoundedProps((short)0);
            return true;
        }

        //selected setMediumRounded option button
        if(methodName.contains("setMediumRounded")){
            getController().getDiagram().setRoundedProps((short)1);
            return true;
        }

        //selected setExtraRounded option button
        if(methodName.contains("setExtraRounded")){
            getController().getDiagram().setRoundedProps((short)2);
            return true;
        }

        //selected yesShadow option button
        if(methodName.contains("setShadowYes")){
            getController().getDiagram().setShadowProps(true);
            return true;
        }

        //selected noShadow option button
        if(methodName.contains("setShadowNo")){
            getController().getDiagram().setShadowProps(false);
            return true;
        }

        // OrganigramPropsDialog
        if(methodName.contains("setOrganigramProps")){
            getController().getDiagram().setActionProps(true);
            getGui().endExecutePropertiesDialog();
            return true;
        }
        //****************************************************************************************************


        //VennDiagramPropsDialog events **********************************************************************
        //selected setBaseColor option button
        if(methodName.endsWith("setBaseColor")){
//            getGui().enableControl(getGui().m_xStartColorLabelControl, false);
//            getGui().enableControl(getGui().m_xEndColorLabelControl, false);
//            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, false);
//            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, false);
//            getGui().enableControl(getGui().m_xColorImageControl, false);
//            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, true);

            if(getController().getGroupType() == Controller.RELATIONGROUP)
                getGui().enableControl(getGui().m_xColorImageControlOfPD, false);

            //getController().getDiagram().setBaseColorsProps(true);
//            getController().getDiagram().setGradientProps(false);
            return true;
        }

        //selected setNoTransparency option button
        if(methodName.contains("setNoTransparency")){
            getController().getDiagram().setTransparencyProps(Diagram.NULL_TRANSP);
            return true;
        }

        //selected setMediumTransparency option button
        if(methodName.contains("setMediumTransparency")){
            getController().getDiagram().setTransparencyProps(Diagram.MEDIUM_TRANSP);
            return true;
        }

        //selected setExtraTransparency option button
        if(methodName.contains("setExtraTransparency")){
            getController().getDiagram().setTransparencyProps(Diagram.EXTRA_TRANSP);
            return true;
        }
        
        //selected setFrameYes option button
        if(methodName.contains("setFrameYes")){
            getController().getDiagram().setFrameProps(true);
            if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                getGui().enableControl(getGui().m_xNoFrameRoundedOBControlOfPD, true);
                getGui().enableControl(getGui().m_xYesFrameRoundedOBControlOfPD, true);
            }
            return true;
        }

        //selected setFrameNo option button
        if(methodName.contains("setFrameNo")){
            getController().getDiagram().setFrameProps(false);
            if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                getGui().enableControl(getGui().m_xNoFrameRoundedOBControlOfPD, false);
                getGui().enableControl(getGui().m_xYesFrameRoundedOBControlOfPD, false);
            }
            return true;
        }

        //selected setFrameRoundedYes option button
        if(methodName.contains("setFrameRoundedYes")){
            getController().getDiagram().setRoundedFrameProps(true);
            return true;
        }

        //selected setFrameRoundedNo option button
        if(methodName.contains("setFrameRoundedNo")){
            getController().getDiagram().setRoundedFrameProps(false);
            return true;
        }

        //pressed OK button
        if(methodName.contains("setVennDiagramProps")){
            getController().getDiagram().setActionProps(true);
            getGui().endExecutePropertiesDialog();
            return true;
        }
        //****************************************************************************************************

        
        //CycleDiagramProps dialog event *********************************************************************
        //pressed OK button
        if(methodName.contains("setCycleDiagramProps")){
            getController().getDiagram().setActionProps(true);
            getGui().endExecutePropertiesDialog();
            return true;
        }
        //****************************************************************************************************


        //TargetDiagramProps dialog event *********************************************************************
        //pressed OK button
        if(methodName.contains("setTargetDiagramProps")){
            getController().getDiagram().setActionProps(true);
            if(getController().getGui().isLefLayoutInTargetDiagram())
                ((TargetDiagram)getController().getDiagram()).setLeftLayoutProperty(true);
            else
                ((TargetDiagram)getController().getDiagram()).setLeftLayoutProperty(false);
            getGui().endExecutePropertiesDialog();
            return true;
        }
        //****************************************************************************************************

        //PyramidDiagramProps dialog event *********************************************************************
        //pressed OK button
        if(methodName.contains("setPyramidDiagramProps")){
            getController().getDiagram().setActionProps(true);
            getGui().endExecutePropertiesDialog();
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

        //changed selected item in combo box
        if(methodName.equals("changedLastHorLevel")){
            m_iLastHorLevel = ((short)((ItemEvent)eventObject).Selected);
            return true;
        }

        //pressed OK button
        if(methodName.equals("convertAction")){
            short convType = getGui().getConversationType();
            getGui().endExecuteConvertDialog();
            getGui().disposeConvertDialog();

            if(getController().getDiagram() != null) {
                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                        if(convType == Controller.ORGANIGRAM){
                            if(getController().getDiagramType() == Controller.ORGANIGRAM && m_iLastHorLevel == OrgChartTree.LASTHORLEVEL)
                                return true;
                            OrgChartTree.LASTHORLEVEL = (short)m_iLastHorLevel;
                        }
                        if(convType != Controller.ORGANIGRAM || (convType == Controller.ORGANIGRAM && getController().getDiagramType() != Controller.ORGANIGRAM))
                            getController().convert(convType);
                        getController().getDiagram().refreshDiagram();
                        ((OrganizationChart)getController().getDiagram()).refreshConnectorProps();
                    }
                }
                if(getController().getGroupType() == Controller.RELATIONGROUP){
                    getController().convert(convType);
                    getController().getDiagram().refreshDiagram();
                }
            }
            return true;
        }
        //****************************************************************************************************


        //DiagramGallery1 events *****************************************************************************
        /*
        if(methodName.equals("Organigram")){
            getController().setDiagramType(Controller.ORGANIGRAM);
            getGui().setSelectDialogText();
            return true;
        }

        if(methodName.equals("VennDiagram")){
            getController().setDiagramType(Controller.VENNDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }

        if(methodName.equals("PyramidDiagram")){
            getController().setDiagramType(Controller.PYRAMIDDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }

        if(methodName.equals("CycleDiagram")){
            getController().setDiagramType(Controller.CYCLEDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }
        */
        //****************************************************************************************************

        return false;
    }

        // XTopWindowListener
    @Override
    public void windowClosing(EventObject event) {

        if(event.Source.equals(getGui().m_xGalleryDialogTopWindow)){
            getController().setLastDiagramName("");
            getGui().disposeGalleryDialog();
        }

        if(event.Source.equals(getGui().m_xControlDialogTopWindow))
            getGui().setVisibleControlDialog(false);

        if(event.Source.equals(getGui().m_xColorTableTopWindow)){
            getGui().setResultColorTable(false, "");
            getGui().disposeColorTable();
        }
        
        if(event.Source.equals(getGui().m_xGradientDialogTopWindow))
           getGui().disposeGradientDialog();

        if(event.Source.equals(getGui().m_xConvertDialogTopWindow))
            getGui().disposeConvertDialog();

        if(event.Source.equals(getGui().m_xPropsDialogTopWindow))
            getGui().endExecutePropertiesDialog();
        
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

}
