package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialog;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;

public class ListenerOfOrganigroupPropsDialogs extends ListenerOfPropsDialogs {

    ListenerOfOrganigroupPropsDialogs(GuiOfPropsDialogs gui, Controller controller){
       super(gui, controller);
    }

    @Override // XDialogEventHandler
    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[23];
        aMethods[0] = "allShape";
        aMethods[1] = "first";
        aMethods[2] = "previous";
        aMethods[3] = "next";
        aMethods[4] = "up";
        aMethods[5] = "down";
        aMethods[6] = "size";
        aMethods[7] = "changedStyle";
        aMethods[8] = "modifyColors";
        aMethods[9] = "setAllShapes";
        aMethods[10] = "setSelectedShapes";
        aMethods[11] = "setSiblingsShapes";
        aMethods[12] = "setBranchShapes";
        aMethods[13] = "setColor";
        aMethods[14] = "showColorTable";
        aMethods[15] = "setPreDefinedColorThemes";
        aMethods[16] = "setPreDefinedColorSchemes";
        aMethods[17] = "setGradients";
        aMethods[18] = "setOutlineYes";
        aMethods[19] = "setOutlineNo";
        aMethods[20] = "setTextToFitToSize";
        aMethods[21] = "setFontSize";
        aMethods[22] = "modifyTextColor";
        //***************************************************
        return aMethods;
    }

    @Override // XDialogEventHandler
    public boolean callHandlerMethod(XDialog xDialog, Object eventObject, String methodName) throws WrappedTargetException {

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

        //f shortcut key
        if(methodName.equals("first")){
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                OrganizationChartTreeItem treeItem = null;
                if(!orgChart.isHiddenRootElementProp())
                    treeItem = orgChart.getDiagramTree().getRootItem();
                else
                    treeItem = orgChart.getDiagramTree().getRootItem().getFirstChild();
                if(treeItem != null)
                    getController().setSelectedShape(treeItem.getRectangleShape());
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
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                    if(treeItem.isDad()){
                         OrganizationChartTreeItem previousItem = treeItem.getDad().getPreviousSibling(treeItem);
                        if(previousItem != null)
                            getController().setSelectedShape(previousItem.getRectangleShape());
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
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                    if(treeItem.isFirstSibling())
                        getController().setSelectedShape(treeItem.getFirstSibling().getRectangleShape());
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //u shortcut key
        if(methodName.equals("up")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                    if(treeItem.isDad())
                        if(!orgChart.isHiddenRootElementProp() || !orgChart.getDiagramTree().getRootItem().equals(treeItem.getDad()))
                            getController().setSelectedShape(treeItem.getDad().getRectangleShape());
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
                    OrganizationChart orgChart = (OrganizationChart)getController().getDiagram();
                    OrganizationChartTreeItem treeItem = orgChart.getDiagramTree().getTreeItem(getController().getSelectedShape());
                    if(treeItem.isFirstChild())
                        getController().setSelectedShape(treeItem.getFirstChild().getRectangleShape());
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }
        
        //s shortcut key
        if(methodName.equals("size")){
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                getController().setNewSize();
                getController().addSelectionListener();
            }  
            return true;
        }

        //item changed in Style ListBox
        if(methodName.equals("changedStyle")){
            if( getController().getDiagram() != null ){
                short itemPos = (short)((ItemEvent)eventObject ).Selected;
                getGui().setOrganigramPropsDialog(itemPos);
            }
            return true;
        }

        //changed modifyColors check box
        if(methodName.equals("modifyColors")){
            getGui().enableOrganigramColorControls(getGui().isStateOfPropsDialogModifyColorCheckBox());
            return true;
        }

        //selected setAllShape option button
        if(methodName.equals("setAllShapes")){
            getGui().setOrganigramPDColorRadioButtons();
            return true;
        }

        //selected setSelectedShapes option button
        if(methodName.equals("setSelectedShapes")){
            getGui().setOrganigramPDColorRadioButtons();
            return true;
        }

        //selected selected shapes and its sibling shapes option button in selected area field
        if(methodName.equals("setSiblingsShapes")){
            getGui().setOrganigramPDColorRadioButtons();
            return true;
        }

        //selected selected shapes and its branches option button in selected area field
        if(methodName.equals("setBranchShapes")){
            getGui().setOrganigramPDColorRadioButtons();
            return true;
        }

        //selected setColor option button
        if(methodName.equals("setColor")){
            getGui().setOrganigramPDColorControlsInColorMode();
            return true;
        }

        //selected setPreDefinedColorThemes option button
        if(methodName.equals("setPreDefinedColorThemes")){
            getGui().setOrganigramPDColorControlsInPreDefinedColorThemeGradinetMode();
            return true;
        }
        
        //selected setPreDefinedColorSchemes option button
        if(methodName.equals("setPreDefinedColorSchemes")){
            getGui().setOrganigramPDColorControlsInPreDefinedSchemeMode();
            return true;
        }

        //selected setGradients option button
        if(methodName.equals("setGradients")){
            getGui().setOrganigramPDColorControlsInGradientMode();
            return true;
        }

        //selected yesOutline option button
        if(methodName.equals("setOutlineYes")){
            getGui().setPropsDialogLineWidthControls(true);
            return true;
        }

        //selected noOutline option button
        if(methodName.equals("setOutlineNo")){
            getGui().setPropsDialogLineWidthControls(false);
            return true;
        }

        //selected textFit option button
        if(methodName.equals("setTextToFitToSize")){
            getGui().enableFontSizeListBox(false);
            return true;
        }

        //selected fontSize option button
        if(methodName.equals("setFontSize")){
            getGui().enableFontSizeListBox(true);
            return true;
        }

        //changed state of modifyTextColorCheckBox
        if(methodName.equals("modifyTextColor")){
            getGui().setTextColorToolsProps();
            return true;
        }

        //click on colorImageControl
        if(methodName.equals("showColorTable")){
            //set source
            XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, ((EventObject)eventObject).Source);
            getGui().m_xEventObjectControl = xControl;
            getGui().executeColorTable();
            return true;
        }

        return false;
    }
}
