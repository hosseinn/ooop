package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialog;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;


public class ListenerOfProcessgroupPropsDialogs extends ListenerOfPropsDialogs {
    
     ListenerOfProcessgroupPropsDialogs(GuiOfPropsDialogs gui, Controller controller){
        super(gui, controller);
    }
     
     @Override // XDialogEventHandler
    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[23];

        aMethods[0] = "allShape";
        aMethods[1] = "first";
        aMethods[2] = "previous";
        aMethods[3] = "next";
        aMethods[4] = "size";
        aMethods[5] = "changedStyle";
        
        aMethods[6] = "modifyColors";
        aMethods[7] = "setAllShapes";
        aMethods[8] = "setSelectedShapes";
        aMethods[9] = "setBaseColors";
        aMethods[10] = "setBaseColorsSettings";
        aMethods[11] = "setColor";
        aMethods[12] = "showColorTable";
        aMethods[13] = "setPreDefinedColorTheme";
        aMethods[14] = "setPreDefinedColorScheme";
        aMethods[15] = "preDefinedColorThemesListBoxChanged";
        aMethods[16] = "preDefinedColorSchemesListBoxChanged";
        aMethods[17] = "modifyArrowShapesColor";
        
        aMethods[18] = "setOutlineNo";
        aMethods[19] = "setOutlineYes";
        aMethods[20] = "setTextToFitToSize";
        aMethods[21] = "setFontSize";
        aMethods[22] = "modifyTextColor";
        //**************************************************
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
                ProcessDiagramItem item = ((ProcessDiagram)getController().getDiagram()).getFirstItem();
                if(item != null){
                    if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                        getController().setSelectedShape(item.getSecondShape());
                    else
                        getController().setSelectedShape(item.getMainShape());
                    setNextColorOnControlDialog();
                }
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
                    XShape xShape = ((ProcessDiagram)getController().getDiagram()).getPreviousShape();
                    if(xShape != null){
                        getController().setSelectedShape(xShape);
                        setNextColorOnControlDialog();
                    }
                    getController().addSelectionListener();
                    getController().setTextFieldOfControlDialog();
                }
            }
            return true;
        }

        //n shortcut key
        if(methodName.equals("next")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    XShape xShape = ((ProcessDiagram)getController().getDiagram()).getNextShape();
                    if(xShape != null){
                        getController().setSelectedShape(xShape);
                        setNextColorOnControlDialog();
                    }
                    getController().addSelectionListener();
                    getController().setTextFieldOfControlDialog();
                }
            }
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
                if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                    getGui().setContinuousBlockProcessPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                    getGui().setStaggeredProcessPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                    getGui().setBendingProcessPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                    getGui().setUpwardArrowProcessPropsDialog(itemPos);
            }
            return true;
        }

        //changed modifyColors check box
        if(methodName.equals("modifyColors")){
            boolean state = getGui().isStateOfPropsDialogModifyColorCheckBox();
            if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS){
                getGui().enableContinuousBlockProcessColorControls(state);
                if(state)
                    getGui().setArrowColorControls();
            }
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                getGui().enableStaggeredProcessColorControls(state);
            if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                getGui().enableBendingProcessColorControls(state);
            if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS){
                getGui().enableUpwardArrowProcessColorControls(state);
                if(state)
                    getGui().setArrowColorControls();
            }
            
            return true;
        }

        //selected setAllShape option button
        //selected setSelectedShapes option button
        if(methodName.equals("setAllShapes") || methodName.equals("setSelectedShapes")){
            if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                getGui().setContinuousBlockProcessPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                getGui().setStaggeredProcessPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                getGui().setBendingProcessPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                getGui().setUpwardArrowProcessPDColorRadioButtons();
            return true;
        }

        //selected setColor option button
        //selected setPreDefinedColorThemes option button
        //selected setPreDefinedColorSchemes option button
        //changed state of preDefinedColorThemesListBox
        //changed state of preDefinedColorSchemesListBox
        if(methodName.equals("setBaseColors") || methodName.equals("setColor") || methodName.equals("setPreDefinedColorTheme") || methodName.equals("setPreDefinedColorScheme") || methodName.equals("preDefinedColorThemesListBoxChanged") || methodName.equals("preDefinedColorSchemesListBoxChanged")){
            if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                getGui().setContinuousBlockProcessPDColorControls();
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                getGui().setStaggeredProcessPDColorControls();
            if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                getGui().setBendingProcessPDColorControls();
            if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                getGui().setUpwardArrowProcessPDColorControls();
            return true;
        }
        
        //push setBaseColorsSettings CommandButton
        if(methodName.equals("setBaseColorsSettings")){
            getGui().showBaseColorsSettingsDialog();
            return true;
        }
      
        //push modifyArrowColor CheckBox
        if(methodName.equals("modifyArrowShapesColor")){
            getGui().enableDisableArrowColorImageControl();
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
                
        //selected textFit option button in every diagrams
        if(methodName.equals("setTextToFitToSize")){
            getGui().enableFontSizeListBox(false);
            return true;
        }

        //selected fontSize option button in every diagrams
        if(methodName.equals("setFontSize")){
            getGui().enableFontSizeListBox(true);
            return true;
        }

        //changed state of modifyTextColorCheckBox in every diagrams
        if(methodName.equals("modifyTextColor")){
            getGui().setTextColorToolsProps();
            return true;
        }
        
        //click on colorImageControl
        if(methodName.equals("showColorTable")){
            XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, ((EventObject)eventObject).Source);
            getGui().m_xEventObjectControl = xControl;
            getGui().executeColorTable();
            return true;
        }
        
        return false;
    }
}
