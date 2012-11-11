package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialog;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;

public class ListenerOfRelationgruopPropsDialogs extends ListenerOfPropsDialogs {

    ListenerOfRelationgruopPropsDialogs(GuiOfPropsDialogs gui, Controller controller){
       super(gui, controller);
    }

    @Override // XDialogEventHandler
    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[24];
        aMethods[0] = "allShape";
        aMethods[1] = "first";
        aMethods[2] = "previous";
        aMethods[3] = "next";
        aMethods[4] = "mainShape";
        aMethods[5] = "textShape";
        aMethods[6] = "size";
        aMethods[7] = "changedStyle";
        aMethods[8] = "modifyColors";
        aMethods[9] = "setAllShapes";
        aMethods[10] = "setSelectedShapes";
        aMethods[11] = "setBaseColors";
        aMethods[12] = "setBaseColorsSettings";
        aMethods[13] = "setColor";
        aMethods[14] = "showColorTable";
        aMethods[15] = "setPreDefinedColorTheme";
        aMethods[16] = "setPreDefinedColorScheme";
        aMethods[17] = "setOutlineNo";
        aMethods[18] = "setOutlineYes";
        aMethods[19] = "setFrameYes";
        aMethods[20] = "setFrameNo";
        aMethods[21] = "setTextToFitToSize";
        aMethods[22] = "setFontSize";
        aMethods[23] = "modifyTextColor";
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
                RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getFirstItem();
                if(item != null){
                    getController().setSelectedShape(item.getMainShape());
                    setNextColorOnControlDialog();
                }
                getController().addSelectionListener();
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }

        //p shortcut key
        if(methodName.equals("previous")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    XShape xShape = ((RelationDiagram)getController().getDiagram()).getPreviousShape();
                    if(xShape != null){
                        getController().setSelectedShape(xShape);
                        setNextColorOnControlDialog();
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
                    XShape xShape = ((RelationDiagram)getController().getDiagram()).getNextShape();
                    if(xShape != null){
                        getController().setSelectedShape(xShape);
                        setNextColorOnControlDialog();
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }
        
        //m shortcut key
        if(methodName.equals("mainShape")){
            if(getController().isOnlySimpleItemIsSelected()){
                if(getController().getDiagram() != null) {
                    getController().removeSelectionListener();
                    RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                    if(item != null){
                        getController().setSelectedShape(item.getMainShape());
                        setNextColorOnControlDialog();
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
                    RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                    if(item != null){
                        XShape xShape = item.getTextShape();
                        if(xShape != null && getController().getDiagram().isInGruopShapes(xShape))
                            getController().setSelectedShape(xShape);
                    }
                    getController().addSelectionListener();
                }
            }
            getController().setTextFieldOfControlDialog();
            return true;
        }
        
        //s shortcut key
        if(methodName.equals("size")){
            /*
            if(getController().getDiagram() != null) {
                getController().removeSelectionListener();
                getController().setNewSize();
                getController().addSelectionListener();
            }
            */
            return true;
        }
        

        //item changed in Style ListBox
        if(methodName.equals("changedStyle")){
            if( getController().getDiagram() != null ){
                short itemPos = (short)((ItemEvent)eventObject ).Selected;
                if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                    getGui().setVennDiagramPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                    getGui().setCycleDiagramPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                    getGui().setPyramidDiagramPropsDialog(itemPos);
                if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                    getGui().setTargetDiagramPropsDialog(itemPos);
            }
            return true;
        }

        //changed modifyColors check box
        if(methodName.equals("modifyColors")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                 getGui().enableVennDiagramColorControls(getGui().isStateOfPropsDialogModifyColorCheckBox());
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                 getGui().enableCycleDiagramColorControls(getGui().isStateOfPropsDialogModifyColorCheckBox());
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                 getGui().enablePyramidDiagramColorControls(getGui().isStateOfPropsDialogModifyColorCheckBox());
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                 getGui().enableTargetDiagramColorControls(getGui().isStateOfPropsDialogModifyColorCheckBox());
            return true;
        }

        //selected setAllShape option button
        if(methodName.equals("setAllShapes")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().setVennDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorRadioButtons();
            return true;
        }

        //selected setSelectedShapes option button
        if(methodName.equals("setSelectedShapes")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().setVennDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorRadioButtons();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorRadioButtons();
            return true;
        }

       
        //selected setColor option button
        if(methodName.equals("setBaseColors")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().setVennDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorControls();
            return true;
        }
        
        //push setBaseColorsSettings CommandButton
        if(methodName.equals("setBaseColorsSettings")){
            getGui().showBaseColorsSettingsDialog();
            return true;
        }
        
        //selected setColor option button
        if(methodName.equals("setColor")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().setVennDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorControls();
            return true;
        }

        //selected setPreDefinedColorThemes option button
        if(methodName.equals("setPreDefinedColorTheme")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().setVennDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorControls();
            return true;
        }
        
        //selected setPreDefinedColorSchemes option button
        if(methodName.equals("setPreDefinedColorScheme")){
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                getGui().setCycleDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                getGui().setPyramidDiagramPDColorControls();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                getGui().setTargetDiagramPDColorControls();
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

        //selected setFrameYes option button
        if(methodName.equals("setFrameYes")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().enableVennDiagramPDRoundedFrameControls(true);
            return true;
        }

        //selected setFrameNo option button
        if(methodName.equals("setFrameNo")){
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                getGui().enableVennDiagramPDRoundedFrameControls(false);
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

        //changed state of modifyTextColorCheckBox in every diagram
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
