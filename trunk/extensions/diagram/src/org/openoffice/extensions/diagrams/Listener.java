package org.openoffice.extensions.diagrams;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.EventObject;
import com.sun.star.awt.XDialogEventHandler;
import org.openoffice.extensions.diagrams.diagram.organizationdiagram.OrganizationDiagram;



public class Listener implements  XDialogEventHandler, XTopWindowListener {

    Gui         m_Gui           = null;
    Controller m_Controller   = null;

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
        String[] aMethods = new String[116];
        aMethods[0] = "Organigram";
        aMethods[1] = "VennDiagram";
        aMethods[2] = "PyramidDiagram";
        aMethods[3] = "CycleDiagram";
        aMethods[4] = "Ok";
        aMethods[5] = "showColorTable";
        aMethods[6] = "addShape";
        aMethods[7] = "removeShape";
        aMethods[8] = "changedStyle";
        aMethods[9] = "changedAddProp";
        aMethods[10] = "setStartColor";
        aMethods[11] = "setEndColor";
        aMethods[12] = "setColors";

        aMethods[13] = "setAllShapes";
        aMethods[14] = "setSelectedShapes";
        aMethods[15] = "setColor";
        aMethods[16] = "setGradients";
        aMethods[17] = "setNoRounded";
        aMethods[18] = "setMediumRounded";
        aMethods[19] = "setExtraRounded";
        aMethods[20] = "setMonographicYes";
        aMethods[21] = "setMonographicNo";
        aMethods[22] = "setOrganigramProps";
        aMethods[23] = "setBaseColor";

        for(int i=1;i<=92;i++)
            aMethods[i+23] = "image" +i;
        return aMethods;
    }

     // XDialogEventHandler
    @Override
    public boolean callHandlerMethod(XDialog xDialog, Object eventObject, String methodName) throws WrappedTargetException {
        // SelectDialog
        if(methodName.equals("Organigram")){
            getController().setDiagramType(Controller.ORGANIGRAM);
            getGui().setSelectDialogText();
            return true;
        }
        // SelectDialog
        if(methodName.equals("VennDiagram")){
            getController().setDiagramType(Controller.VENNDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }
        // SelectDialog
        if(methodName.equals("PyramidDiagram")){
            getController().setDiagramType(Controller.PYRAMIDDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }
        // SelectDialog
        if(methodName.equals("CycleDiagram")){
            getController().setDiagramType(Controller.CYCLEDIAGRAM);
            getGui().setSelectDialogText();
            return true;
        }
        // SelectDialog
        if(methodName.equals("Ok")){
            getGui().setVisibleSelectWindow(false);
            getController().instantiateDiagram();
            if(getController().getDiagram() != null){
                getGui().setVisibleControlDialog(true);
                getController().getDiagram().createDiagram();
                getController().getDiagram().initDiagram();
            }
            return true;
        }
        // ControlDialog1, ControlDialog2
        if(methodName.equals("showColorTable")){

            if(((EventObject)eventObject).Source.equals(getGui().m_xColorImageControl))
                getGui().m_xEventObjectControl = getGui().m_xColorImageControl;
            if(((EventObject)eventObject).Source.equals(getGui().m_xStartColorImageControl))
                getGui().m_xEventObjectControl = getGui().m_xStartColorImageControl;
            if(((EventObject)eventObject).Source.equals(getGui().m_xEndColorImageControl))
                getGui().m_xEventObjectControl = getGui().m_xEndColorImageControl;

            getGui().showColorTable();
            return true;
        }
        // ControlDialog1, ControlDialog2
        if(methodName.equals("addShape")){
            if(getController().getDiagram() != null) {
                if(getController().getDiagramType() == Controller.ORGANIGRAM){
                    OrganizationDiagram orgDiagram = (OrganizationDiagram)getController().getDiagram();
                    if(orgDiagram.searchErrorInTree()){
                        getGui().askUserForRepair(orgDiagram);
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
        // ControlDialog1, ControlDialog2
        if(methodName.equals("removeShape")){
            if(getController().getDiagram() != null) {
                if(getController().getDiagramType() == Controller.ORGANIGRAM){
                    OrganizationDiagram orgDiagram = (OrganizationDiagram)getController().getDiagram();
                    if(orgDiagram.searchErrorInTree()){
                        getGui().askUserForRepair(orgDiagram);
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
        // ControlDialog1, ControlDialog2
        if(methodName.equals("changedStyle")){
            //if( getController().getDrawPage() != null && getController().getShapes() != null){
            if( getController().getDiagram() != null ){
                getController().getDiagram().setChangedMode((short)( (ItemEvent)eventObject ).Selected);
                getController().getDiagram().refreshShapeProperties();
                getController().getDiagram().refreshDiagram();
            }
            return true;
        }
        // ControlDialog2
        if(methodName.equals("changedAddProp")){
            ((OrganizationDiagram)getController().getDiagram()).setNewItemHType((short)((ItemEvent)eventObject).Selected);
            return true;
        }
        if(methodName.equals("setStartColor")){
            getGui().m_sImageType = "StartImage";
            getGui().showColorTable();
            return true;
        }
        if(methodName.equals("setEndColor")){
            getGui().m_sImageType = "EndImage";
            getGui().showColorTable();
            return true;
        }
        if(methodName.equals("setColors")){
            getGui().m_xGradientTopWindow.removeTopWindowListener(this);
            ((OrganizationDiagram)getController().getDiagram()).setGradientAction(true);
            OrganizationDiagram._startColor = getGui().getImageColorOfControlDialog(getGui().m_xStartImage);
            OrganizationDiagram._endColor = getGui().getImageColorOfControlDialog(getGui().m_xEndImage);
            getGui().m_xGradientDialog.endExecute();
            getGui().m_xGradientWindow = null;
            getGui().m_xGradientTopWindow = null;
            getGui().m_xGradientDialog = null;
            getGui().m_xControlDialogWindow.setEnable(true);
            getGui().m_xControlDialogWindow.setFocus();
            return true;
        }
        // ColorTable
        if(methodName.contains("image")){

            int color = getGui().getCurrImageColor(xDialog, getGui().getNum(methodName));

            getGui().m_xPaletteDialog.endExecute();
            getGui().m_xPaletteTopWindow = null;
            getGui().m_xPaletteDialog = null;
            if(getGui().m_xGradientWindow != null){
                if(color != -1){
                    if(getGui().m_sImageType.equals("StartImage"))
                        getGui().setImageColorOfImageControl(getGui().m_xStartImage, color);
                    if(getGui().m_sImageType.equals("EndImage"))
                        getGui().setImageColorOfImageControl(getGui().m_xEndImage, color);
                }
                getGui().m_xGradientWindow.setFocus();
            }else{
                if(getGui().m_xPropsDialog != null){
                    if(getGui().m_xEventObjectControl != null)
                        if(color != -1)
                            getGui().setImageColorOfImageControl(getGui().m_xEventObjectControl, color);
                    getGui().m_xEventObjectControl = null;
                }else{
                    if(color != -1)
                        getGui().setImageColorOfImageControl(getGui().m_xImageControl, color);
                    getGui().m_xControlDialogWindow.setEnable(true);
                    getGui().m_xControlDialogWindow.setFocus();
                }
            }
            return true;
        }


        // DiagramPropsDialog
        if(methodName.contains("setAllShapes")){
            if(getGui().m_xBaseColorRadioButton != null)
                if(getGui().m_xBaseColorRadioButton.getState())
                    getGui().m_xColorRadioButton.setState(true);
            getGui().enableControl(getGui().m_xBaseColorControl, true);
            getGui().enableControl(getGui().m_xMonographicOBYesControl, true);
            getGui().enableControl(getGui().m_xMonographicOBNoControl, true);
            getGui().enableControl(getGui().m_xFrameOBYesControl, true);
            getGui().enableControl(getGui().m_xFrameOBNoControl, true);
            getGui().enableControl(getGui().m_xFrameRoundedOBNoControl, true);
            getGui().enableControl(getGui().m_xFrameRoundedOBYesControl, true);
            if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                if(getGui().m_xBaseColorRadioButton != null)
                    getGui().m_xColorRadioButton.setState(getGui().m_xColorRadioButton.getState());
                getGui().enableControl(getGui().m_xColorCBControl, false);
                getGui().enableVisibleControl(getGui().m_xColorImageControl, false);
            }

            getController().getDiagram().setSelectedDiagramProps(true);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.contains("setSelectedShapes")){
            if(getGui().m_xBaseColorRadioButton != null)
                if(getGui().m_xBaseColorRadioButton.getState())
                    getGui().m_xColorRadioButton.setState(true);
            getGui().enableControl(getGui().m_xBaseColorControl, false);
            getGui().enableControl(getGui().m_xMonographicOBYesControl, false);
            getGui().enableControl(getGui().m_xMonographicOBNoControl, false);
            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
            getGui().enableControl(getGui().m_xFrameOBYesControl, false);
            getGui().enableControl(getGui().m_xFrameOBNoControl, false);
            getGui().enableControl(getGui().m_xFrameRoundedOBNoControl, false);
            getGui().enableControl(getGui().m_xFrameRoundedOBYesControl, false);
            if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                getGui().enableControl(getGui().m_xColorCBControl, true);
                getGui().enableVisibleControl(getGui().m_xColorImageControl, true);
            }



            getController().getDiagram().setSelectedDiagramProps(false);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.contains("setColor")){
            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
            getGui().enableControl(getGui().m_xStartColorLabelControl, false);
            getGui().enableControl(getGui().m_xEndColorLabelControl, false);
            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, false);
            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, false);
            getGui().enableVisibleControl(getGui().m_xColorImageControl, true);

            getController().getDiagram().setGradientProps(false);
            getController().getDiagram().setBaseColorsProps(false);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.contains("setGradients")){
            getGui().enableVisibleControl(getGui().m_xColorImageControl, false);
            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, false);
            getGui().enableControl(getGui().m_xStartColorLabelControl, true);
            getGui().enableControl(getGui().m_xEndColorLabelControl, true);
            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, true);
            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, true);

            getController().getDiagram().setGradientProps(true);
            getController().getDiagram().setBaseColorsProps(false);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.endsWith("setBaseColor")){
            getGui().enableControl(getGui().m_xStartColorLabelControl, false);
            getGui().enableControl(getGui().m_xEndColorLabelControl, false);
            getGui().enableVisibleControl(getGui().m_xStartColorImageControl, false);
            getGui().enableVisibleControl(getGui().m_xEndColorImageControl, false);
            getGui().enableVisibleControl(getGui().m_xColorImageControl, false);
            getGui().enableControl(getGui().m_xGradientsCheckBoxControl, true);

            getController().getDiagram().setBaseColorsProps(true);
            getController().getDiagram().setGradientProps(false);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setNoRounded")){
            getController().getDiagram().setRoundedProps((short)0);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setMediumRounded")){
            getController().getDiagram().setRoundedProps((short)1);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setExtraRounded")){
            getController().getDiagram().setRoundedProps((short)2);
            return true;
        }


        // DiagramPropsDialog
        if(methodName.contains("setNoTransparency")){
            getController().getDiagram().setTransparencyProps((short)0);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setMediumTransparency")){
            getController().getDiagram().setTransparencyProps((short)1);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setExtraTransparency")){
            getController().getDiagram().setTransparencyProps((short)2);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.contains("setMonographicYes")){
            getController().getDiagram().setMonographicProps(true);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setMonographicNo")){
            getController().getDiagram().setMonographicProps(false);
            return true;
        }

        // DiagramPropsDialog
        if(methodName.contains("setFrameYes")){
            getGui().enableControl(getGui().m_xFrameRoundedOBNoControl, true);
            getGui().enableControl(getGui().m_xFrameRoundedOBYesControl, true);

            getController().getDiagram().setFrameProps(true);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setFrameNo")){
            getGui().enableControl(getGui().m_xFrameRoundedOBNoControl, false);
            getGui().enableControl(getGui().m_xFrameRoundedOBYesControl, false);

            getController().getDiagram().setFrameProps(false);
            return true;
        }


        // DiagramPropsDialog
        if(methodName.contains("setFrameRoundedYes")){
            getController().getDiagram().setRoundedFrameProps(true);
            return true;
        }
        // DiagramPropsDialog
        if(methodName.contains("setFrameRoundedNo")){
            getController().getDiagram().setRoundedFrameProps(true);
            return true;
        }



        // DiagramPropsDialog
        if(methodName.contains("setOrganigramProps")){

            getController().getDiagram().setActionProps(true);
            if(getGui().m_xGradientsCheckBox != null){
                boolean bool = false;
                if(getGui().m_xGradientsCheckBox.getState() == 1)
                    bool = true;
                getController().getDiagram().setBaseColorsWithGradientsProps(bool);
            }

            if(getGui().m_xColorCheckBox != null){
                boolean bool = false;
                if(getGui().m_xColorCheckBox.getState() == 1)
                    bool = true;
                getController().getDiagram().setColorProps(bool);
            }

            getController().getDiagram().setColorProps(getGui().getImageColorOfControlDialog(getGui().m_xColorImageControl));
            getController().getDiagram().setStartColorProps(getGui().getImageColorOfControlDialog(getGui().m_xStartColorImageControl));
            getController().getDiagram().setEndColorProps(getGui().getImageColorOfControlDialog(getGui().m_xEndColorImageControl));

            getGui().m_xPropsTopWindow.removeTopWindowListener(this);
            getGui().m_xPropsDialog.endExecute();
            getGui().m_xPropsTopWindow = null;
            getGui().m_xPropsDialog = null;

            getGui().m_xControlDialogWindow.setEnable(true);
            getGui().m_xControlDialogWindow.setFocus();

            return true;
        }


        return false;
    }

        // XTopWindowListener
    @Override
    public void windowClosing(EventObject event) {

        if(event.Source.equals(getGui().m_xSelectDTopWindow)){
            getController().setLastDiagramName("");
            getGui().m_xSelectDTopWindow.removeTopWindowListener(this);
        }
        if(event.Source.equals(getGui().m_xControlDialogWindow))
            getGui().setVisibleControlDialog(false);

        if(event.Source.equals(getGui().m_xPaletteTopWindow)){
            getGui().m_xPaletteTopWindow.removeTopWindowListener(this);
            getGui().m_xPaletteDialog.endExecute();
            getGui().m_xPaletteTopWindow = null;
            getGui().m_xPaletteDialog = null;
            if(getGui().m_xGradientWindow != null || getGui().m_xPropsTopWindow != null){
                if(getGui().m_xGradientWindow != null)
                    getGui().m_xGradientWindow.setFocus();

            }else{
                getGui().m_xControlDialogWindow.setEnable(true);
                getGui().m_xControlDialogWindow.setFocus();
            }
        }
        if(event.Source.equals(getGui().m_xGradientTopWindow)){
           getGui().m_xGradientTopWindow.removeTopWindowListener(this);
           getGui().m_xGradientDialog.endExecute();
           getGui().m_xGradientWindow = null;
           getGui().m_xGradientTopWindow = null;
           getGui().m_xGradientDialog = null;
           getGui().m_xControlDialogWindow.setEnable(true);
           getGui().m_xControlDialogWindow.setFocus();
        }

        if(event.Source.equals(getGui().m_xPropsTopWindow)){
            getGui().m_xPropsTopWindow.removeTopWindowListener(this);
            getGui().m_xPropsDialog.endExecute();
            getGui().m_xPropsTopWindow = null;
            getGui().m_xPropsDialog = null;

            getGui().m_xControlDialogWindow.setEnable(true);
            getGui().m_xControlDialogWindow.setFocus();

        }


    }

    // XTopWindowListener
    @Override
    public void windowOpened(EventObject arg0) {

    }

    // XTopWindowListener
    @Override
    public void windowClosed(EventObject arg0) {

    }

    // XTopWindowListener
    @Override
    public void windowMinimized(EventObject arg0) {
    }

    // XTopWindowListener
    @Override
    public void windowNormalized(EventObject arg0) {
    }

    // XTopWindowListener
    @Override
    public void windowActivated(EventObject arg0) {

    }

    // XTopWindowListener
    @Override
    public void windowDeactivated(EventObject arg0) {

    }

    // XTopWindowListener
    @Override
    public void disposing(EventObject arg0) {

    }


}
