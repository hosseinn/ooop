package org.openoffice.extensions.diagrams;

import com.sun.star.awt.ImageAlign;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XRadioButton;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.OrganizationChart;


import org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram.OrganizationDiagram;


public class Gui {

    private     Controller          m_Controller                = null;
    private     XComponentContext   m_xContext                  = null;
    private     XFrame              m_xFrame                    = null;
    private     Listener            m_oListener                 = null;
    private     XDialog             m_xSelectDialog             = null;
    protected   XTopWindow          m_xSelectDTopWindow         = null;
    private     XFixedText          m_XDiagramNameText          = null;
    private     XFixedText          m_XDiagramDescriptionText   = null;

    private     XDialog             m_xControlDialog            = null;
    protected   XWindow             m_xControlDialogWindow      = null;
    private     boolean             isVisibleControlDialog      = false;
    private     XTopWindow          m_xControlDialogTopWindow   = null;
    private     XTextComponent      m_xTextField                = null;
    protected   XControl            m_xImageControl             = null;
    private     Object              m_oUpAndDownButton          = null;
    private     boolean             m_isShownTextField          = true;
    protected   XDialog             m_xPaletteDialog            = null;
    protected   XTopWindow          m_xPaletteTopWindow         = null;
    private     int                 m_iInfoBoxOk                = -1;

    protected   XDialog             m_xGradientDialog           = null;
    protected   XWindow             m_xGradientWindow           = null;
    protected   XTopWindow          m_xGradientTopWindow        = null;
    protected   XControl            m_xStartImage               = null;
    protected   XControl            m_xEndImage                 = null;
    protected   String              m_sImageType                = "";

    protected   XDialog             m_xConvertDialog            = null;
    protected   XWindow             m_xConvertWindow            = null;
    protected   XTopWindow          m_xConvertTopWindow         = null;
    protected   XComboBox           m_xConvertComboBox          = null;

    protected   XDialog             m_xPropsDialog              = null;
    protected   XTopWindow          m_xPropsTopWindow           = null;
    protected   XRadioButton        m_xColorRadioButton         = null;
    protected   XRadioButton        m_xBaseColorRadioButton     = null;
    protected   XCheckBox           m_xGradientsCheckBox        = null;

    protected   XControl            m_xEventObjectControl       = null;
    protected   XControl            m_xColorCBControl           = null;
    protected   XCheckBox           m_xColorCheckBox            = null;

    protected   XControl            m_xColorImageControl        = null;
    protected   XControl            m_xStartColorImageControl   = null;
    protected   XControl            m_xEndColorImageControl     = null;

    protected   XControl            m_xStartColorLabelControl   = null;
    protected   XControl            m_xEndColorLabelControl     = null;

    protected   XControl            m_xBaseColorControl         = null;
    protected   XControl            m_xGradientsCheckBoxControl = null;
    protected   XControl            m_xMonographicOBYesControl  = null;
    protected   XControl            m_xMonographicOBNoControl   = null;
    protected   XControl            m_xFrameOBYesControl        = null;
    protected   XControl            m_xFrameOBNoControl         = null;
    protected   XControl            m_xFrameRoundedOBNoControl  = null;
    protected   XControl            m_xFrameRoundedOBYesControl = null;


    public Gui(){ }

    public boolean isShownTextField(){
        if(m_xTextField != null)
            return m_isShownTextField;
        return false;
    }

    public Gui(Controller controller, XComponentContext xContext, XFrame xFrame){
        m_Controller = controller;
        m_xContext = xContext;
        m_xFrame = xFrame;
        m_oListener = new Listener(this, m_Controller);
    }

    public Controller getController(){
        return m_Controller;
    }

    public void setVisibleSelectWindow(boolean bool){
        if(bool){
            if(isVisibleControlDialog)
                setVisibleControlDialog(false);
            getController().setGroupType(Controller.ORGANIGROUP);
            getController().setDiagramType(Controller.SIMPLEORGANIGRAM); //ORGANIGRAM
            createSelectDialog2();
            getController().removeDiagram();
            m_xSelectDialog.execute();  
        } else {
            m_xSelectDialog.endExecute();
            m_xSelectDialog = null;
        }
    }

    public XWindow getControlDialogWindow(){
        return m_xControlDialogWindow;
    }
    
    public void closeControlDialog(){
        if(m_xControlDialogWindow != null)
            m_xControlDialogWindow.setVisible(false);
        if(m_xControlDialogTopWindow != null)
            m_xControlDialogTopWindow.removeTopWindowListener(m_oListener);
        m_xControlDialogTopWindow = null;
        m_xControlDialogWindow = null;
        m_xControlDialog = null;
    }

    public void setVisibleControlDialog(boolean bool){

        int newDiagramId = getController().getDiagram().getDiagramID();
        // need to creat new controlDialog when a new diagram selected (not same the gui panels of diagrams)
        if( ( getController().getLastDiagramType() != 0 || getController().getLastDiagramID() != -1 ) && ( getController().getLastDiagramType() != getController().getDiagramType() || getController().getLastDiagramID() != newDiagramId ) ){
            if(m_xControlDialogWindow != null)
               m_xControlDialogWindow.setVisible(false);
            createControlDialog();
        }
        //if( bool && m_xControlDialogWindow == null)
        if( m_xControlDialogWindow == null)
            createControlDialog();
        if(m_xControlDialogWindow != null)
            m_xControlDialogWindow.setVisible(bool);
        if(bool){
            m_xControlDialogWindow.setFocus();
            isVisibleControlDialog = true;
        }else{
            isVisibleControlDialog = false;
        }
        getController().setLastDiagramType(getController().getDiagramType());
        getController().setLastDiagramID(newDiagramId);
    }

    public void setFocusControlDialog(){
        m_xControlDialogWindow.setFocus();
    }

    public boolean isVisibleControlDialog(){
        return isVisibleControlDialog;
    }

    public void createSelectDialog(){
        try {
            String sPackageURL              = getPackageLocation();
            String sDialogURL               = sPackageURL + "/dialogs/DiagramGallery.xdl";
            XDialogProvider2 xDialogProv    = getDialogProvider();
            m_xSelectDialog                 = xDialogProv.createDialogWithHandler( sDialogURL, m_oListener );
            m_xSelectDTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xSelectDialog);
            m_xSelectDTopWindow.addTopWindowListener(m_oListener);

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xSelectDialog);

            Object oFixedText = xControlContainer.getControl("diagramNameLabel");
            m_XDiagramNameText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            oFixedText = xControlContainer.getControl("diagramDescriptionLabel");
            m_XDiagramDescriptionText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            Object oButton = xControlContainer.getControl("OrganigramButton");
            setImageOfButton(oButton, sPackageURL + "/images/orgchart.png", (short)-1);

            oButton = xControlContainer.getControl("VennDiagramButton");
            setImageOfButton(oButton, sPackageURL + "/images/venn.png", (short)-1);

            oButton = xControlContainer.getControl("PyramidDiagramButton");
            setImageOfButton(oButton, sPackageURL + "/images/pyramid.png", (short)-1);

            oButton = xControlContainer.getControl("CycleDiagramButton");
            setImageOfButton(oButton, sPackageURL + "/images/ring.png", (short)-1);

        }catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void createSelectDialog2(){
        try {
            String sPackageURL              = getPackageLocation();
            String sDialogURL               = sPackageURL + "/dialogs/DiagramGallery2.xdl";
            XDialogProvider2 xDialogProv    = getDialogProvider();
            m_xSelectDialog                 = xDialogProv.createDialogWithHandler( sDialogURL, m_oListener );
            m_xSelectDTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xSelectDialog);
            m_xSelectDTopWindow.addTopWindowListener(m_oListener);

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xSelectDialog);

            Object oFixedText = xControlContainer.getControl("Label1");
            m_XDiagramNameText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            oFixedText = xControlContainer.getControl("Label2");
            m_XDiagramDescriptionText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            Object oButton = xControlContainer.getControl("Item0");
            setImageOfButton(oButton, sPackageURL + "/images/orgchart.png", (short)-1);
            oButton = xControlContainer.getControl("Item1");
            setImageOfButton(oButton, sPackageURL + "/images/hororgchart.png", (short)-1);
            oButton = xControlContainer.getControl("Item2");
            setImageOfButton(oButton, sPackageURL + "/images/tablediagram.png", (short)-1);
            oButton = xControlContainer.getControl("Item3");
            setImageOfButton(oButton, sPackageURL + "/images/orgchart2.png", (short)-1);
            oButton = xControlContainer.getControl("Item4");
            setImageOfButton(oButton, sPackageURL + "/images/venn.png", (short)-1);
            oButton = xControlContainer.getControl("Item5");
            setImageOfButton(oButton, sPackageURL + "/images/ring.png", (short)-1);
            oButton = xControlContainer.getControl("Item6");
            setImageOfButton(oButton, sPackageURL + "/images/pyramid.png", (short)-1);
            //oButton = xControlContainer.getControl("Item7");

            setSelectDialog2Images();
            setSelectDialogText(Controller.SIMPLEORGANIGRAM);
            
        }catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }
/*
    public void setFirstItemOnFocus(){
        if(m_xSelectDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xSelectDialog);
            Object oButton = xControlContainer.getControl("Item0");
            XControl xButtonControl = (XControl)UnoRuntime.queryInterface(XControl.class, oButton);
            if(xButtonControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xButtonControl.getModel());
                xPropImage.setPropertyValue("FocusOnClick", new Boolean(true));
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        }
    }
*/
    public void  setSelectDialog2Images(){
        int n = getController().getGroupType() * 4;
        int m = n + 3;
        if(getController().getGroupType() == 1)
            m -= 1;
        if(getController().getGroupType() == 4)
            m -= 2;
        if(m_xSelectDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xSelectDialog);
            if(xControlContainer != null){
                boolean bool;
                for(int i = 0; i <= 7; i++){
                    bool = false;
                    if(i >= n && i <= m)
                        bool = true;
                    XControl xItemControl = (XControl)UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl("Item" + i));
                    enableVisibleControl(xItemControl, bool);
                }
            }
        }
    }

    public void testProps(Object obj){
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, obj);
            Property[] props = xProp.getPropertySetInfo().getProperties();
            for (Property p : props)
                System.out.println(p.Name + " "  + p.Type.getTypeName());
    }
    
    public void createControlDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ControlDialog" + (getController().getGroupType() != Controller.ORGANIGROUP ? 1 : 2) + ".xdl";
            m_xControlDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xControlDialog != null) {
                m_xControlDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xControlDialog);
                m_xControlDialogTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xControlDialogWindow);
                m_xControlDialogTopWindow.addTopWindowListener(m_oListener);
            }

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            Object oTextField = xControlContainer.getControl("textField");
            m_xTextField = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, oTextField);

            m_xImageControl = xControlContainer.getControl("ImageControl");

            if(getController().getGroupType() == Controller.ORGANIGROUP){

                m_oUpAndDownButton = xControlContainer.getControl("downUpButton");
                m_isShownTextField = true;
                textFieldDownUp();
                
                if(getController().getDiagram() != null)
                    ((OrganizationChart)getController().getDiagram()).setNewItemHType(OrganizationDiagram.UNDERLING);

                Object oConvertButton = xControlContainer.getControl("convertButton");
                setImageOfButton(oConvertButton, sPackageURL + "/images/convert.png", ImageAlign.LEFT);

                
                
            }

            Object oComboBox = xControlContainer.getControl("StyleComboBox");
            XComboBox xComboBox = (XComboBox) UnoRuntime.queryInterface(XComboBox.class, oComboBox);
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                setImageColorOfControlDialog(10079487);
                xComboBox.removeItems((short)1, (short)3);
                String[] newStyles = { getDialogPropertyValue("Strings", "CycleStyleComboBox1"),
                                       getDialogPropertyValue("Strings", "CycleStyleComboBox2"),
                                       getDialogPropertyValue("Strings", "CycleStyleComboBox3")
                                     };
                xComboBox.addItems(newStyles, (short)1);
            }
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
                setImageColorOfControlDialog(10079487);
                xComboBox.removeItems((short)1, (short)3);
                String[] newStyles = { getDialogPropertyValue("Strings", "PyramidStyleComboBox1"),
                                       getDialogPropertyValue("Strings", "PyramidStyleComboBox2"),
                                       getDialogPropertyValue("Strings", "PyramidStyleComboBox3")
                                     };
                xComboBox.addItems(newStyles, (short)1);
            }


            /*
             *
                // create add and remove button images
                Object oAddButton = xControlContainer.getControl("addShapeButton");
                Object oRemoveButton = xControlContainer.getControl("removeShapeButton");

                if(getController().getDiagramType() == Controller.ORGANIGRAM){
                    setImageOfButton(oAddButton, sPackageURL + "/images/organigram_add.png", ImageAlign.LEFT);
                    setImageOfButton(oRemoveButton, sPackageURL + "/images/organigram_remove.png", ImageAlign.LEFT);
                }
                if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                    setImageOfButton(oAddButton, sPackageURL + "/images/venn_add.png", ImageAlign.LEFT);
                    setImageOfButton(oRemoveButton, sPackageURL + "/images/venn_remove.png", ImageAlign.LEFT);
                }
                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
                    setImageOfButton(oAddButton, sPackageURL + "/images/pyramid_add.png", ImageAlign.LEFT);
                    setImageOfButton(oRemoveButton, sPackageURL + "/images/pyramid_remove.png", ImageAlign.LEFT);
                }
                if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                    setImageOfButton(oAddButton, sPackageURL + "/images/ring_add.png", ImageAlign.LEFT);
                    setImageOfButton(oRemoveButton, sPackageURL + "/images/ring_remove.png", ImageAlign.LEFT);
                }
            */
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void test(Object obj){
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, obj);
            Property[] props = xProp.getPropertySetInfo().getProperties();
            for (Property p : props)
                System.out.println(p.Name + " "  + p.Type.getTypeName());
    }

    public void textFieldModified(){
        if(m_xTextField != null)
            setTextOfSelectedShape(m_xTextField.getText());
    }

    public String getTextField(){
        if(m_xTextField != null)
            return m_xTextField.getText();
        return "";
    }

    public void setTextField(String str){
        m_xTextField.setText(str);
    }

    public void setTextOfSelectedShape(String str){
        if(getController().getSelectedShapes().getCount() == 1){
            XShape xSelectedShape = getController().getSelectedShape();
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
            String selectedShapeName = xNamed.getName();
            if((selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram")) && selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0")){
                XText xText = (XText)UnoRuntime.queryInterface(XText.class, xSelectedShape);
                xText.setString(str);
            }else{
                //error message
            }
        }else{
            //error message
        }
    }

    public void setFocusTextField(){
        if(m_xTextField != null){
            XWindow xWindow = (XWindow)UnoRuntime.queryInterface(XWindow.class, m_xTextField);
            if(xWindow != null)
                xWindow.setFocus();
        }
    }

    public void setControllerTextField(){
        if(m_xTextField != null){
            if(getController().isOnlySimpleItemIsSelected()){
                XText xText = (XText)UnoRuntime.queryInterface(XText.class, getController().getSelectedShape());
                setTextField(xText.getString());
            }else{
                setTextField("Inactive function - this option can be used if only simple item has been selected");
            }
        }
    }

    public void enableVisibleTextField(boolean bool){
        enableVisibleControl((XControl) UnoRuntime.queryInterface(XControl.class, m_xTextField), bool);
    }

    public void enableTextField(boolean bool){
        enableControl((XControl) UnoRuntime.queryInterface(XControl.class, m_xTextField), bool);
        if(!bool)
            setFocusControlDialog();
    }

    public void textFieldDownUp(){
        try {
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, m_xControlDialog);
            XPropertySet xPS = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());

            String sPackageURL = getPackageLocation();
            m_isShownTextField = !m_isShownTextField;

            if(m_isShownTextField){
                enableVisibleTextField(true);
                if(getController().isOnlySimpleItemIsSelected())
                    setFocusTextField();
                else
                    enableTextField(false);
                xPS.setPropertyValue("Width", new Integer(272));
                xPS.setPropertyValue("Height", new Integer(35));
                if(m_oUpAndDownButton != null)
                    setImageOfButton(m_oUpAndDownButton, sPackageURL + "/images/up.png", (short)-1);
            }else{
                enableVisibleTextField(false);
                xPS.setPropertyValue("Width", new Integer(272));
                xPS.setPropertyValue("Height", new Integer(15));
                if(m_oUpAndDownButton != null)
                    setImageOfButton(m_oUpAndDownButton, sPackageURL + "/images/down.png", (short)-1);
            }
        } catch (PropertyVetoException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.out.println(ex.getLocalizedMessage());
        }
    }

    public void showConvertDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ConvertDialog.xdl";
            m_xConvertDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);

            if (m_xConvertDialog != null){
                m_xConvertComboBox = (XComboBox)UnoRuntime.queryInterface(XComboBox.class, xControlContainer.getControl("lastHorLevelComboBox"));
                setConvertDialogRadioButtons();
                m_xConvertWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xConvertDialog);
                m_xConvertTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xConvertDialog);
                m_xConvertTopWindow.addTopWindowListener(m_oListener);
                m_xControlDialogWindow.setEnable(false);
                m_xConvertDialog.execute();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public short getConversationType(){
        if(m_xConvertDialog != null){
            XRadioButton xRadioButton = null;
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton1"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.SIMPLEORGANIGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.HORIZONTALORGANIGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton3"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.TABLEHIERARCHYDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton4"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.ORGANIGRAM;
            }
        }
        return -1;
    }

    public void setConvertDialogRadioButtons(){
        if(m_xConvertDialog != null){
            try {
                XRadioButton xRadioButton2          = null;
                XControl xControl                   = null;
                XPropertySet xProps                 = null;
                String disabledButton               = "convertOptionButton";
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);

                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM){
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
                        xRadioButton2.setState(true);
                        disabledButton += 1;
                    }
                    if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                        disabledButton += 2;
                    if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                        disabledButton += 3;
                    //if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    //    disabledButton += 4;

                    if(getController().getDiagramType() != Controller.ORGANIGRAM){
                        xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl(disabledButton));
                        xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                        xProps.setPropertyValue("Enabled", new Boolean(false));
                    }
                }
                //return AnyConverter.toBoolean(xPropImage.getPropertyValue("EnableVisible"));
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void setConvertComboBox(boolean bool){
        try {
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl("lastHorLevelComboBox"));
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            xProps.setPropertyValue("Enabled", new Boolean(bool));
            xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl("Label"));
            xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            xProps.setPropertyValue("Enabled", new Boolean(bool));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
/*
    public void closeConvertDialog(){
        if(m_xConvertDialog != null)
            m_xConvertDialog.endExecute();
        if(m_xControlDialogWindow != null){
            m_xControlDialogWindow.setEnable(true);
            setFocusControlDialog();
        }
    }
*/
    public void showGradientDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/GradientDialog.xdl";
            m_xGradientDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xGradientDialog != null) {
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
                m_xStartImage = xControlContainer.getControl("startColor");
                m_xEndImage = xControlContainer.getControl("endColor");

                m_xGradientWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xGradientDialog);
                m_xGradientTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xGradientDialog);
                m_xGradientTopWindow.addTopWindowListener(m_oListener);
                m_xControlDialogWindow.setEnable(false);
                m_xGradientDialog.execute();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void showColorTable() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ColorTable.xdl";
            m_xPaletteDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xPaletteDialog != null) {
                m_xPaletteTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xPaletteDialog);
                m_xPaletteTopWindow.addTopWindowListener(m_oListener);
                m_xControlDialogWindow.setEnable(false);
                m_xPaletteDialog.execute();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void showDiagramPropsDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String diagramDefine = "";
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                diagramDefine = "OrganigramPropsDialog.xdl";
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                diagramDefine = "VennDiagramPropsDialog.xdl";
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                diagramDefine = "PyramidDiagramPropsDialog.xdl";
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                diagramDefine = "CycleDiagramPropsDialog.xdl";
            String sDialogURL = sPackageURL + "/dialogs/" + diagramDefine;
            m_xPropsDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);

            if (m_xPropsDialog != null) {

                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
                
                Object oRadioButton = xControlContainer.getControl("colorOptionButton");
                m_xColorRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, oRadioButton);

                m_xColorImageControl = xControlContainer.getControl("colorImageControl");

                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
                    m_xGradientsCheckBoxControl = xControlContainer.getControl("gradientsCheckBox");
                    m_xGradientsCheckBox = (XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, m_xGradientsCheckBoxControl);
                }else{
                   m_xGradientsCheckBoxControl = null;
                   m_xGradientsCheckBox = null;
                }

                if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                    m_xColorCBControl = xControlContainer.getControl("colorOptionButton");
                    m_xColorCheckBox = UnoRuntime.queryInterface(XCheckBox.class, m_xColorCBControl);
                    m_xFrameRoundedOBYesControl = xControlContainer.getControl("frameRoundedOptionButtonYes");
                    m_xFrameRoundedOBNoControl = xControlContainer.getControl("frameRoundedOptionButtonNo");
                }else{
                    m_xFrameRoundedOBYesControl = null;
                    m_xFrameRoundedOBNoControl = null;
                }

                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getGroupType() == Controller.ORGANIGROUP){
                    m_xStartColorImageControl = xControlContainer.getControl("startColorImageControl");
                    m_xEndColorImageControl = xControlContainer.getControl("endColorImageControl");

                    m_xStartColorLabelControl = xControlContainer.getControl("Label0");
                    m_xEndColorLabelControl = xControlContainer.getControl("Label1");
                }else{
                    m_xStartColorImageControl = null;
                    m_xEndColorImageControl = null;
                    m_xStartColorLabelControl = null;
                    m_xEndColorLabelControl = null;
                }

                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                    m_xBaseColorControl = xControlContainer.getControl("baseColorOptionButton");
                    m_xBaseColorRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorControl);
                }else{
                    m_xBaseColorControl = null;
                    m_xBaseColorRadioButton = null;
                }

                if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getDiagramType() == Controller.CYCLEDIAGRAM || getController().getDiagramType() == Controller.VENNDIAGRAM){
                    m_xMonographicOBYesControl = xControlContainer.getControl("monographicOptionButtonYes");
                    m_xMonographicOBNoControl = xControlContainer.getControl("monographicOptionButtonNo");
                }else{
                    m_xMonographicOBYesControl = null;
                    m_xMonographicOBNoControl = null;
                }
                

                if(getController().getDiagramType() == Controller.CYCLEDIAGRAM || getController().getDiagramType() == Controller.VENNDIAGRAM){
                    m_xFrameOBYesControl = xControlContainer.getControl("frameOptionButtonYes");
                    m_xFrameOBNoControl = xControlContainer.getControl("frameOptionButtonNo");
                }else{
                    m_xFrameOBYesControl = null;
                    m_xFrameOBNoControl = null;
                }

                m_xPropsTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xPropsDialog);
                m_xPropsTopWindow.addTopWindowListener(m_oListener);
                m_xControlDialogWindow.setEnable(false);
                m_xPropsDialog.execute();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            location =  xPIP.getPackageLocation("org.openoffice.extensions.diagrams.Diagrams");
        } catch (NoSuchElementException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return location;
    }

    public XDialogProvider2 getDialogProvider(){
        XDialogProvider2 xDialogProv = null;
        try {
            XModel xModel = m_xFrame.getController().getModel();
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object obj;
            if (xModel != null) {
                Object[] args = new Object[1];
                args[0] = xModel;
                obj = xMCF.createInstanceWithArgumentsAndContext("com.sun.star.awt.DialogProvider2", args, m_xContext);
            } else {
                obj = xMCF.createInstanceWithContext("com.sun.star.awt.DialogProvider2", m_xContext);
            }
            xDialogProv = (XDialogProvider2) UnoRuntime.queryInterface(XDialogProvider2.class, obj);
        }catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xDialogProv;
    }

    public void showMessageBox(String sTitle, String sMessage){
        try{
            Object oToolkit = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit",m_xContext);
            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);
            if ( m_xFrame != null && xToolkit != null ) {
                WindowDescriptor aDescriptor = new WindowDescriptor();
                aDescriptor.Type              = WindowClass.MODALTOP;
                aDescriptor.WindowServiceName = "infobox";
                aDescriptor.ParentIndex       = -1;
                aDescriptor.Parent            = (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class, m_xFrame.getContainerWindow());
                //aDescriptor.Bounds            = new Rectangle(0,0,300,200);
                aDescriptor.WindowAttributes  = WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.CLOSEABLE;
                XWindowPeer xPeer = xToolkit.createWindow( aDescriptor );
                if ( null != xPeer ) {
                   XMessageBox xMessageBox = (XMessageBox)UnoRuntime.queryInterface(XMessageBox.class, xPeer);
                    if ( null != xMessageBox ){
                        xMessageBox.setCaptionText( sTitle );
                        xMessageBox.setMessageText( sMessage );
                        m_xControlDialogWindow.setEnable(false);
                        xMessageBox.execute();
                        m_xControlDialogWindow.setEnable(true);
                        m_xControlDialogWindow.setFocus();
                    }
                }
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void showMessageBox2(String sTitle, String sMessage){
        XComponent xComponent = null;
        try {
            Object oToolkit = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", m_xContext);
            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);
            XMessageBoxFactory xMessageBoxFactory = (XMessageBoxFactory)UnoRuntime.queryInterface(XMessageBoxFactory.class, oToolkit);
            if ( m_xFrame != null && xToolkit != null ) {
                WindowDescriptor aDescriptor = new WindowDescriptor();
                aDescriptor.Type              = WindowClass.MODALTOP;
                aDescriptor.ParentIndex       = -1;
                aDescriptor.Parent            = (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class, m_xFrame.getContainerWindow());
                XWindowPeer xPeer = xToolkit.createWindow( aDescriptor );
                if ( null != xPeer ) {
                    // rectangle may be empty if position is in the center of the parent peer
                    Rectangle aRectangle = new Rectangle();
                    XMessageBox xMessageBox = xMessageBoxFactory.createMessageBox( xPeer, aRectangle, "messbox", com.sun.star.awt.MessageBoxButtons.BUTTONS_OK_CANCEL, sTitle, sMessage);
                    xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                    if (xMessageBox != null){
                        m_xControlDialogWindow.setEnable(false);
                        m_iInfoBoxOk = xMessageBox.execute();
                        m_xControlDialogWindow.setEnable(true);
                        m_xControlDialogWindow.setFocus();
                    }
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
             System.err.println(ex.getLocalizedMessage());
        }finally{
            if (xComponent != null)
                xComponent.dispose();
        }
    }

    public void setSelectDialogText(){
        setSelectDialogText(getController().getDiagramType());
    }

    public void setSelectDialogText(short type){
        String sType = "";
        String sourceFileName = "Strings";
        if( type == Controller.ORGANIGRAM){
            sType = "Organigram2";
            sourceFileName += "2";
        }
        if(type == Controller.VENNDIAGRAM )
            sType = "Venndiagram";
        if( type == Controller.PYRAMIDDIAGRAM )
            sType = "Pyramiddiagram";
        if( type == Controller.CYCLEDIAGRAM )
            sType = "Cyclediagram";
        if( type == Controller.HORIZONTALORGANIGRAM ){
            sType = "HorizontalOrganigram";
            sourceFileName += "2";
        }
        if( type == Controller.TABLEHIERARCHYDIAGRAM ){
            sType = "TableHierarchyDiagram";
            sourceFileName += "2";
        }
        if( type == Controller.SIMPLEORGANIGRAM)
            sType = "Organigram";
        if(type == Controller.NOTDIAGRAM){
            m_XDiagramNameText.setText("");
            m_XDiagramDescriptionText.setText("");
        }else{
            String diagramNameProperty = sourceFileName + "." + sType + ".Label";
            String diagramDescProperty = sourceFileName + "." + sType + "Description.Label";
            m_XDiagramNameText.setText(getDialogPropertyValue(sourceFileName, diagramNameProperty));
            m_XDiagramDescriptionText.setText(getDialogPropertyValue(sourceFileName, diagramDescProperty));
        }
    }

    public String getDialogPropertyValue(String dialogName, String propertyName){
        String result = null;
        XStringResourceWithLocation xResources = null;
        String m_resRootUrl = getPackageLocation() + "/dialogs/";
        try {
            xResources = StringResourceWithLocation.create(m_xContext, m_resRootUrl, true, getController().getLocation(), dialogName, "", null);
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        // map properties
        if(xResources != null){
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++)
                if(ids[i].contains(propertyName))
                    result = xResources.resolveString(ids[i]);
        }
        return result;
    }

    public int getNum(String name){
        String s ="";
        char[] charName = name.toCharArray();
        int i = 5;
        while(i<name.length())
           s +=  charName[i++];
        return getController().parseInt(s);
    }

    public void setImageOfButton(Object oButton, String sImageUrl, short imageAlign){
        if(oButton != null){
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, oButton);
            if(xControl != null){
                try {
                    XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                    if(imageAlign == ImageAlign.LEFT || imageAlign == ImageAlign.RIGHT || imageAlign == ImageAlign.TOP || imageAlign == ImageAlign.BOTTOM)
                        xProps.setPropertyValue("ImageAlign", new Short(imageAlign));
                    xProps.setPropertyValue("ImageURL", sImageUrl);
                } catch (PropertyVetoException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (IllegalArgumentException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (UnknownPropertyException ex) {
                    System.err.println(ex.getLocalizedMessage());
                }
            }
        }
    }

    public int getCurrImageColor(XDialog xDialog, int i){
        int color = -1;
        try {
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, xDialog);
            XControl xImageControl = xControlContainer.getControl("ImageControl" + i);
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xImageControl.getModel());
            color = AnyConverter.toInt(xPropImage.getPropertyValue("BackgroundColor"));
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return color;
    }

    public boolean isEnableControlDialogImageColor() throws IllegalArgumentException, UnknownPropertyException, WrappedTargetException{
        if(m_xImageControl != null){
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xImageControl.getModel());
            return AnyConverter.toBoolean(xPropImage.getPropertyValue("EnableVisible"));
        }
        return false;
    }


    public void enableVisibleControl(XControl xControl, boolean bool){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("EnableVisible", bool);
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void enableControl(XControl xControl, boolean bool){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("Enabled", new Boolean(bool));
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void enableControlDialogImageColor() throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException{
        if(m_xImageControl != null){
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xImageControl.getModel());
            xPropImage.setPropertyValue("EnableVisible", new Boolean(true));
        }
    }

    public void disableControlDialogImageColor() throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException{
        if(m_xImageControl != null){
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xImageControl.getModel());
            xPropImage.setPropertyValue("EnableVisible", new Boolean(false));
        }
    }

    public void setImageColorOfControlDialog(int color){
        setImageColorOfImageControl(m_xImageControl, color);
    }

    public void setImageColorOfImageControl(XControl xImageControl, int color){
        if(xImageControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xImageControl.getModel());
                xPropImage.setPropertyValue("BackgroundColor", new Integer(color));
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public int getImageColorOfControlDialog(){
        return getImageColorOfControlDialog(m_xImageControl);
    }

    public int getImageColorOfControlDialog(XControl xControl){
        int color = -1;
        try {
            if(xControl != null){
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                color = AnyConverter.toInt(xPropImage.getPropertyValue("BackgroundColor"));
            }

        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return color;
    }
  
    public void askUserForRepair(OrganizationChart organigram){
        String sTitle = getDialogPropertyValue("Strings", "TreeBuildError.Title");
        String sMessage = getDialogPropertyValue("Strings", "TreeBuildError.Message");
        showMessageBox2(sTitle, sMessage);
        if(m_iInfoBoxOk == 1){
            organigram.repairDiagram();
            organigram.refreshDiagram();
        }
        m_iInfoBoxOk = -1;
    }   

    public void setEnableControlDialogWindow(boolean bool){
        if(m_xControlDialogWindow != null)
            m_xControlDialogWindow.setEnable(bool);
    }
    
}
