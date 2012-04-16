package oxygenoffice.extensions.smart;

import com.sun.star.awt.ImageAlign;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XRadioButton;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChart;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.buttdiagram.TargetDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram.CycleDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram.PyramidDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram.VennDiagram;


public class Gui {


    public      boolean             isShownTips;

    private     XComponentContext   m_xContext                  = null;
    private     XFrame              m_xFrame                    = null;
    private     Controller          m_Controller                = null;
    private     Listener            m_oListener                 = null;

    protected   String              m_sImageType                = "ColorImage";
    protected   XControl            m_xEventObjectControl       = null;

    //DiagramGalleryDialog *************************************************
    private     XDialog             m_xGalleryDialog                = null;
    protected   XTopWindow          m_xGalleryDialogTopWindow       = null;
    private     XFixedText          m_XDiagramNameText              = null;
    private     XFixedText          m_XDiagramDescriptionText       = null;
    //**********************************************************************


    //ControlDialog ********************************************************
    private     XDialog             m_xControlDialog                = null;
    private     XWindow             m_xControlDialogWindow          = null;
    protected   XTopWindow          m_xControlDialogTopWindow       = null;
    private     XControl            m_xICOfControlDialog            = null;
    private     Object              m_oUpAndDownBOfControlDialog    = null;
    private     XTextComponent      m_xTFOfControlDialog            = null;

    private     boolean             isVisibleControlDialog          = false;
    private     boolean             m_isShownTFOfControlDialog      = true;
    //**********************************************************************


    //ColorTable dialog ****************************************************
    private     XDialog             m_xColorTableDialog             = null;
    protected   XTopWindow          m_xColorTableTopWindow          = null;
    //**********************************************************************


    //GradientDialog *******************************************************
    private     XDialog             m_xGradientDialog               = null;
    private     XWindow             m_xGradientDialogWindow         = null;
    protected   XTopWindow          m_xGradientDialogTopWindow      = null;
    private     XControl            m_xStartImageOfGradientDialog   = null;
    private     XControl            m_xEndImageOfGradientDialog     = null;
    //**********************************************************************


    //ConvertDialog *******************************************************
    private     XDialog             m_xConvertDialog                = null;
    //private     XWindow             m_xConvertDialogWindow        = null;
    protected   XTopWindow          m_xConvertDialogTopWindow       = null;
    private     XListBox            m_xListBoxOfConvertDialog       = null;
    //**********************************************************************


    //Properties dialogs ***************************************************
    private     boolean             m_IsShownPropsDialog            = false;

    private     XDialog             m_xPropsDialog                  = null;
    private     XWindow             m_xPropsDialogWindow            = null;
    protected   XTopWindow          m_xPropsDialogTopWindow         = null;

    //protected   XControl            m_xColorCBControl             = null;
    //protected   XCheckBox           m_xColorCheckBox              = null;
    //protected   XCheckBox           m_xGradientsCheckBox          = null;
    //protected   XControl            m_xGradientsCheckBoxControl   = null;
    protected   XCheckBox           m_xModifyColorsCheckBoxOfPD     = null;

    protected   XControl            m_xColorOBControlOfPD           = null;
    //private     XRadioButton        m_xColorRadioButtonOfPD       = null;

    protected   XControl            m_xBaseColorOBControlOfPD       = null;
    protected   XRadioButton        m_xBaseColorRadioButtonOfPD     = null;
    public      XControl            m_xColorImageControlOfPD        = null;
    protected   XControl            m_xGradientsOBControlOfPD       = null;
    private     XControl            m_xGradientModeLBControlOfPD    = null;
    public      XControl            m_xStartColorImageControlOfPD   = null;
    public      XControl            m_xEndColorImageControlOfPD     = null;
    private     XControl            m_xStartColorLabelControlOfPD   = null;
    private     XControl            m_xEndColorLabelControlOfPD     = null;
    private     XControl            m_xYesOutlineOBControlOfPD      = null;
    private     XControl            m_xNoOutlineOBControlOfPD       = null;

    protected   XControl            m_xYesShadowOBOfPD              = null;
    protected   XControl            m_xNoShadowOBOfPD               = null;

    private     XControl            m_xYesFrameOBControlOfPD        = null;
    private     XControl            m_xNoFrameOBControlOfPD         = null;
    protected   XControl            m_xNoFrameRoundedOBControlOfPD  = null;
    protected   XControl            m_xYesFrameRoundedOBControlOfPD = null;

    protected   XFixedText          m_xAreaLabelOfPD                = null;
    protected   XFixedText          m_xModifiesColorLabelOfPD       = null;
    protected   XFixedText          m_xColorModeLabelOfPD           = null;
    protected   XFixedText          m_xRoundedLabelOfPD             = null;
    protected   XFixedText          m_xOutlineLabelOfPD             = null;
    protected   XFixedText          m_xShadowLabelOfPD              = null;
    protected   XFixedText          m_xTransparencyLabelOfPD        = null;
    protected   XFixedText          m_xTFrameLabelOfPD              = null;
    protected   XFixedText          m_xRoundedTFrameLabelOfPD       = null;
    protected   XFixedText          m_xLayoutLabelOfPD              = null;

    //**********************************************************************

    //MessageBoxes *********************************************************
    private     int                 m_iInfoBoxOk                = -1;
    //**********************************************************************
   

    public Gui(){ }

    public Gui(Controller controller, XComponentContext xContext, XFrame xFrame){
        isShownTips = true;
        m_Controller = controller;
        m_xContext = xContext;
        m_xFrame = xFrame;
        m_oListener = new Listener(this, m_Controller);
    }

    public Controller getController(){
        return m_Controller;
    }

    //DiagramGalleryDialog's methods ************************************************************************
    public void executeGalleryDialog(boolean bool){
        if(bool){
            if(isVisibleControlDialog)
                setVisibleControlDialog(false);
            getController().setGroupType(Controller.ORGANIGROUP);
            getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
            getController().removeDiagram();
            createDiagramGallery2();
            m_xGalleryDialog.execute();
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xGalleryDialog);
            xComp.dispose();
        } else {
            if(m_xGalleryDialog != null)
                m_xGalleryDialog.endExecute();
            disposeGalleryDialog();
        }
    }


    public boolean isShownTips(){
        return isShownTips;
    }

    public void setShownTips(){
        isShownTips = false;
        if(m_xGalleryDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            if(xControlContainer != null){
                XCheckBox xTipsCheckBox = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("TipsCheckBox"));
                if(xTipsCheckBox.getState() == 1)
                    isShownTips = true;
            }
        }
    }

    public void createDiagramGallery2(){
        try {
            String sPackageURL              = getPackageLocation();
            String sDialogURL               = sPackageURL + "/dialogs/DiagramGallery2.xdl";
            XDialogProvider2 xDialogProv    = getDialogProvider();
            m_xGalleryDialog = xDialogProv.createDialogWithHandler( sDialogURL, m_oListener );
            m_xGalleryDialogTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xGalleryDialog);
            m_xGalleryDialogTopWindow.addTopWindowListener(m_oListener);

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            XListBox diagramGroupList = (XListBox) UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("ListBox"));
            diagramGroupList.removeItems((short)2, (short)4);
            diagramGroupList.selectItemPos((short)0, true);
            m_XDiagramNameText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label1"));
            m_XDiagramDescriptionText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label2"));
            Object oButton = xControlContainer.getControl("Item0");
            setImageOfObject(oButton, sPackageURL + "/images/simpleOrgchart.png", (short)-1);
            oButton = xControlContainer.getControl("Item1");
            setImageOfObject(oButton, sPackageURL + "/images/hororgchart.png", (short)-1);
            oButton = xControlContainer.getControl("Item2");
            setImageOfObject(oButton, sPackageURL + "/images/tableHierarchy.png", (short)-1);
            oButton = xControlContainer.getControl("Item3");
            setImageOfObject(oButton, sPackageURL + "/images/orgchart.png", (short)-1);
            oButton = xControlContainer.getControl("Item4");
            setImageOfObject(oButton, sPackageURL + "/images/venn.png", (short)-1);
            oButton = xControlContainer.getControl("Item5");
            setImageOfObject(oButton, sPackageURL + "/images/ring.png", (short)-1);
            oButton = xControlContainer.getControl("Item6");
            setImageOfObject(oButton, sPackageURL + "/images/pyramid.png", (short)-1);
            oButton = xControlContainer.getControl("Item7");
            setImageOfObject(oButton, sPackageURL + "/images/target.png", (short)-1);

            setGalleryDialog2Images();
            setGalleryDialogText(Controller.SIMPLEORGANIGRAM);

            XCheckBox xTipsCheckBox = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("TipsCheckBox"));
            if(isShownTips)
                xTipsCheckBox.setState((short)1);
            else
                xTipsCheckBox.setState((short)0);

        }catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void  setGalleryDialog2Images(){
        int numOfButtons = 7;
        int n = getController().getGroupType() * 4;
        int m = n + 3;
        if(getController().getGroupType() == Controller.MATRIXGROUP)
            m -= 2;
        if(m_xGalleryDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            if(xControlContainer != null){
                boolean bool;
                for(int i = 0; i <= numOfButtons; i++){
                    bool = false;
                    if(i >= n && i <= m)
                        bool = true;
                    XControl xItemControl = (XControl)UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl("Item" + i));
                    enableVisibleControl(xItemControl, bool);
                }
            }
        }
    }

    public void setGalleryDialogText(){
        setGalleryDialogText(getController().getDiagramType());
    }

    public void setGalleryDialogText(short type){
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
        if( type == Controller.TARGETDIAGRAM ){
            sType = "ButtDiagram";
            sourceFileName += "2";
        }
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

    public void disposeGalleryDialog(){
        if(m_xGalleryDialogTopWindow != null)
            m_xGalleryDialogTopWindow.removeTopWindowListener(m_oListener);
        //m_xGalleryDialogTopWindow = null;
        //m_xGalleryDialog = null;
    }
    //*******************************************************************************************************


    //ControlDialog1s methods *******************************************************************************
    public void setVisibleControlDialog(boolean bool){
        int newDiagramId = getController().getDiagram().getDiagramID();
        // need to creat new controlDialog when a new diagram selected (not same the gui panels of diagrams)
        if( ( getController().getLastDiagramType() != -1 || getController().getLastDiagramID() != -1 ) && ( getController().getLastDiagramType() != getController().getDiagramType() || getController().getLastDiagramID() != newDiagramId ) ){
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
            
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
                m_xICOfControlDialog = xControlContainer.getControl("ImageControl");
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    getController().getDiagram().setGradientProps(getController().getDiagram().isGradientProps());
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    setColorModeOfImageOfControlDialog();
                //XListBox xListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("gradientModeListBox"));
                if(getController().getGroupType() == Controller.ORGANIGROUP && getController().getDiagram() != null)
                    ((OrganizationChart)getController().getDiagram()).setNewItemHType(OrgChart.UNDERLING);
                setImageOfObject(xControlContainer.getControl("convertButton"), sPackageURL + "/images/convert.png", ImageAlign.LEFT);
                m_oUpAndDownBOfControlDialog = xControlContainer.getControl("downUpButton");
                m_xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
                m_isShownTFOfControlDialog = true;
                moveTextFieldOfControlDialog();
//              if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
//                  setImageColorOfControlDialog(10079487);
 //             if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
 //                 setImageColorOfControlDialog(10079487);
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void closeAndDisposeControlDialog(){
        if(m_xControlDialogTopWindow != null)
            m_xControlDialogTopWindow.removeTopWindowListener(m_oListener);
        if(m_xControlDialogWindow != null){
            m_xControlDialogWindow.setVisible(false);
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xControlDialogWindow);
            if(xComp != null)
                xComp.dispose();
        }
        m_xControlDialogTopWindow = null;
        m_xControlDialogWindow = null;
        m_xControlDialog = null;
    }

    public boolean isVisibleControlDialog(){
        return isVisibleControlDialog;
    }

    public void enableControlDialogWindow(boolean bool){
        if(m_xControlDialogWindow != null)
            m_xControlDialogWindow.setEnable(bool);
    }
 
    public void setFocusControlDialog(){
        if(m_xControlDialogWindow != null)
            m_xControlDialogWindow.setFocus();
    }

    public XWindow getControlDialogWindow(){
        return m_xControlDialogWindow;
    }

    public void setColorModeOfImageOfControlDialog(){
        
        if(getController().getGroupType() == Controller.ORGANIGROUP){
            if(getController().getDiagram().isGradientProps())
                setGraphic(m_xICOfControlDialog, "/images/gradient_mode.png");
            else if(getController().getDiagram().isBlueGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/blueGradient_mode.png");
            else if(getController().getDiagram().isRedGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/redGradient_mode.png");
            else
                setGraphic(m_xICOfControlDialog, "");
        }

        if(getController().getDiagramType() == Controller.VENNDIAGRAM){
            if(getController().getDiagram().isBaseColorsProps())
                setGraphic(m_xICOfControlDialog, "/images/basecolors_mode.png");
            else
                setGraphic(m_xICOfControlDialog, "");
        }

        if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
            if(getController().getDiagram().isBlueGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/blueGradient_mode.png");
            else if(getController().getDiagram().isRedGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/redGradient_mode.png");
            else if(getController().getDiagram().isBaseColorsProps())
                setGraphic(m_xICOfControlDialog, "/images/basecolors_mode.png");
            else
                setGraphic(m_xICOfControlDialog, "");
        }

        if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
            if(getController().getDiagram().isBlueGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/blueGradient_mode.png");
            else if(getController().getDiagram().isRedGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/redGradient_mode.png");
            else if(getController().getDiagram().isBaseColorsWithGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/basecolorsWithGradients_mode.png");
            else if(getController().getDiagram().isBaseColorsProps())
                setGraphic(m_xICOfControlDialog, "/images/basecolors_mode.png");
            else
                setGraphic(m_xICOfControlDialog, "");
        }

        if(getController().getDiagramType() == Controller.TARGETDIAGRAM){
            if(getController().getDiagram().isBlueGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/blueGradient_mode.png");
            else if(getController().getDiagram().isRedGradientsProps())
                setGraphic(m_xICOfControlDialog, "/images/redGradient_mode.png");
            else if(getController().getDiagram().isBaseColorsProps())
                setGraphic(m_xICOfControlDialog, "/images/basecolors_mode.png");
            else
                setGraphic(m_xICOfControlDialog, "");
        }
    }

    public void setImageColorOfControlDialog(int color){
        setImageColorOfControl(m_xICOfControlDialog, color);
    }

    public int getImageColorOfControlDialog(){
        return getImageColorOfControl(m_xICOfControlDialog);
    }
 
    public void enableVisibleTextFieldOfControlDialog(boolean bool){
        enableVisibleControl((XControl) UnoRuntime.queryInterface(XControl.class, m_xTFOfControlDialog), bool);
    }

    public void enableTextFieldOfControlDialog(boolean bool){
        enableControl((XControl) UnoRuntime.queryInterface(XControl.class, m_xTFOfControlDialog), bool);
        if(!bool)
            setFocusControlDialog();
    }

    public boolean isShownTextFieldOfControlDialog(){
        if(m_xTFOfControlDialog != null)
            return m_isShownTFOfControlDialog;
        return false;
    }

    public void moveTextFieldOfControlDialog(){
        try {
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, m_xControlDialog);
            XPropertySet xPS = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());

            String sPackageURL = getPackageLocation();
            m_isShownTFOfControlDialog = !m_isShownTFOfControlDialog;

            int dialogWidth = 258 + 3;
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                dialogWidth = 224 + 3;

            if(m_isShownTFOfControlDialog){
                enableVisibleTextFieldOfControlDialog(true);
                if(getController().isOnlySimpleItemIsSelected()){
                    setTextFieldOfControlDialog();
                    setFocusTextFieldOfControlDialog();
                } else
                    enableTextFieldOfControlDialog(false);
                xPS.setPropertyValue("Width", new Integer(dialogWidth));
                xPS.setPropertyValue("Height", new Integer(39));
                if(m_oUpAndDownBOfControlDialog != null)
                    setImageOfObject(m_oUpAndDownBOfControlDialog, sPackageURL + "/images/up.png", (short)-1);
            }else{
                enableVisibleTextFieldOfControlDialog(false);
                xPS.setPropertyValue("Width", new Integer(dialogWidth));
                xPS.setPropertyValue("Height", new Integer(16));
                if(m_oUpAndDownBOfControlDialog != null)
                    setImageOfObject(m_oUpAndDownBOfControlDialog, sPackageURL + "/images/down.png", (short)-1);
            }
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void modifiedTextFieldOfControlDialog(){
        if(m_xTFOfControlDialog != null)
            getController().setTextOfSelectedShape(m_xTFOfControlDialog.getText());
    }

    public String getTextFieldValueControlDialog(){
        if(m_xTFOfControlDialog != null)
            return m_xTFOfControlDialog.getText();
        return "";
    }

    public void setFocusTextFieldOfControlDialog(){
        if(m_xTFOfControlDialog != null){
            XWindow xWindow = (XWindow)UnoRuntime.queryInterface(XWindow.class, m_xTFOfControlDialog);
            if(xWindow != null)
                xWindow.setFocus();
        }
    }

    public void setTextFieldValueOfControlDialog(String str){
        m_xTFOfControlDialog.setText(str);
    }

    public void setTextFieldOfControlDialog(){
        if(m_xTFOfControlDialog != null){
            if(getController().isOnlySimpleItemIsSelected()){
                XText xText = null;
                if(getController().getGroupType() == Controller.RELATIONGROUP){
                    RelationDiagramItem item = ((RelationDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                    if(item != null){
                        XShape xTextShape = item.getTextShape();
                        if(xTextShape != null && getController().getDiagram().isInGruopShapes(xTextShape)){
                            xText = (XText)UnoRuntime.queryInterface(XText.class, xTextShape);
                            setTextFieldValueOfControlDialog(xText.getString());
                        }else
                            setTextFieldValueOfControlDialog("");
                    }
                }else{
                    xText = (XText)UnoRuntime.queryInterface(XText.class, getController().getSelectedShape());
                    setTextFieldValueOfControlDialog(xText.getString());
                }
            }else{
                String message = getDialogPropertyValue("Strings2", "Strings2.InactiveFunction.Label");
                setTextFieldValueOfControlDialog(message);
            }
        }
    }
    //*******************************************************************************************************


    //ColorTable's methods *********************************************************************************
    public void executeColorTable() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ColorTable.xdl";
            m_xColorTableDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            addListenerAndExecuteColorTable();
        } catch (IllegalArgumentException e) {
            System.err.println(e.getLocalizedMessage());
        }
    }

    public void executeColorTable2() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ColorTable2.xdl";
            m_xColorTableDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            addListenerAndExecuteColorTable();
        } catch (IllegalArgumentException e) {
            System.err.println(e.getLocalizedMessage());
        }
    }

    public void executeColorTable3() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ColorTable3.xdl";
            m_xColorTableDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if(m_xColorTableDialog != null){
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
                XCheckBox xBaseColorCheckBox = (XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("baseColorsModeCheckBox"));
                if(getController().getDiagram() != null && !getController().getDiagram().isBaseColorsProps())
                    xBaseColorCheckBox.setState((short)0);
                if(getController().getDiagramType() == Controller.CYCLEDIAGRAM || getController().getDiagramType() == Controller.PYRAMIDDIAGRAM || getController().getDiagramType() == Controller.TARGETDIAGRAM){
                    if(getController().getDiagram().isBlueGradientsProps() || getController().getDiagram().isRedGradientsProps())
                        enableImageControlsOfColorTable(false);
                }
                addListenerAndExecuteColorTable();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void enableImageControlsOfColorTable(boolean bool){
        if(m_xColorTableDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
            if(xControlContainer != null){
                for(int i = 1; i <= 92; i++){
                    XControl xControl = xControlContainer.getControl("ImageControl" + i);
                    enableControl(xControl, bool);
                }
            }
        }
    }

    public void addListenerAndExecuteColorTable() {
        if (m_xColorTableDialog != null) {
            m_xColorTableTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xColorTableDialog);
            m_xColorTableTopWindow.addTopWindowListener(m_oListener);
            enableControlDialogWindow(false);
            m_xColorTableDialog.execute();
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xColorTableDialog);
            if(xComp != null)
                xComp.dispose();
        }
    }

    public void setResultColorTable(boolean isAction, String methodName){
        int color = -1;

        if(isAction && !methodName.equals(""))
            color = getCurrColorOfColorTable(getNum(methodName));
        else
            isAction = false;
        if(color == -1)
            isAction = false;

        if(m_xGradientDialogWindow != null){
            if(isAction){
                if(m_sImageType.equals("StartImage"))
                    setImageColorOfControl(m_xStartImageOfGradientDialog, color);
                if(m_sImageType.equals("EndImage"))
                    setImageColorOfControl(m_xEndImageOfGradientDialog, color);
            }
            m_xGradientDialogWindow.setFocus();
        }else if(m_xPropsDialogWindow != null && isShownPropsDialog()){
            if(isAction)
                if(m_xEventObjectControl != null)
                    setImageColorOfControl(m_xEventObjectControl, color);
            m_xPropsDialogWindow.setFocus();
        }else{
            if(isAction){
                setImageColorOfControlDialog(color);
                getController().getDiagram().setColorProps(color);
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    setBaseColorsProps();
                //if(getController().getGroupType() == Controller.ORGANIGROUP){
                //    getController().getDiagram().setGradientProps(false);
                //    getController().getDiagram().setBlueGradientsProps(false);
                //    getController().getDiagram().setRedGradientsProps(false);
                //    setColorModeOfImageOfControlDialog();
                //}
            }
            enableControlDialogWindow(true);
            setFocusControlDialog();
        }
    }

    public void setBaseColorsProps(){
        if(isBaseColorModeOfColorTable()){
            getController().getDiagram().setBaseColorsProps(true);
            getController().getDiagram().setBlueGradientsProps(false);
            getController().getDiagram().setRedGradientsProps(false);
        }else{
            getController().getDiagram().setBaseColorsProps(false);
            getController().getDiagram().setBaseColorsWithGradientsProps(false);
        }
        setColorModeOfImageOfControlDialog();
    }

    public boolean isBaseColorModeOfColorTable(){
        if(m_xColorTableDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
            XCheckBox xBaseColorCheckBox = (XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("baseColorsModeCheckBox"));
            if(xBaseColorCheckBox != null)
                if(xBaseColorCheckBox.getState() == 1)
                    return true;
        }
        return false;
    }

    public void endExecuteColorTable(){
        if(m_xColorTableDialog != null)
            m_xColorTableDialog.endExecute();
    }

    public int getCurrColorOfColorTable(int i){
        int color = -1;
        try {
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
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

    public void disposeColorTable(){
        if(m_xColorTableTopWindow != null)
            m_xColorTableTopWindow.removeTopWindowListener(m_oListener);
        m_xColorTableTopWindow = null;
        m_xColorTableDialog = null;
    }
    //********************************************************************************************************


    //GradientDialog's methods *******************************************************************************
    public void executeGradientDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/GradientDialog.xdl";
            m_xGradientDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xGradientDialog != null) {
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
                m_xStartImageOfGradientDialog = xControlContainer.getControl("startColor");
                m_xEndImageOfGradientDialog = xControlContainer.getControl("endColor");
                setImageColorOfControl(m_xStartImageOfGradientDialog, getController().getDiagram().getStartColorProps());
                setImageColorOfControl(m_xEndImageOfGradientDialog, getController().getDiagram().getEndColorProps());
                XListBox xListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("gradientModeListBox"));
                xListBox.selectItemPos(getController().getDiagram().getGradientDirectionProps(), true);

                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    if(getController().getDiagram().isBlueGradientsProps() || getController().getDiagram().isRedGradientsProps()){
                        XRadioButton xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("setColorModeOptionButton"));
                        xRadioButton.setState(true);
                        enableControl(xControlContainer.getControl("setGradientOptionButton"), false);
                        enableGradientDialogControls(false);
                    }
                }

                m_xGradientDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xGradientDialog);
                m_xGradientDialogTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xGradientDialog);
                m_xGradientDialogTopWindow.addTopWindowListener(m_oListener);
                enableControlDialogWindow(false);
                m_xGradientDialog.execute();
                XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xGradientDialog);
                if(xComp != null)
                    xComp.dispose();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void enableGradientDialogControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
        enableControl(xControlContainer.getControl("label1"), bool);
        enableControl(m_xStartImageOfGradientDialog, bool);
        enableControl(xControlContainer.getControl("label2"), bool);
        enableControl(m_xEndImageOfGradientDialog, bool);
        enableControl(xControlContainer.getControl("gradientModeListBox"), bool);
        if(bool){
            setImageColorOfControl(m_xStartImageOfGradientDialog, getController().getDiagram().getStartColorProps());
            setImageColorOfControl(m_xEndImageOfGradientDialog, getController().getDiagram().getEndColorProps());
        }
    }

    public int getColorOfStartImageOfGradientDialog(){
        if (m_xGradientDialog != null)
            return getImageColorOfControl(m_xStartImageOfGradientDialog);
        return -1;
    }

    public int getColorOfEndImageOfGradientDialog(){
        if (m_xGradientDialog != null)
            return getImageColorOfControl(m_xEndImageOfGradientDialog);
        return -1;
    }

    public void disposeGradientDialog(){
        if(m_xGradientDialog != null){
            m_xGradientDialogTopWindow.removeTopWindowListener(m_oListener);
            m_xGradientDialog.endExecute();
            m_xGradientDialogWindow = null;
            m_xGradientDialogTopWindow = null;
            m_xGradientDialog = null;
            enableControlDialogWindow(true);
            setFocusControlDialog();
        }

    }

    public void setGradientProps(){
        getController().getDiagram().setBlueGradientsProps(false);
        getController().getDiagram().setRedGradientsProps(false);
        if(isGradientModeOfGradientDialog())
            getController().getDiagram().setGradientProps(true);
        else
            getController().getDiagram().setGradientProps(false);
        setColorModeOfImageOfControlDialog();
    }

    public boolean isGradientModeOfGradientDialog(){
        if(m_xGradientDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
            if(xControlContainer != null){
                XRadioButton xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("setColorModeOptionButton"));
                if(xRadioButton != null)
                    return !xRadioButton.getState();
            }
        }
        return false;
    }

    //********************************************************************************************************


    //ConvertDialog's methods ********************************************************************************
    public void executeConvertDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL;
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                sDialogURL += "/dialogs/OrganigroupConvertDialog.xdl";
            if(getController().getGroupType() == Controller.RELATIONGROUP)
                sDialogURL += "/dialogs/RelationgroupConvertDialog.xdl";
            m_xConvertDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xConvertDialog != null){
                setConvertDialogControls();
                //m_xConvertDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xConvertDialog);
                m_xConvertDialogTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xConvertDialog);
                m_xConvertDialogTopWindow.addTopWindowListener(m_oListener);
                enableControlDialogWindow(false);
                m_xConvertDialog.execute();
                XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xConvertDialog);
                if(xComp != null)
                    xComp.dispose();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setConvertDialogControls(){
        if(m_xConvertDialog != null){
            XRadioButton xRadioButton2          = null;
            XControl xControl                   = null;
            String disabledOptionButton         = "convertOptionButton";
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            if(xControlContainer != null){

                if(getController().getGroupType() == Controller.ORGANIGROUP){
                    if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM){
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 1;
                    }
                    if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                        disabledOptionButton += 2;
                    if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                        disabledOptionButton += 3;
                    //if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    //    disabledOptionButton += 4;
                    if(getController().getDiagramType() != Controller.ORGANIGRAM){
                        xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl(disabledOptionButton));
                        enableControl(xControl, false);
                    }else{
                        m_xListBoxOfConvertDialog = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lastHorLevelListBox"));
                        m_xListBoxOfConvertDialog.selectItemPos(OrgChartTree.LASTHORLEVEL, true);
                    }
                }

                if(getController().getGroupType() == Controller.RELATIONGROUP){
                    if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 1;
                    }
                    if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                        disabledOptionButton += 2;
                    if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                        disabledOptionButton += 3;
                    if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                        disabledOptionButton += 4;
                    xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl(disabledOptionButton));
                    enableControl(xControl, false);
                }
            }
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
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    return Controller.VENNDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.HORIZONTALORGANIGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    return Controller.CYCLEDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton3"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.TABLEHIERARCHYDIAGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    return Controller.PYRAMIDDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton4"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.ORGANIGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    return Controller.TARGETDIAGRAM;
            }
        }
        return -1;
    }

    public void enableListBoxOfConvertDialog(boolean bool){
        try {
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl("lastHorLevelListBox"));
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
    
    public void endExecuteConvertDialog(){
        if(m_xConvertDialog != null)
            m_xConvertDialog.endExecute();
    }

    public void disposeConvertDialog(){
        if(m_xConvertDialogTopWindow != null)
            m_xConvertDialogTopWindow.removeTopWindowListener(m_oListener);
        //m_xConvertDialogWindow = null;
        m_xConvertDialogTopWindow = null;
        m_xConvertDialog = null;
        enableControlDialogWindow(true);
        setFocusControlDialog();
    }
    //********************************************************************************************************


    //Properties dialogs' methods ****************************************************************************
    public boolean isShownPropsDialog(){
        return m_IsShownPropsDialog;
    }

    public void setShownPropsDialog(boolean bool){
        m_IsShownPropsDialog = bool;
    }

     public void executePropertiesDialog() {
        if(m_xPropsDialog == null){
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
                if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                    diagramDefine = "TargetDiagramPropsDialog.xdl";
                String sDialogURL = sPackageURL + "/dialogs/" + diagramDefine;
                m_xPropsDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
                if (m_xPropsDialog != null) {
                    m_xPropsDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xPropsDialog);
                    m_xPropsDialogTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xPropsDialog);
                    XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        m_xModifyColorsCheckBoxOfPD = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"));
                        m_xColorOBControlOfPD = xControlContainer.getControl("colorOptionButton");
                        m_xColorImageControlOfPD = xControlContainer.getControl("colorImageControl");
                        m_xGradientsOBControlOfPD = xControlContainer.getControl("gradientsOptionButton");
                        m_xGradientModeLBControlOfPD = xControlContainer.getControl("gradientModeListBox");
                        m_xStartColorLabelControlOfPD = xControlContainer.getControl("label0");
                        m_xStartColorImageControlOfPD = xControlContainer.getControl("startColorImageControl");
                        m_xEndColorImageControlOfPD = xControlContainer.getControl("endColorImageControl");
                        m_xEndColorLabelControlOfPD = xControlContainer.getControl("label1");

                        m_xAreaLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label7"));
                        m_xModifiesColorLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label8"));
                        m_xColorModeLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label9"));
                        m_xRoundedLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label10"));
                        m_xOutlineLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label11"));
                        m_xShadowLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label12"));
                    }

                    if(getController().getDiagramType() == Controller.VENNDIAGRAM){
                        m_xModifyColorsCheckBoxOfPD = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"));
                        m_xBaseColorOBControlOfPD = xControlContainer.getControl("baseColorOptionButton");
                        m_xBaseColorRadioButtonOfPD = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorOBControlOfPD);
                        m_xColorOBControlOfPD = xControlContainer.getControl("colorOptionButton");
                        m_xColorImageControlOfPD = xControlContainer.getControl("colorImageControl");

                        m_xYesOutlineOBControlOfPD = xControlContainer.getControl("yesOutlineOptionButton");
                        m_xNoOutlineOBControlOfPD = xControlContainer.getControl("noOutlineOptionButton");
                        m_xYesFrameOBControlOfPD = xControlContainer.getControl("yesFrameOptionButton");
                        m_xNoFrameOBControlOfPD = xControlContainer.getControl("noFrameOptionButton");
                        m_xYesFrameRoundedOBControlOfPD = xControlContainer.getControl("yesFrameRoundedOptionButton");
                        m_xNoFrameRoundedOBControlOfPD = xControlContainer.getControl("noFrameRoundedOptionButton");

                        m_xAreaLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label8"));
                        m_xModifiesColorLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label9"));
                        m_xColorModeLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label10"));
                        m_xTransparencyLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label11"));
                        m_xOutlineLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label12"));
                        m_xTFrameLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label13"));
                        m_xRoundedTFrameLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label14"));
                    }

                    if(getController().getDiagramType() == Controller.CYCLEDIAGRAM){
                        m_xModifyColorsCheckBoxOfPD = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"));
                        m_xBaseColorOBControlOfPD = xControlContainer.getControl("baseColorOptionButton");
                        m_xBaseColorRadioButtonOfPD = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorOBControlOfPD);
                        m_xColorOBControlOfPD = xControlContainer.getControl("colorOptionButton");
                        m_xColorImageControlOfPD = xControlContainer.getControl("colorImageControl");
                        m_xYesShadowOBOfPD = xControlContainer.getControl("yesShadowOptionButton");
                        m_xNoShadowOBOfPD = xControlContainer.getControl("noShadowOptionButton");

                        m_xAreaLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label7"));
                        m_xModifiesColorLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label8"));
                        m_xColorModeLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label9"));
                        m_xOutlineLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label10"));
                        m_xTFrameLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label11"));
                        m_xShadowLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label12"));
                    }

                    if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
                        m_xModifyColorsCheckBoxOfPD = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"));
                        m_xBaseColorOBControlOfPD = xControlContainer.getControl("baseColorOptionButton");
                        m_xBaseColorRadioButtonOfPD = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorOBControlOfPD);
                        m_xColorOBControlOfPD = xControlContainer.getControl("colorOptionButton");
                        m_xColorImageControlOfPD = xControlContainer.getControl("colorImageControl");

                        m_xAreaLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label6"));
                        m_xModifiesColorLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label7"));
                        m_xColorModeLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label8"));
                        m_xOutlineLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label9"));
                        m_xShadowLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label10"));
                    }

                    if(getController().getDiagramType() == Controller.TARGETDIAGRAM){
                        m_xModifyColorsCheckBoxOfPD = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"));
                        m_xBaseColorOBControlOfPD = xControlContainer.getControl("baseColorOptionButton");
                        m_xBaseColorRadioButtonOfPD = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorOBControlOfPD);
                        m_xColorOBControlOfPD = xControlContainer.getControl("colorOptionButton");
                        m_xColorImageControlOfPD = xControlContainer.getControl("colorImageControl");

                        m_xYesOutlineOBControlOfPD = xControlContainer.getControl("yesOutlineOptionButton");
                        m_xNoOutlineOBControlOfPD = xControlContainer.getControl("noOutlineOptionButton");
                        m_xYesFrameOBControlOfPD = xControlContainer.getControl("yesFrameOptionButton");
                        m_xNoFrameOBControlOfPD = xControlContainer.getControl("noFrameOptionButton");

                        m_xAreaLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label7"));
                        m_xModifiesColorLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label8"));
                        m_xColorModeLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label9"));
                        m_xLayoutLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label10"));
                        m_xOutlineLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label11"));
                        m_xTFrameLabelOfPD = (XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label12"));
                    }

                    setControlScaleImageProp(xControlContainer.getControl("imageControl"), false);
                    
                    getController().getDiagram().setFontPropertyValues();
                }
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        if(m_xPropsDialog != null){

            setTextOptionControls();

            if(getController().getGroupType() == Controller.ORGANIGROUP)
                setOrganigramPropsDialog();
            if(getController().getDiagramType() == Controller.VENNDIAGRAM)
                setVennDiagramPropsDialog();
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                setCycleDiagramPropsDialog();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                setPyramidDiagramPropsDialog();
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                setTargetDiagramPropsDialog();

            m_xPropsDialogTopWindow.addTopWindowListener(m_oListener);
            enableControlDialogWindow(false);
            setShownPropsDialog(true);
            m_xPropsDialog.execute();
        }
    }

    public void setTextOptionControls(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            boolean isFitText = getController().getDiagram().isTextFitProps();
            if(isFitText){
                XRadioButton textFitRB = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("textFitOptionButton"));
                textFitRB.setState(true);
            }else{
                XRadioButton fontSizeRB = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("fontSizeOptionButton"));
                fontSizeRB.setState(true);
                setMarkInFontSizeLB();
            }
            enableFontSizeListBox(!isFitText);
        }
    }

    public void enableFontSizeListBox(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("fontSizeListBox"), bool);
        }
    }

    public void setMarkInFontSizeLB(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox fontSizeLB = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("fontSizeListBox"));

            //short selectedItemPos = fontSizeLB.getSelectedItemPos();
            short index = FontSize._getIndexOfFontSize(getController().getDiagram().getFontSizeProps());
            fontSizeLB.selectItemPos(index, true);
            String label = fontSizeLB.getSelectedItem();
            if(!label.startsWith("*"))
                label = "*" + label.substring(1);
            fontSizeLB.removeItems(index, (short)1);
            fontSizeLB.addItem(label, index);
            fontSizeLB.selectItemPos(index, true);
        }
    }

    public void setNoMarkInFontSizeLB(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox fontSizeLB = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("fontSizeListBox"));
            short index = FontSize._getIndexOfFontSize(getController().getDiagram().getFontSizeProps());
            fontSizeLB.selectItemPos(index, true);
            String label = fontSizeLB.getSelectedItem();
            if(label.startsWith("*"))
                label = " " + label.substring(1);
            fontSizeLB.removeItems(index, (short)1);
            fontSizeLB.addItem(label, index);
            fontSizeLB.selectItemPos(index, true);
        }
    }

    public void setTextProperties(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XRadioButton textFitRB = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("textFitOptionButton"));
            boolean isFitText = textFitRB.getState();
            getController().getDiagram().setTextFitProps(isFitText);
            if(isFitText){
                setNoMarkInFontSizeLB();
            } else {
                XListBox fontSizeLB = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("fontSizeListBox"));
                float newFontSize = FontSize._getFontSize(fontSizeLB.getSelectedItemPos());
                setNoMarkInFontSizeLB();
                getController().getDiagram().setFontSizeProps(newFontSize);    
            }
            
            boolean isTextColorChange = isCheckedModifyTextColorCheckBox();
            getController().getDiagram().setTextColorChange(isTextColorChange);
            if(isTextColorChange)
                getController().getDiagram().setTextColorProps(getImageColorOfControl(getTextColorImageControl()));
         }
    }

    public XControl getTextColorImageControl(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return xControlContainer.getControl("textColorImageControl");
        }
        return null;
    }

    public void enableTextColorImageControl(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("textColorImageControl"), bool);
        }
    }
    
    public void enableTextColorLabel(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("textColorLabel"), bool);
        }
    }



    public boolean isCheckedModifyTextColorCheckBox(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            short state = ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).getState();
            if(state==1)
                return true;
        }
        return false;
    }

    public void setTextColorToolsProps(){
        boolean state = isCheckedModifyTextColorCheckBox();
        enableTextColorImageControl(state);
        enableTextColorLabel(state);
    }

    public void setDefaultTextColorToolsProps(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).setState((short)0);
            enableTextColorImageControl(false);
            enableTextColorLabel(false);
        }
    }

    public void endExecutePropertiesDialog(){
        if(m_xPropsDialogTopWindow != null)
                m_xPropsDialogTopWindow.removeTopWindowListener(m_oListener);
            m_xPropsDialog.endExecute();
            setShownPropsDialog(false);
            enableControlDialogWindow(true);
            setFocusControlDialog();
            setDefaultTextColorToolsProps();
    }

    public void disposePropertiesDialog(){
        if(m_xPropsDialog != null){
            if(m_xPropsDialogTopWindow != null)
                m_xPropsDialogTopWindow.removeTopWindowListener(m_oListener);
            m_xPropsDialog.endExecute();
            setShownPropsDialog(false);
            XComponent xComponent = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xPropsDialog);
            if(xComponent != null)
                xComponent.dispose();
            m_xPropsDialogTopWindow = null;
            m_xPropsDialog = null;
            enableControlDialogWindow(true);
            setFocusControlDialog();
        }
    }
     
    public void setOrganigramPropsDialog(){

        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        short style = getController().getDiagram().getStyle();

        String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
        String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
        String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
        String localBlueGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBlueGradients.Label");
        String localRedGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalRedGradients.Label");

        if(style != OrganizationChart.USER_DEFINE){
            enableVisibleOrganigramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfOrganigroupPD(true);
            if(style == OrganizationChart.DEFAULT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_default.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_default.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_default.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_default.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localYes, localYes, localNo);
                enableColorFunctionFieldFieldOfOrganigroupPD(false);
            }
            if(style == OrganizationChart.WITHOUT_OUTLINE){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_withoutOutline.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_withoutOutline.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_withoutOutline.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_withoutOutline.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localYes, localNo, localNo);
                enableColorFunctionFieldFieldOfOrganigroupPD(false);
            }
            if(style == OrganizationChart.NOT_ROUNDED){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_notRounded.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_notRounded.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_notRounded.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_notRounded.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localNo, localYes, localNo);
                enableColorFunctionFieldFieldOfOrganigroupPD(false);
            }
            if(style == OrganizationChart.WITH_SHADOW){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_withShadow.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_withShadow.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_withShadow.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_withShadow.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localYes, localYes, localYes);
                enableColorFunctionFieldFieldOfOrganigroupPD(false);
            }
            if(style == OrganizationChart.BLUE_GRADIENTS){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_blueGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_blueGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_blueGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_blueGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localBlueGradients, localYes, localYes, localNo);
            }
            if(style == OrganizationChart.RED_GRADIENTS){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_redGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_redGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_redGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_redGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localRedGradients, localYes, localYes, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleOrganigramPropsControls(true);
            enablePropertiesFieldOfOrganigroupPD(false);
        }
    }
    
    public void enablePropertiesFieldOfOrganigroupPD(boolean bool){
        if(m_xPropsDialog != null){
            if(bool == false){
                if(m_xAreaLabelOfPD != null)
                    m_xAreaLabelOfPD.setText("");
                if(m_xModifiesColorLabelOfPD != null)
                    m_xModifiesColorLabelOfPD.setText("");
                if(m_xColorModeLabelOfPD != null)
                    m_xColorModeLabelOfPD.setText("");
                if(m_xRoundedLabelOfPD != null)
                    m_xRoundedLabelOfPD.setText("");
                if(m_xOutlineLabelOfPD != null)
                    m_xOutlineLabelOfPD.setText("");
                if(m_xShadowLabelOfPD != null)
                    m_xShadowLabelOfPD.setText("");
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("Label" + i), bool);
        }
    }

    public void enableColorFunctionFieldFieldOfOrganigroupPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("Label3"), bool);
        }
    }

    public  void setDescriptionLabelOfOrganigroupPD(String area, String modifiesColor, String colorMode, String rounded, String outline, String shadow){
        if(m_xAreaLabelOfPD != null)
            m_xAreaLabelOfPD.setText(area);
        if(m_xModifiesColorLabelOfPD != null)
            m_xModifiesColorLabelOfPD.setText(modifiesColor);
        if(m_xColorModeLabelOfPD != null)
            m_xColorModeLabelOfPD.setText(colorMode);
        if(m_xRoundedLabelOfPD != null)
            m_xRoundedLabelOfPD.setText(rounded);
        if(m_xOutlineLabelOfPD != null)
            m_xOutlineLabelOfPD.setText(outline);
        if(m_xShadowLabelOfPD != null)
            m_xShadowLabelOfPD.setText(shadow);
    }

    public void enableVisibleOrganigramPropsControls(boolean bool){

        XRadioButton xRadioButton = null;
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            getController().getDiagram().setSelectedAllShapesProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);

        getController().getDiagram().setModifyColorsProps(false);
        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);

        enableVisibleControl(m_xColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorImageControlOfPD, bool);
        enableVisibleControl(m_xGradientsOBControlOfPD, bool);
        enableVisibleControl(m_xGradientModeLBControlOfPD, bool);
        enableVisibleControl(m_xStartColorLabelControlOfPD, bool);
        enableVisibleControl(m_xStartColorImageControlOfPD, bool);
        enableVisibleControl(m_xEndColorImageControlOfPD, bool);
        enableVisibleControl(m_xEndColorLabelControlOfPD, bool);

        if(bool){
            setImageColorOfControl(m_xColorImageControlOfPD, getController().getDiagram().getColorProps());
            setImageColorOfControl(m_xStartColorImageControlOfPD, getController().getDiagram().getStartColorProps());
            setImageColorOfControl(m_xEndColorImageControlOfPD, getController().getDiagram().getEndColorProps());
            XListBox xListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, m_xGradientModeLBControlOfPD);
            xListBox.selectItemPos(getController().getDiagram().getGradientDirectionProps(), true);

            if(m_xModifyColorsCheckBoxOfPD.getState() == 0){
                enableOrganigramColorControls(false);
            }else{
                getController().getDiagram().setModifyColorsProps(true);
                enableControl(m_xColorOBControlOfPD, true);
                enableControl(m_xGradientsOBControlOfPD, true);
                xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xColorOBControlOfPD);
                boolean isBaseColorsMode = getController().getDiagram().isBaseColorsProps();
                xRadioButton.setState(isBaseColorsMode);
                //getController().getDiagram().setGradientProps(!state);
                enableOrganigramControlsIsColorProps(isBaseColorsMode);
            }
        }

        if(bool)
            getController().getDiagram().setRoundedProps((short)1);
        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("noRoundedOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noRoundedOptionButton"));
            if(xRadioButton.getState()==true)
                getController().getDiagram().setRoundedProps((short)0);
        }
        enableVisibleControl(xControlContainer.getControl("mediumRoundedOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("extraRoundedOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraRoundedOptionButton"));
            if(xRadioButton.getState()==true)
                getController().getDiagram().setRoundedProps((short)2);
        }

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"));
            getController().getDiagram().setOutlineProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);


        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"));
            getController().getDiagram().setShadowProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

    }

    public void enableOrganigramControlsIsColorProps(boolean bool){
        enableControl(m_xColorImageControlOfPD, bool);
        enableControl(m_xGradientModeLBControlOfPD, !bool);
        enableControl(m_xStartColorLabelControlOfPD, !bool);
        enableControl(m_xStartColorImageControlOfPD, !bool);
        enableControl(m_xEndColorImageControlOfPD, !bool);
        enableControl(m_xEndColorLabelControlOfPD, !bool);
    }

    public void enableOrganigramColorControls(boolean bool){
        enableControl(m_xColorOBControlOfPD, bool);
        enableControl(m_xColorImageControlOfPD, bool);
        enableControl(m_xGradientsOBControlOfPD, bool);
        enableControl(m_xGradientModeLBControlOfPD, bool);
        enableControl(m_xStartColorLabelControlOfPD, bool);
        enableControl(m_xStartColorImageControlOfPD, bool);
        enableControl(m_xEndColorLabelControlOfPD, bool);
        enableControl(m_xEndColorImageControlOfPD, bool);
    }

    public boolean isGradientPropsInDiagramPropsDialog(){
        if(m_xPropsDialog != null){
            XRadioButton xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xGradientsOBControlOfPD);
            if(xRadioButton != null)
                return xRadioButton.getState();
        }
        return false;
    }

    public void setVennDiagramPropsDialog(){

        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        short style = getController().getDiagram().getStyle();

        String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
        String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
        String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
        String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
        String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");

        if(style != VennDiagram.USER_DEFINE){
            enableVisibleVennDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfVennDiagramPD(true);
            
            if(style == VennDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/venn_default.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localBaseColors, localMedium, localYes, localYes, localYes);
            }
            if (style == VennDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/venn_withoutOutline.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localBaseColors, localMedium, localNo, localYes, localYes);
            }
            if (style == VennDiagram.WITHOUT_FRAME){
                setGraphic(xImageControl, "/images/venn_withoutFrame.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localBaseColors, localMedium, localYes, localNo, "");
                enableRoundedTextFieldFieldOfVennDiagramPD(false);
            }
            if (style == VennDiagram.NOT_ROUNDED){
                setGraphic(xImageControl, "/images/venn_notRounded.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localBaseColors, localMedium, localYes, localYes, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleVennDiagramPropsControls(true);
            enablePropertiesFieldOfVennDiagramPD(false);
        }
    }

    public void enablePropertiesFieldOfVennDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            if(bool == false){
                if(m_xAreaLabelOfPD != null)
                    m_xAreaLabelOfPD.setText("");
                if(m_xColorModeLabelOfPD != null)
                    m_xColorModeLabelOfPD.setText("");
                if(m_xModifiesColorLabelOfPD != null)
                    m_xModifiesColorLabelOfPD.setText("");
                if(m_xTransparencyLabelOfPD != null)
                    m_xTransparencyLabelOfPD.setText("");
                if(m_xOutlineLabelOfPD != null)
                    m_xOutlineLabelOfPD.setText("");
                if(m_xTFrameLabelOfPD != null)
                    m_xTFrameLabelOfPD.setText("");
                if(m_xRoundedTFrameLabelOfPD != null)
                    m_xRoundedTFrameLabelOfPD.setText("");
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 14; i++)
                enableControl(xControlContainer.getControl("Label" + i), bool);
        }
    }

    public void enableRoundedTextFieldFieldOfVennDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("Label7"), bool);
        }
    }

    public  void setDescriptionLabelOfVennDiagramPD(String area, String modifiesColor, String colorMode, String transp, String outline, String TFrame, String roundedTFrame){
        if(m_xAreaLabelOfPD != null)
            m_xAreaLabelOfPD.setText(area);
        if(m_xModifiesColorLabelOfPD != null)
            m_xModifiesColorLabelOfPD.setText(modifiesColor);
        if(m_xColorModeLabelOfPD != null)
            m_xColorModeLabelOfPD.setText(colorMode);
        if(m_xTransparencyLabelOfPD != null)
            m_xTransparencyLabelOfPD.setText(transp);
        if(m_xOutlineLabelOfPD != null)
            m_xOutlineLabelOfPD.setText(outline);
        if(m_xTFrameLabelOfPD != null)
            m_xTFrameLabelOfPD.setText(TFrame);
        if(m_xRoundedTFrameLabelOfPD != null)
            m_xRoundedTFrameLabelOfPD.setText(roundedTFrame);
    }

    public void enableVisibleVennDiagramPropsControls(boolean bool){

        XRadioButton xRadioButton = null;
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            getController().getDiagram().setSelectedAllShapesProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);

        getController().getDiagram().setModifyColorsProps(false);
        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("baseColorOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
        enableVisibleControl(m_xColorImageControlOfPD, bool);
        if(bool){
            XRadioButton xBaseColorsRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorOptionButton"));
            XRadioButton xColorRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"));
            if(getController().getDiagram().isBaseColorsProps())
                xBaseColorsRadioButton.setState(true);
            else
                xColorRadioButton.setState(true);
            if(m_xModifyColorsCheckBoxOfPD.getState() == 0){
                enableControl(xControlContainer.getControl("baseColorOptionButton"), false);
                enableControl(xControlContainer.getControl("colorOptionButton"), false);
                enableControl(m_xColorImageControlOfPD, false);
            }else{
                getController().getDiagram().setModifyColorsProps(true);
                enableControl(xControlContainer.getControl("baseColorOptionButton"), true);
                enableControl(xControlContainer.getControl("colorOptionButton"), true);
                xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorOptionButton"));
                if(xRadioButton.getState())
                    enableControl(m_xColorImageControlOfPD, false);
            }
        }

        if(bool)
            getController().getDiagram().setTransparencyProps(Diagram.MEDIUM_TRANSP);
        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("noTransparencyOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noTransparencyOptionButton"));
            if(xRadioButton.getState()==true)
                getController().getDiagram().setTransparencyProps(Diagram.NULL_TRANSP);
        }
        enableVisibleControl(xControlContainer.getControl("mediumTransparencyOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("extraTransparencyOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraTransparencyOptionButton"));
            if(xRadioButton.getState()==true)
                getController().getDiagram().setTransparencyProps(Diagram.EXTRA_TRANSP);
        }

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(m_xYesOutlineOBControlOfPD, bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xYesOutlineOBControlOfPD);
            getController().getDiagram().setOutlineProps(xRadioButton.getState());
        }
        enableVisibleControl(m_xNoOutlineOBControlOfPD, bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(m_xYesFrameOBControlOfPD, bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xYesFrameOBControlOfPD);
            getController().getDiagram().setFrameProps(xRadioButton.getState());
        }
        enableVisibleControl(m_xNoFrameOBControlOfPD, bool);

        enableVisibleControl(xControlContainer.getControl("frameControl6"), bool);
        enableVisibleControl(m_xYesFrameRoundedOBControlOfPD, bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xYesFrameRoundedOBControlOfPD);
            getController().getDiagram().setRoundedFrameProps(xRadioButton.getState());
        }
        enableVisibleControl(m_xNoFrameRoundedOBControlOfPD, bool);
    }

    public void enableVennDiagramNotAllShapeControls(boolean bool){
        enableControl(m_xYesOutlineOBControlOfPD, bool);
        enableControl(m_xNoOutlineOBControlOfPD, bool);
        enableControl(m_xYesFrameOBControlOfPD, bool);
        enableControl(m_xNoFrameOBControlOfPD, bool);
        enableControl(m_xNoFrameRoundedOBControlOfPD, bool);
        enableControl(m_xYesFrameRoundedOBControlOfPD, bool);
    }

    public void setCycleDiagramPropsDialog(){

        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        short style = getController().getDiagram().getStyle();

        String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
        String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
        String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
        String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
        String localBlueGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBlueGradients.Label");
        String localRedGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalRedGradients.Label");

        if(style != CycleDiagram.USER_DEFINE){
            enableVisibleCycleDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfCycleDiagramPD(true);

            if(style == CycleDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/ring_default.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo, localNo);
            }
            if(style == CycleDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/ring_withoutOutline.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localNo, localNo, localNo);
            }
            if(style == CycleDiagram.WITH_SHADOW){
                setGraphic(xImageControl, "/images/ring_withShadow.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo, localYes);
            }
            if(style == CycleDiagram.WITH_FRAME){
                setGraphic(xImageControl, "/images/ring_withFrame.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localYes, localNo);
            }
            if(style == CycleDiagram.BLUE_GRADIENTS){
                setGraphic(xImageControl, "/images/ring_blueGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBlueGradients, localYes, localNo, localNo);
            }
            if(style == CycleDiagram.RED_GRADIENTS){
                setGraphic(xImageControl, "/images/ring_redGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localRedGradients, localYes, localNo, localNo);
            }
        }
        if(style == CycleDiagram.USER_DEFINE){
            enableVisibleControl(xImageControl, false);
            enableVisibleCycleDiagramPropsControls(true);
            enablePropertiesFieldOfCycleDiagramPD(false);
        }
    }

    public void enablePropertiesFieldOfCycleDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            if(bool == false){
                if(m_xAreaLabelOfPD != null)
                    m_xAreaLabelOfPD.setText("");
                if(m_xModifiesColorLabelOfPD != null)
                    m_xModifiesColorLabelOfPD.setText("");
                if(m_xColorModeLabelOfPD != null)
                    m_xColorModeLabelOfPD.setText("");
                if(m_xOutlineLabelOfPD != null)
                    m_xOutlineLabelOfPD.setText("");
                if(m_xTFrameLabelOfPD != null)
                    m_xTFrameLabelOfPD.setText("");
                if(m_xShadowLabelOfPD != null)
                    m_xShadowLabelOfPD.setText("");
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("Label" + i), bool);
        }
    }
    
    public  void setDescriptionLabelOfCycleDiagramPD(String area, String modifiesColor, String colorMode, String outline, String textFrame, String shadow){
        if(m_xAreaLabelOfPD != null)
            m_xAreaLabelOfPD.setText(area);
        if(m_xModifiesColorLabelOfPD != null)
            m_xModifiesColorLabelOfPD.setText(modifiesColor);
        if(m_xColorModeLabelOfPD != null)
            m_xColorModeLabelOfPD.setText(colorMode);
        if(m_xOutlineLabelOfPD != null)
            m_xOutlineLabelOfPD.setText(outline);
        if(m_xTFrameLabelOfPD != null)
            m_xTFrameLabelOfPD.setText(textFrame);
        if(m_xShadowLabelOfPD != null)
            m_xShadowLabelOfPD.setText(shadow);
    }

    public void enableVisibleCycleDiagramPropsControls(boolean bool){
        XRadioButton xRadioButton = null;
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            getController().getDiagram().setSelectedAllShapesProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);

        getController().getDiagram().setModifyColorsProps(false);
        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(m_xBaseColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorImageControlOfPD, bool);
        if(bool){
            XRadioButton xColorRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xColorOBControlOfPD);
            if(getController().getDiagram().isBaseColorsProps())
                m_xBaseColorRadioButtonOfPD.setState(true);
            else
                xColorRadioButton.setState(true);
            if(m_xModifyColorsCheckBoxOfPD.getState() == 0){
                enableControl(m_xBaseColorOBControlOfPD, false);
                enableControl(m_xColorOBControlOfPD, false);
                enableControl(m_xColorImageControlOfPD, false);
            }else{
                getController().getDiagram().setModifyColorsProps(true);
                enableControl(m_xBaseColorOBControlOfPD, true);
                enableControl(m_xColorOBControlOfPD, true);
                if(m_xBaseColorRadioButtonOfPD.getState())
                    enableControl(m_xColorImageControlOfPD, false);
            }
        }

        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"));
            getController().getDiagram().setShadowProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"));
            getController().getDiagram().setOutlineProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesFrameOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"));
            getController().getDiagram().setFrameProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noFrameOptionButton"), bool);
       
    }

    public void setPyramidDiagramPropsDialog(){

        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        short style = getController().getDiagram().getStyle();

        String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
        String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
        String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
        String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
        String localBlueGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBlueGradients.Label");
        String localRedGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalRedGradients.Label");
        String localBaseColorsGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColorsGradients.Label");

        if(style != PyramidDiagram.USER_DEFINE){
            enableVisiblePyramidDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfPyramidDiagramPD(true);

            if(style == PyramidDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/pyramid_default.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo);
            }
            if(style == PyramidDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/pyramid_withoutOutline.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localNo, localNo);
            }
            if(style == PyramidDiagram.WITH_SHADOW){
                setGraphic(xImageControl, "/images/pyramid_withShadow.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localYes, localYes);
            }
            if(style == PyramidDiagram.BC_WITH_GRADIENTS){
                setGraphic(xImageControl, "/images/pyramid_basecolorsWithGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColorsGradients, localYes, localNo);
            }
            if(style == PyramidDiagram.BLUE_GRADIENTS){
                setGraphic(xImageControl, "/images/pyramid_blueGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBlueGradients, localYes, localNo);
            }
            if(style == PyramidDiagram.RED_GRADIENTS){
                setGraphic(xImageControl, "/images/pyramid_redGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localRedGradients, localYes, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisiblePyramidDiagramPropsControls(true);
            enablePropertiesFieldOfPyramidDiagramPD(false);
        }
    }

    public void enablePropertiesFieldOfPyramidDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            if(bool == false){
                if(m_xAreaLabelOfPD != null)
                    m_xAreaLabelOfPD.setText("");
                if(m_xModifiesColorLabelOfPD != null)
                    m_xModifiesColorLabelOfPD.setText("");
                if(m_xColorModeLabelOfPD != null)
                    m_xColorModeLabelOfPD.setText("");
                if(m_xOutlineLabelOfPD != null)
                    m_xOutlineLabelOfPD.setText("");
                if(m_xShadowLabelOfPD != null)
                    m_xShadowLabelOfPD.setText("");
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 10; i++)
                enableControl(xControlContainer.getControl("Label" + i), bool);
        }
    }

    public  void setDescriptionLabelOfPyramidDiagramPD(String area, String modifiesColor, String colorMode, String outline, String shadow){
        if(m_xAreaLabelOfPD != null)
            m_xAreaLabelOfPD.setText(area);
        if(m_xModifiesColorLabelOfPD != null)
            m_xModifiesColorLabelOfPD.setText(modifiesColor);
        if(m_xColorModeLabelOfPD != null)
            m_xColorModeLabelOfPD.setText(colorMode);
        if(m_xOutlineLabelOfPD != null)
            m_xOutlineLabelOfPD.setText(outline);
        if(m_xShadowLabelOfPD != null)
            m_xShadowLabelOfPD.setText(shadow);
    }

    public void enableVisiblePyramidDiagramPropsControls(boolean bool){
        XRadioButton xRadioButton = null;
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            getController().getDiagram().setSelectedAllShapesProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
        
        getController().getDiagram().setModifyColorsProps(false);
        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(m_xBaseColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorImageControlOfPD, bool);
        if(bool){
            XRadioButton xColorRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xColorOBControlOfPD);
            if(getController().getDiagram().isBaseColorsProps())
                m_xBaseColorRadioButtonOfPD.setState(true);
            else
                xColorRadioButton.setState(true);
            if(m_xModifyColorsCheckBoxOfPD.getState() == 0){
                enableControl(m_xBaseColorOBControlOfPD, false);
                enableControl(m_xColorOBControlOfPD, false);
                enableControl(m_xColorImageControlOfPD, false);
            }else{
                getController().getDiagram().setModifyColorsProps(true);
                enableControl(m_xBaseColorOBControlOfPD, true);
                enableControl(m_xColorOBControlOfPD, true);
                if(m_xBaseColorRadioButtonOfPD.getState())
                    enableControl(m_xColorImageControlOfPD, false);
            }
        }

        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"));
            getController().getDiagram().setOutlineProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"));
            getController().getDiagram().setShadowProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

    }

    public void setTargetDiagramPropsDialog(){

        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        short style = getController().getDiagram().getStyle();

        String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
        String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
        String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
        String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
        String localBlueGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBlueGradients.Label");
        String localRedGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalRedGradients.Label");
        String localCenter = getDialogPropertyValue("Strings2", "Strings2.Common.LocalCentert.Label");
        String localLeft = getDialogPropertyValue("Strings2", "Strings2.Common.LocalLeft.Label");

        if(style != TargetDiagram.USER_DEFINE){
            enableVisibleTargetDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfTargetDiagramPD(true);

            if(style == TargetDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/target_default.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localYes, localNo);
            }
            if (style == TargetDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/target_withoutOutline.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localNo, localNo);
            }
            if (style == TargetDiagram.WITH_FRAME){
                setGraphic(xImageControl, "/images/target_withTextFrame.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localYes, localYes);
            }
            if (style == TargetDiagram.BLUE_GRADIENTS){
                setGraphic(xImageControl, "/images/target_blueGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBlueGradients, localCenter, localYes, localNo);
            }
            if (style == TargetDiagram.RED_GRADIENTS){
                setGraphic(xImageControl, "/images/target_redGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localRedGradients, localCenter, localYes, localNo);
            }
            if (style == TargetDiagram.LEFT_LAYOUT){
                setGraphic(xImageControl, "/images/target_leftLayout.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localLeft, localYes, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleTargetDiagramPropsControls(true);
            enablePropertiesFieldOfTargetDiagramPD(false);
        }
    }

        public void enablePropertiesFieldOfTargetDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            if(bool == false){
                if(m_xAreaLabelOfPD != null)
                    m_xAreaLabelOfPD.setText("");
                if(m_xModifiesColorLabelOfPD != null)
                    m_xModifiesColorLabelOfPD.setText("");
                if(m_xColorModeLabelOfPD != null)
                    m_xColorModeLabelOfPD.setText("");
                if(m_xLayoutLabelOfPD != null)
                    m_xLayoutLabelOfPD.setText("");
                if(m_xOutlineLabelOfPD != null)
                    m_xOutlineLabelOfPD.setText("");
                if(m_xTFrameLabelOfPD != null)
                    m_xTFrameLabelOfPD.setText("");
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("Label" + i), bool);
        }
    }

    public  void setDescriptionLabelOfTargetDiagramPD(String area, String modifiesColor, String colorMode, String layout, String outline, String tFrame){
        if(m_xAreaLabelOfPD != null)
            m_xAreaLabelOfPD.setText(area);
        if(m_xModifiesColorLabelOfPD != null)
            m_xModifiesColorLabelOfPD.setText(modifiesColor);
        if(m_xColorModeLabelOfPD != null)
            m_xColorModeLabelOfPD.setText(colorMode);
        if(m_xLayoutLabelOfPD != null)
            m_xLayoutLabelOfPD.setText(layout);
        if(m_xOutlineLabelOfPD != null)
            m_xOutlineLabelOfPD.setText(outline);
        if(m_xTFrameLabelOfPD != null)
            m_xTFrameLabelOfPD.setText(tFrame);
    }

    public void enableVisibleTargetDiagramPropsControls(boolean bool){

        XRadioButton xRadioButton = null;
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            getController().getDiagram().setSelectedAllShapesProps(xRadioButton.getState());
        }
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);

        getController().getDiagram().setModifyColorsProps(false);
        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(m_xBaseColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorOBControlOfPD, bool);
        enableVisibleControl(m_xColorImageControlOfPD, bool);
        if(bool){
            XRadioButton xColorRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xColorOBControlOfPD);
            if(getController().getDiagram().isBaseColorsProps())
                m_xBaseColorRadioButtonOfPD.setState(true);
            else
                xColorRadioButton.setState(true);
            if(m_xModifyColorsCheckBoxOfPD.getState() == 0){
                enableControl(m_xBaseColorOBControlOfPD, false);
                enableControl(m_xColorOBControlOfPD, false);
                enableControl(m_xColorImageControlOfPD, false);
            }else{
                getController().getDiagram().setModifyColorsProps(true);
                enableControl(m_xBaseColorOBControlOfPD, true);
                enableControl(m_xColorOBControlOfPD, true);
                xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorOBControlOfPD);
                if(xRadioButton.getState())
                    enableControl(m_xColorImageControlOfPD, false);
            }
        }

        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("centerLayoutOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("leftLayoutOptionButton"), bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"));
            if(xRadioButton.getState()){
                if(((TargetDiagram)getController().getDiagram()).isLeftLayoutProperty()){
                    xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("leftLayoutOptionButton"));
                    xRadioButton.setState(true);
                } else{
                    xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("centerLayoutOptionButton"));
                    xRadioButton.setState(true);
                }
            }else{
                enableControl(xControlContainer.getControl("centerLayoutOptionButton"), false);
                enableControl(xControlContainer.getControl("leftLayoutOptionButton"), false);
            }
        }

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(m_xYesOutlineOBControlOfPD, bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xYesOutlineOBControlOfPD);
            getController().getDiagram().setOutlineProps(xRadioButton.getState());
        }
        enableVisibleControl(m_xNoOutlineOBControlOfPD, bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(m_xYesFrameOBControlOfPD, bool);
        if(bool){
            xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xYesFrameOBControlOfPD);
            getController().getDiagram().setFrameProps(xRadioButton.getState());
        }
        enableVisibleControl(m_xNoFrameOBControlOfPD, bool);
    }

    public boolean isLefLayoutInTargetDiagram(){
        XControlContainer xControlContainer = null;
        if(m_xPropsDialog != null)
            xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        if(xControlContainer != null){
            XRadioButton xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("leftLayoutOptionButton"));
            return xRadioButton.getState();
        }
        return false;
    }

    public void enableTargetDiagramNotAllShapeControls(boolean bool){
        XControlContainer xControlContainer = null;
        if(m_xPropsDialog != null)
            xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        if(xControlContainer != null){
            if(bool == false){
                XRadioButton xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("centerLayoutOptionButton"));
                xRadioButton.setState(true);
                xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("leftLayoutOptionButton"));
                xRadioButton.setState(false);
            }
            enableControl(xControlContainer.getControl("centerLayoutOptionButton"), bool);
            enableControl(xControlContainer.getControl("leftLayoutOptionButton"), bool);
        }
    }

    public boolean isModifyColorsPropsInDiagramPropsDialog(){
        if(m_xPropsDialog != null)
            if(m_xModifyColorsCheckBoxOfPD != null)
                if(m_xModifyColorsCheckBoxOfPD.getState() == 1)
                    return true;
        return false;
    }

    public boolean isBaseColorsPropsInDiagramPropsDialog(){
        if(m_xPropsDialog != null){
            if(m_xBaseColorRadioButtonOfPD != null){
                XRadioButton xRadioButton = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, m_xBaseColorRadioButtonOfPD);
                if(xRadioButton != null)
                    return xRadioButton.getState();
            }
        }
        return false;
    }
    //********************************************************************************************************


    //MessageBoxes *******************************************************************************************
    //MessageBox with OK button
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
                        enableControlDialogWindow(false);
                        xMessageBox.execute();
                        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                        if(xComponent != null)
                            xComponent.dispose();
                        enableControlDialogWindow(true);
                        setFocusControlDialog();
                    }
                }
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    //MessageBox with OK and Cancel button
    public void showMessageBox2(String sTitle, String sMessage){
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
                    if (xMessageBox != null){
                        enableControlDialogWindow(false);
                        m_iInfoBoxOk = xMessageBox.execute();
                        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                        if(xComponent != null)
                            xComponent.dispose();
                        enableControlDialogWindow(true);
                        setFocusControlDialog();
                    }
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
             System.err.println(ex.getLocalizedMessage());
        }
    }
    //********************************************************************************************************

    
    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            location =  xPIP.getPackageLocation("oxygenoffice.extensions.smart.SmART");
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

    public XGraphic setGraphic(XControl xImageControl, String sImageUrl){
        XGraphic xGraphic = null;
        try{
            Object oGraphicProvider = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.graphic.GraphicProvider", m_xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);
            PropertyValue[] aPropertyValues = new PropertyValue[1];
            PropertyValue aPropertyValue = new PropertyValue();
            aPropertyValue.Name = "URL";
            if(!sImageUrl.equals(""))
                aPropertyValue.Value = getPackageLocation() + sImageUrl;
            else
                aPropertyValue.Value = "";
            aPropertyValues[0] = aPropertyValue;
            xGraphic = xGraphicProvider.queryGraphic(aPropertyValues);
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xImageControl.getModel());
            xProps.setPropertyValue("Graphic", xGraphic);
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xGraphic;
    }

    public void setControlScaleImageProp(XControl xControl, boolean bool){
        try{
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            xProps.setPropertyValue("ScaleImage", new Boolean(bool));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
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

    public void visibleControl(XControl xControl, boolean bool){
        //XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, obj);
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("Visible", new Boolean(bool));
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

    public void setImageColorOfControl(XControl xControl, int color){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
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

    public int getImageColorOfControl(XControl xControl){
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

    public void setImageOfObject(Object object, String sImageUrl, short imageAlign){
        if(object != null){
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, object);
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

    public int getNum(String name){
        String s ="";
        char[] charName = name.toCharArray();
        int i = 5;
        while(i<name.length())
           s +=  charName[i++];
        return getController().parseInt(s);
    }

/*
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
 */
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
    
}