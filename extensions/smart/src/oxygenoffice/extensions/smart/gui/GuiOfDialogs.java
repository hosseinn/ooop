package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ImageAlign;
import com.sun.star.awt.ImageScaleMode;
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
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XInitialization;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.ui.dialogs.XExecutableDialog;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.ui.dialogs.XFilePickerControlAccess;
import com.sun.star.ui.dialogs.XFilterManager;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class GuiOfDialogs extends GuiBaseClass {


    private     boolean             isShownTips;
    private     ListenerOfDialogs   m_oListener = null;
    protected   XControl            m_xEventObjectControl       = null;

    
    //DiagramGalleryDialog *********************************************
    private     XDialog             m_xGalleryDialog            = null;
    public static int               NUM_OF_DIAGRAMS             = 12;
    //******************************************************************
    

    //ControlDialog ****************************************************
    private     XDialog             m_xControlDialog            = null;
    private     XWindow             m_xControlDialogWindow      = null;
    protected   XTopWindow          m_xControlDialogTopWindow   = null;
    private     boolean             isVisibleControlDialog      = false;
    private     String              m_sWorkUrl                  = "";
//    private     boolean             m_isShownTFOfControlDialog  = true;
    //******************************************************************

    
    //ColorTable dialog ************************************************
    private     XDialog             m_xColorTableDialog         = null;
    //protected   XTopWindow          m_xColorTableTopWindow      = null;
    private     XDialog             m_xColorTableDialog3        = null;
    //protected   XTopWindow          m_xColorTableTopWindow3     = null;
    //******************************************************************

    
    //GradientDialog ***************************************************
    private     XDialog             m_xGradientDialog           = null;
    private     XWindow             m_xGradientDialogWindow     = null;
    private     boolean             m_IsShownGradientDialog     = false;
    //******************************************************************


    //ConvertDialog ***************************************************
    private     XDialog             m_xConvertDialog            = null;
    //******************************************************************

    
    //Properties dialogs ***********************************************
    private     boolean             m_IsShownPropsDialog        = false;
    protected     XWindow           m_xPropsDialogWindow        = null;
    //******************************************************************
    
    
    //BaseColorsSettings dialogs ***************************************
    private XDialog                 m_xBCSettingsDialog         = null;
    private     boolean             m_IsShownBCSettingsDialog   = false;
    //******************************************************************
    
    
    //MessageBoxes *****************************************************
    private     int                 m_iInfoBoxOk                = -1;
    //******************************************************************
  

    public GuiOfDialogs(){ }

    public GuiOfDialogs(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
        m_oListener = new ListenerOfDialogs(this, m_Controller);
        isShownTips = true;
    }


    //DiagramGalleryDialog's methods ***********************************************************************    
    public short executeGalleryDialog(){
        if(isVisibleControlDialog)
            setVisibleControlDialog(false);
        getController().setGroupType(Controller.ORGANIGROUP);
        getController().setDiagramType(Controller.SIMPLEORGANIGRAM);
        getController().setNullDiagram();
        createDiagramGallery2();
        short exec = m_xGalleryDialog.execute();
        setShownTipsProp();
        if(exec != 1)
            getController().setLastDiagramName("");
        XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xGalleryDialog);
        xComp.dispose();
        return exec;
    }

    public boolean isShownTips(){
        return isShownTips;
    }

    public void setShownTipsProp(){
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
    
    public XGraphic getGraphic(String sImageUrl){
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
            return xGraphicProvider.queryGraphic(aPropertyValues);
        } catch (java.lang.Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }

    public void createDiagramGallery2(){
        try {
            String sPackageURL              = getPackageLocation();
            String sDialogURL               = sPackageURL + "/dialogs/DiagramGallery2.xdl";
            XDialogProvider2 xDialogProv    = getDialogProvider();
            m_xGalleryDialog = xDialogProv.createDialogWithHandler( sDialogURL, m_oListener );
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            XListBox diagramGroupList = (XListBox) UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("ListBox"));
            diagramGroupList.removeItems((short)3, (short)4);
            diagramGroupList.selectItemPos((short)0, true);
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
            oButton = xControlContainer.getControl("Item8");
            setImageOfObject(oButton, sPackageURL + "/images/continuousBlockProcess.png", (short)-1);
            oButton = xControlContainer.getControl("Item9");
            setImageOfObject(oButton, sPackageURL + "/images/staggeredProcess.png", (short)-1);
            oButton = xControlContainer.getControl("Item10");
            setImageOfObject(oButton, sPackageURL + "/images/bendingProcess.png", (short)-1);
            oButton = xControlContainer.getControl("Item11");
            setImageOfObject(oButton, sPackageURL + "/images/upwardArrowProcess.png", (short)-1);

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
        int numOfButtons = NUM_OF_DIAGRAMS;
        int n = getController().getGroupType() * 4;
        int m = n + 3;
        if(getController().getGroupType() == Controller.MATRIXGROUP)
            m -= 2;
        if(m_xGalleryDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            if(xControlContainer != null){
                boolean bool;
                for(int i = 0; i < numOfButtons; i++){
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
        if(m_xGalleryDialog != null){
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
                sType = "TargetDiagram";
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
            
            if( type == Controller.CONTINUOUSBLOCKPROCESS){
                sType = "ContinuousBlockProcess";
                sourceFileName += "2";
            }
            if( type == Controller.STAGGEREDPROCESS){
                sType = "StaggeredProcess";
                sourceFileName += "2";
            }
            if( type == Controller.BENDINGPROCESS){
                sType = "BendingProcess";
                sourceFileName += "2";
            }
            if( type == Controller.UPWARDARROWPROCESS){
                sType = "UpwardArrowProcess";
                sourceFileName += "2";
            }
            
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGalleryDialog);
            XFixedText xDiagramNameText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label1"));
            XFixedText xDiagramDescriptionText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("Label2"));
            if(xDiagramNameText != null && xDiagramDescriptionText != null){
                if(type == Controller.NOTDIAGRAM){
                    xDiagramNameText.setText("");
                    xDiagramDescriptionText.setText("");
                }else{
                    String diagramNameProperty = sourceFileName + "." + sType + ".Label";
                    String diagramDescProperty = sourceFileName + "." + sType + "Description.Label";
                    xDiagramNameText.setText(getDialogPropertyValue(sourceFileName, diagramNameProperty));
                    xDiagramDescriptionText.setText(getDialogPropertyValue(sourceFileName, diagramDescProperty));
                }
            }
        }
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
        if( m_xControlDialogWindow == null)
            createControlDialog();
        if(m_xControlDialogWindow != null){
            m_xControlDialogWindow.setVisible(bool);
            if(bool){
                isVisibleControlDialog = true;
                setColorModeOfImageOfControlDialog();
                if(getController().getDiagram().isBaseColorsMode()){
                    if(getController().getGroupType() == Controller.RELATIONGROUP)
                        setImageColorOfControlDialog(((RelationDiagram)getController().getDiagram()).getNextColor());
                    if(getController().getGroupType() == Controller.PROCESSGROUP)
                        setImageColorOfControlDialog(((ProcessDiagram)getController().getDiagram()).getNextColor());
                }
                m_xControlDialogWindow.setFocus();
            }else{
                isVisibleControlDialog = false;
            }
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
                setControlScaleModeProp(xControlContainer.getControl("ImageControl"), ImageScaleMode.None);
                //setControlScaleImageProp(xControlContainer.getControl("ImageControl"), false);
              
                if(getController().getGroupType() == Controller.ORGANIGROUP && getController().getDiagram() != null)
                    ((OrganizationChart)getController().getDiagram()).setNewItemHType(OrgChart.UNDERLING);
                setImageOfObject(xControlContainer.getControl("convertButton"), sPackageURL + "/images/convert.png", ImageAlign.LEFT);
                setImageOfObject(xControlContainer.getControl("exportButton"), sPackageURL + "/images/export.png", ImageAlign.LEFT);
//                m_isShownTFOfControlDialog = true;
//                moveTextFieldOfControlDialog();
                setTextFieldOfControlDialog();
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

    public void enableAndSetFocusControlDialog(){
        enableControlDialogWindow(true);
        setFocusControlDialog();
    }

    public XWindow getControlDialogWindow(){
        return m_xControlDialogWindow;
    }

    public void setColorThemeModeOnImageOfControlDialog(){
        short colorMode = getController().getDiagram().getColorModeProp();
        if(colorMode == Diagram.GREEN_DARK_COLOR_MODE || colorMode == Diagram.GREEN_DARK_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/greenDarkTheme_mode.png");
        if(colorMode == Diagram.GREEN_BRIGHT_COLOR_MODE || colorMode == Diagram.GREEN_BRIGHT_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/greenBrightTheme_mode.png");
        if(colorMode == Diagram.BLUE_DARK_COLOR_MODE || colorMode == Diagram.BLUE_DARK_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/blueDarkTheme_mode.png");
        if(colorMode == Diagram.BLUE_BRIGHT_COLOR_MODE || colorMode == Diagram.BLUE_BRIGHT_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/blueBrightTheme_mode.png");
        if(colorMode == Diagram.PURPLE_DARK_COLOR_MODE || colorMode == Diagram.PURPLE_DARK_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/purpleDarkTheme_mode.png");
        if(colorMode == Diagram.PURPLE_BRIGHT_COLOR_MODE || colorMode == Diagram.PURPLE_BRIGHT_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/purpleBrightTheme_mode.png");
        if(colorMode == Diagram.ORANGE_DARK_COLOR_MODE || colorMode == Diagram.ORANGE_DARK_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/orangeDarkTheme_mode.png");
        if(colorMode == Diagram.ORANGE_BRIGHT_COLOR_MODE || colorMode == Diagram.ORANGE_BRIGHT_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/orangeBrightTheme_mode.png");
        if(colorMode == Diagram.YELLOW_DARK_COLOR_MODE || colorMode == Diagram.YELLOW_DARK_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/yellowDarkTheme_mode.png");
        if(colorMode == Diagram.YELLOW_BRIGHT_COLOR_MODE || colorMode == Diagram.YELLOW_BRIGHT_GRADIENT_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/yellowBrightTheme_mode.png");
    }
    
    public void setSchemeModeOnImageOfControlDialog(){
        short colorMode = getController().getDiagram().getColorModeProp();
        if(colorMode == Diagram.BLUE_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/blueDesign_mode.png");
        if(colorMode == Diagram.AQUA_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/aquaDesign_mode.png");
        if(colorMode == Diagram.RED_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/redDesign_mode.png");
        if(colorMode == Diagram.FIRE_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/fireDesign_mode.png");
        if(colorMode == Diagram.SUN_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/sunDesign_mode.png");
        if(colorMode == Diagram.GREEN_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/greenDesign_mode.png");
        if(colorMode == Diagram.OLIVE_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/oliveDesign_mode.png");
        if(colorMode == Diagram.PURPLE_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/purpleDesign_mode.png");
        if(colorMode == Diagram.PINK_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/pinkDesign_mode.png");
        if(colorMode == Diagram.INDIAN_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/indianDesign_mode.png");
        if(colorMode == Diagram.MAROON_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/maroonDesign_mode.png");
        if(colorMode == Diagram.BROWN_SCHEME_COLOR_MODE)
            setGraphic(getImageControlOfControlDialog(), "/images/brownDesign_mode.png");
        
    }

    public void setColorModeOfImageOfControlDialog(){
        if(isVisibleControlDialog && getImageControlOfControlDialog() != null){
            if(getController().getDiagram().isSimpleColorMode()){
                setGraphic(getImageControlOfControlDialog(), "");
                setImageColorOfControlDialog(getController().getDiagram().getColorProp());
            }
            if(getController().getDiagram().isBaseColorsMode()) {
                setGraphic(getImageControlOfControlDialog(), "/images/basecolors_mode.png");
            }
            if(getController().getDiagram().isBaseColorsWithGradientMode()){
                setGraphic(getImageControlOfControlDialog(), "/images/basecolorsWithGradients_mode.png");
            }
            if (getController().getDiagram().isGradientColorMode()) {
                setGraphic(getImageControlOfControlDialog(), "/images/gradient_mode.png");
            }
            if (getController().getDiagram().isColorThemeGradientMode()) {
                setColorThemeModeOnImageOfControlDialog();
            }
            if (getController().getDiagram().isColorSchemeMode()) {
                setSchemeModeOnImageOfControlDialog();
            }
            if (getController().getDiagram().isColorThemeMode()) {
                setColorThemeModeOnImageOfControlDialog();
            }
        }
    }

    public XControl getImageControlOfControlDialog(){
        XControl xControl = null;
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            xControl = xControlContainer.getControl("ImageControl");
        }
        return xControl;
    }
    
    public void setImageColorOfControlDialog(int color){
        setImageColorOfControl(getImageControlOfControlDialog(), color);
    }

    public int getImageColorOfControlDialog(){
       return getImageColorOfControl(getImageControlOfControlDialog());
    }
/* 
    public void enableVisibleTextFieldOfControlDialog(boolean bool){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            enableVisibleControl((XControl) UnoRuntime.queryInterface(XControl.class, xTFOfControlDialog), bool);
         }
    }
*/
    public void enableTextFieldOfControlDialog(boolean bool){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));  
            enableControl((XControl) UnoRuntime.queryInterface(XControl.class, xTFOfControlDialog), bool);
        }
    }
/*
    public boolean isShownTextFieldOfControlDialog(){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            if(xTFOfControlDialog != null)
                return m_isShownTFOfControlDialog;
        }
        return false;
    }

    public void moveTextFieldOfControlDialog(){
        try {
            if(m_xControlDialog != null){
                XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, m_xControlDialog);
                XPropertySet xPS = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                String sPackageURL = getPackageLocation();
                m_isShownTFOfControlDialog = !m_isShownTFOfControlDialog;
                
                int plus = 0;
                double version = 0.0;
                try{
                    version = Double.parseDouble(getController().getSmartPH().getOOSetupVersion());
                }catch(NumberFormatException ex){
                    System.err.println(ex.getLocalizedMessage());
                }
                if(version < 3.6)
                    plus = 1;
                int dialogWidth = 259 + plus;
                if(getController().getGroupType() != Controller.ORGANIGROUP)
                    dialogWidth = 224 + plus;
                int posX = AnyConverter.toInt(xPS.getPropertyValue("PositionX"));
                int posY = AnyConverter.toInt(xPS.getPropertyValue("PositionY"));
                
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
                Object oUpAndDownBOfControlDialog = xControlContainer.getControl("downUpButton");

                if(m_isShownTFOfControlDialog){
                    enableVisibleTextFieldOfControlDialog(true);
                    if(getController().isOnlySimpleItemIsSelected()){
                        setTextFieldOfControlDialog();
                        setFocusTextFieldOfControlDialog();
                    }else{
                        enableTextFieldOfControlDialog(false);
                    }
                    xPS.setPropertyValue("Width", new Integer(dialogWidth));
                    xPS.setPropertyValue("Height", new Integer(39));
                    if(oUpAndDownBOfControlDialog != null)
                        setImageOfObject(oUpAndDownBOfControlDialog, sPackageURL + "/images/up.png", (short)-1);
                }else{
                    enableVisibleTextFieldOfControlDialog(false);
                    xPS.setPropertyValue("Width", new Integer(dialogWidth));
                    xPS.setPropertyValue("Height", new Integer(16));
                    if(oUpAndDownBOfControlDialog != null)
                        setImageOfObject(oUpAndDownBOfControlDialog, sPackageURL + "/images/down.png", (short)-1);
                }
                //set old position (LO 3.6 move dialog after resize)
                xPS.setPropertyValue("PositionX", new Integer(posX));
                xPS.setPropertyValue("PositionY", new Integer(posY));
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
*/
    public void modifiedTextFieldOfControlDialog(){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            
            if(xTFOfControlDialog != null)
                getController().setTextOfSelectedShape(xTFOfControlDialog.getText());
        }
    }

    public String getTextFieldValueControlDialog(){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            if(xTFOfControlDialog != null)
                return xTFOfControlDialog.getText();
        }
        return "";
    }

    public void setFocusTextFieldOfControlDialog(){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            if(xTFOfControlDialog != null){
                XWindow xWindow = (XWindow)UnoRuntime.queryInterface(XWindow.class, xTFOfControlDialog);
                if(xWindow != null)
                    xWindow.setFocus();
            }
        }
    }

    public void setTextFieldValueOfControlDialog(String str){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));
            if(xTFOfControlDialog != null)
                xTFOfControlDialog.setText(str);
        }
    }

    public void setTextFieldOfControlDialog(){
        if(m_xControlDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xControlDialog);
            XTextComponent xTFOfControlDialog = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xControlContainer.getControl("textField"));

            if(xTFOfControlDialog != null){
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
                    }
                    if(getController().getGroupType() == Controller.PROCESSGROUP){
                        ProcessDiagramItem item = ((ProcessDiagram)getController().getDiagram()).getItem(getController().getSelectedShape());
                        if(item != null){
                            XShape xTextShape = item.getTextShape();
                            if(xTextShape != null && getController().getDiagram().isInGruopShapes(xTextShape)){
                                xText = (XText)UnoRuntime.queryInterface(XText.class, xTextShape);
                                setTextFieldValueOfControlDialog(xText.getString());
                            }else
                                setTextFieldValueOfControlDialog("");
                        }
                        
                    }
                    if(getController().getGroupType() == Controller.ORGANIGROUP){
                        xText = (XText)UnoRuntime.queryInterface(XText.class, getController().getSelectedShape());
                        setTextFieldValueOfControlDialog(xText.getString());
                    }
                    enableTextFieldOfControlDialog(true);
                }else{
//                    String message = getDialogPropertyValue("Strings2", "Strings2.InactiveFunction.Label");
//                    setTextFieldValueOfControlDialog(message);
                    setTextFieldValueOfControlDialog("");
                    enableTextFieldOfControlDialog(false);
                }
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
            m_xColorTableDialog.execute();
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xColorTableDialog);
            if(xComp != null)
                xComp.dispose();
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
            m_xColorTableDialog.execute();
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xColorTableDialog);
            if(xComp != null)
                xComp.dispose();
        } catch (IllegalArgumentException e) {
            System.err.println(e.getLocalizedMessage());
        }
    }

    public void executeColorTable3() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/ColorTable3.xdl";
            m_xColorTableDialog3 = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if(m_xColorTableDialog3 != null){
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog3);
                enableSetBaseColorsSettingsCommandButtonOnColorTable3(false);
                if(getController().getDiagram().isSimpleColorMode()){
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("simpleColorModeOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorThemeModeOptionButton"), false);
                    enableControl(xControlContainer.getControl("colorSchemeModeOptionButton"), false);
                }
                if(getController().getDiagram().isBaseColorsMode() || getController().getDiagram().isBaseColorsWithGradientMode()){
                    if(getController().getDiagram().isBaseColorsMode())
                        ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsModeOptionButton"))).setState(true);
                    if(getController().getDiagram().isBaseColorsWithGradientMode()) 
                        ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsGradientsModeOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorThemeModeOptionButton"), false);
                    enableControl(xControlContainer.getControl("colorSchemeModeOptionButton"), false);
                    enableSetBaseColorsSettingsCommandButtonOnColorTable3(true);
                }
                if(getController().getDiagram().isColorThemeMode() || getController().getDiagram().isColorThemeGradientMode()) {
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorThemeModeOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorSchemeModeOptionButton"), false);
                    enableImageControlsOfColorTable(false);
                }
                if(getController().getDiagram().isColorSchemeMode()) {
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorSchemeModeOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorThemeModeOptionButton"), false);
                    enableImageControlsOfColorTable(false);
                }
                if(getController().getDiagramType() == Controller.VENNDIAGRAM) {
                    enableControl(xControlContainer.getControl("colorSchemeModeOptionButton"), false);
                }
                if(getController().getDiagramType() != Controller.PYRAMIDDIAGRAM && getController().getDiagramType() != Controller.UPWARDARROWPROCESS) {
                    enableControl(xControlContainer.getControl("baseColorsGradientsModeOptionButton"), false);
                }
                m_xColorTableDialog3.execute();
                XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xColorTableDialog3);
                if(xComp != null)
                    xComp.dispose();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void showColorTableHelpDialog(){
        String title = getDialogPropertyValue("Strings2", "Strings2.ColorTableHelp.Title.Label");
        String message = getDialogPropertyValue("Strings2", "Strings2.ColorTableHelp.Message.Label");
        showMessageBox(title, message);
        if(m_xColorTableDialog3 != null)
            ((XWindow) UnoRuntime.queryInterface(XWindow.class, m_xColorTableDialog3)).setFocus();         
    }
    
    public void enableSetBaseColorsSettingsCommandButtonOnColorTable3(boolean bool){
        if(m_xColorTableDialog3 != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog3);
            enableControl(xControlContainer.getControl("baseColorsSettingsCommanButton"), bool);
        }
    }

    public void setColorModeBasedColorTable3(){
        if(m_xColorTableDialog3 != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog3);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("simpleColorModeOptionButton"))).getState())    
                getController().getDiagram().setColorModeProp(Diagram.SIMPLE_COLOR_MODE);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsModeOptionButton"))).getState())    
                getController().getDiagram().setColorModeProp(Diagram.BASE_COLORS_MODE);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsGradientsModeOptionButton"))).getState())    
                getController().getDiagram().setColorModeProp(Diagram.BASE_COLORS_WITH_GRADIENT_MODE);
        }
    }
    
    public void showBaseColorsSettingsDialog(){
        setShownBaseColorsSettingsDialog(true);
        short exec = executeBCSettingDialog();
        setShownBaseColorsSettingsDialog(false);
        if(exec == 1)
            setNewBaseColors();
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
   
    public void setResultColorTable(String methodName){
        int color = -1;
        if(!methodName.equals(""))
            color = getCurrColorOfColorTable(getNum(methodName));
        if(color != -1){
//            System.out.println(Integer.toHexString(color));
            if(isShownBaseColorsSettingsDialog() && m_xBCSettingsDialog != null){
                if(m_xEventObjectControl != null)
                    setImageColorOfControl(m_xEventObjectControl, color);
            }else if(isShownGradientDialog() || (m_xPropsDialogWindow != null && isShownPropsDialog())){
                if(m_xEventObjectControl != null)
                    setImageColorOfControl(m_xEventObjectControl, color);
            }else{
                setImageColorOfControlDialog(color);
                getController().getDiagram().setColorProp(color);
//                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
//                    setBaseOrSimpleColorModeProp();
            }
        }
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
        if(m_xColorTableDialog != null){
            m_xColorTableDialog.endExecute();
            m_xColorTableDialog = null;
            setFocusParentWindow();
        }
    }
    
    public void endExecuteColorTable3(){
        if(m_xColorTableDialog3 != null){
            m_xColorTableDialog3.endExecute();
            m_xColorTableDialog3 = null;
            setFocusParentWindow();
        }
    }

    public void setFocusParentWindow(){

        if(isShownBaseColorsSettingsDialog()) {
            setFocusBaseColorsSettingsDialog();
        } else if(isShownGradientDialog()) {
            setFocusGradientDialog();
        } else if (isShownPropsDialog()) {
            setFocusPropsDialog();
        } else if (isVisibleControlDialog()) {
            enableControlDialogWindow(true);
            setFocusControlDialog();
        }
    }

    public int getCurrColorOfColorTable(int i){
        int color = -1;
        try {
            XControlContainer xControlContainer = null;
            if(isShownBaseColorsSettingsDialog()){
                xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
            } else if(m_xColorTableDialog3 != null){
                xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog3);
            } else {
                xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xColorTableDialog);
            }
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
    //********************************************************************************************************


    //GradientDialog's methods *******************************************************************************
    public boolean isShownGradientDialog(){
        return m_IsShownGradientDialog;
    }

    public void setFocusGradientDialog(){
        if(m_xGradientDialogWindow != null)
            m_xGradientDialogWindow.setFocus();
    }

    public void executeGradientDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/GradientDialog.xdl";
            m_xGradientDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xGradientDialog != null) {
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
                setStartColorOfGradientDialog(getController().getDiagram().getStartColorProp());
                setEndColorOfGradientDialog(getController().getDiagram().getEndColorProp());
                XListBox xListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("gradientModeListBox"));
                xListBox.selectItemPos(getController().getDiagram().getGradientDirectionProp(), true);

                m_xGradientDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xGradientDialog);
                m_IsShownGradientDialog = true;
                short exec = m_xGradientDialog.execute();
                m_IsShownGradientDialog = false;
                if(exec == 1){
                    XRadioButton xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("setGradientOptionButton"));
                    if(xRadioButton.getState()){
                        getController().getDiagram().setStartColorProp(getImageColorOfControl(xControlContainer.getControl("startColor")));
                        getController().getDiagram().setEndColorProp(getImageColorOfControl(xControlContainer.getControl("endColor")));
                        getController().getDiagram().setGradientDirectionProp(xListBox.getSelectedItemPos());
                        getController().getDiagram().setColorModeProp(Diagram.GRADIENT_COLOR_MODE);
                    }else{
                        enableControlDialogWindow(false);
                        executeColorTable2();
                        enableControlDialogWindow(true);
                        setFocusControlDialog();
                    }
                }
                XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xGradientDialog);
                if(xComp != null)
                    xComp.dispose();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setStartColorOfGradientDialog(int color){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
        if(xControlContainer != null)
            setImageColorOfControl(xControlContainer.getControl("startColor"), color);
    }

    public void setEndColorOfGradientDialog(int color){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
        if(xControlContainer != null)
            setImageColorOfControl(xControlContainer.getControl("endColor"), color);
    }

    public void enableGradientDialogControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xGradientDialog);
        if(xControlContainer != null){
            enableControl(xControlContainer.getControl("label1"), bool);
            enableControl(xControlContainer.getControl("startColor"), bool);
            enableControl(xControlContainer.getControl("label2"), bool);
            enableControl(xControlContainer.getControl("endColor"), bool);
            enableControl(xControlContainer.getControl("gradientModeListBox"), bool);
        }
    }
    //********************************************************************************************************

    
    //ConvertDialog's methods ********************************************************************************
    public short executeConvertDialog() {
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL;
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                sDialogURL += "/dialogs/OrganigroupConvertDialog.xdl";
            if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                sDialogURL += "/dialogs/RelationAndProcessgroupConvertDialog.xdl";
            m_xConvertDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xConvertDialog != null){
                setConvertDialogControls();
                enableControlDialogWindow(false);
                short exec = m_xConvertDialog.execute();
                enableAndSetFocusControlDialog();
                return exec;
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return -1;
    }
    
    public void disposeConvertDialog(){
        if(m_xConvertDialog != null){
            XComponent xComponent = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xConvertDialog);
            if(xComponent != null)
                xComponent.dispose();
            enableAndSetFocusControlDialog();
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
                        XListBox xListBoxOfConvertDialog = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lastHorLevelListBox"));
                        xListBoxOfConvertDialog.selectItemPos(OrgChartTree.LASTHORLEVEL, true);
                    }
                }

                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP){
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
                    if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS) {
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton6"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 5;
                    }
                    if(getController().getDiagramType() == Controller.STAGGEREDPROCESS) {
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton5"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 6;
                    }
                    if(getController().getDiagramType() == Controller.BENDINGPROCESS) {
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton5"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 7;
                    }
                    if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS) {
                        xRadioButton2 = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton5"));
                        xRadioButton2.setState(true);
                        disabledOptionButton += 8;
                    }
                    xControl = (XControl) UnoRuntime.queryInterface(XControl.class, xControlContainer.getControl(disabledOptionButton));
                    enableControl(xControl, false);
                }
            }
        }
    }

    public short getSelectedItemPosOfConvertDialogListBox(){
        if(m_xConvertDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            XListBox xListBoxOfConvertDialog = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lastHorLevelListBox"));
            return xListBoxOfConvertDialog.getSelectedItemPos();
        }
        return -1;
    }
    
    public short getConversationType(){
        if(m_xConvertDialog != null){
            XRadioButton xRadioButton = null;
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xConvertDialog);
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton1"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.SIMPLEORGANIGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.VENNDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton2"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.HORIZONTALORGANIGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.CYCLEDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton3"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.TABLEHIERARCHYDIAGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.PYRAMIDDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton4"));
            if(xRadioButton.getState() == true){
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    return Controller.ORGANIGRAM;
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.TARGETDIAGRAM;
            }
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton5"));
            if(xRadioButton.getState() == true)
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.CONTINUOUSBLOCKPROCESS;
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton6"));
            if(xRadioButton.getState() == true)
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.STAGGEREDPROCESS;
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton7"));
            if(xRadioButton.getState() == true)
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.BENDINGPROCESS;
            xRadioButton = (XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("convertOptionButton8"));
            if(xRadioButton.getState() == true)
                if(getController().getGroupType() == Controller.RELATIONGROUP || getController().getGroupType() == Controller.PROCESSGROUP)
                    return Controller.UPWARDARROWPROCESS;
            
            
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
    //********************************************************************************************************


    //Properties dialogs' methods ****************************************************************************

    public boolean isShownPropsDialog(){
        return m_IsShownPropsDialog;
    }

    public void setShownPropsDialog(boolean bool){
        m_IsShownPropsDialog = bool;
    }
    
    public void setFocusPropsDialog(){
        if(m_xPropsDialogWindow != null)
            m_xPropsDialogWindow.setFocus();
    }
    //********************************************************************************************************

    
    //BaseColorsSettingsDiaglog's methods *******************************************************************
    public boolean isShownBaseColorsSettingsDialog(){
        return m_IsShownBCSettingsDialog;
    }
    
    public void setShownBaseColorsSettingsDialog(boolean bool){
        m_IsShownBCSettingsDialog = bool;
    }
    
    public void setFocusBaseColorsSettingsDialog(){
        if(m_xBCSettingsDialog != null)
            ((XWindow)UnoRuntime.queryInterface(XWindow.class, m_xBCSettingsDialog)).setFocus();
    }
    
    public short executeBCSettingDialog(){
        try {
            XDialogProvider2 xDialogProv = getDialogProvider();
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/BaseColorsSettings.xdl";
            m_xBCSettingsDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListener);
            if (m_xBCSettingsDialog != null){
                setBCSettingDialog();
                return m_xBCSettingsDialog.execute();
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return -1;
    }
    
    public void setBCSettingDialog(){
        if(m_xBCSettingsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xBCSettingsDialog);
                if(xControlContainer != null)
                    for(int i = 1; i<= 8; i++)
                        setImageColorOfControl(xControlContainer.getControl("ImageControl" + i), Diagram._aBaseColors[i-1]);
        }
    }
    
    public void setDefaultBaseColors(){
        if(m_xBCSettingsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xBCSettingsDialog);
            if(xControlContainer != null)
                for(int i = 1; i<= 8; i++){
                    if(getController().getGroupType() == Controller.RELATIONGROUP)
                        setImageColorOfControl(xControlContainer.getControl("ImageControl" + i), Diagram.aBASECOLORS[i-1]);
                    if(getController().getGroupType() == Controller.PROCESSGROUP)
                        setImageColorOfControl(xControlContainer.getControl("ImageControl" + i), Diagram.aRAINBOWCOLORS[i-1]);
                }
        }
    }
    
    public void setNewBaseColors(){
        if(m_xBCSettingsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xBCSettingsDialog);
            if(xControlContainer != null)
                for(int i = 0; i < 8; i++)
                    Diagram._aBaseColors[i] = getImageColorOfControl(xControlContainer.getControl("ImageControl" + (i + 1)));
        }
    }
    //*******************************************************************************************************

    
    //TipsMessageBoxes *******************************************************************************************
    //MessageBox with OK button
    public void showTipsMessageBox(){
        String title = getDialogPropertyValue("Strings2", "Strings2.Tips.Name.Label");
        String message = "";
        if(getController().getGroupType() == Controller.ORGANIGROUP)
            message = getDialogPropertyValue("Strings2", "Strings2.Organizationcharts.Tips.Label");
        if(getController().getGroupType() == Controller.RELATIONGROUP)
            message = getDialogPropertyValue("Strings2", "Strings2.Relationdiagrams.Tips.Label");
        if(getController().getGroupType() == Controller.PROCESSGROUP)
            message = getDialogPropertyValue("Strings2", "Strings2.Processes.Tips.Label");
        message += "\n" + getDialogPropertyValue("Strings2", "Strings2.AllDiagram.Tips2.Label");
        message += getDialogPropertyValue("Strings2", "Strings2.Help.Message.Label");
        showMessageBox(title, message);
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
                //aDescriptor.Bounds            = new Rectangle(0,0,60000,1000);
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
    //*******************************************************************************
    
    public String raiseSaveAsDialog() {
        String sStorePath = "";
        XComponent xComponent = null;
        try {
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object oFilePicker = xMCF.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", m_xContext);
            XFilePicker xFilePicker = (XFilePicker)UnoRuntime.queryInterface(XFilePicker.class, oFilePicker);
            if(m_sWorkUrl.equals("")){
                Object oPathSettings = xMCF.createInstanceWithContext("com.sun.star.util.PathSettings", m_xContext);
                XPropertySet xPropertySet = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, oPathSettings);
                m_sWorkUrl = (String) xPropertySet.getPropertyValue("Work");
            }
            xFilePicker.setDisplayDirectory(m_sWorkUrl);
            
            // choose the template that defines the capabilities of the filepicker dialog
            XInitialization xInitialize = (XInitialization)UnoRuntime.queryInterface(XInitialization.class, xFilePicker);
            Short[] listAny = new Short[] { new Short(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_AUTOEXTENSION)};
            xInitialize.initialize(listAny);
            
            String fileName = getController().getDiagram().getDiagramTypeName();
            xFilePicker.setDefaultName(fileName);
            
            // add a control to the dialog to add the extension automatically to the filename...
            XFilePickerControlAccess xFilePickerControlAccess =(XFilePickerControlAccess)UnoRuntime.queryInterface(XFilePickerControlAccess.class, xFilePicker);
            xFilePickerControlAccess.setValue(com.sun.star.ui.dialogs.ExtendedFilePickerElementIds.CHECKBOX_AUTOEXTENSION, (short) 0, new Boolean(true));
            
            XFilterManager xFilterManager = (XFilterManager)UnoRuntime.queryInterface(XFilterManager.class, xFilePicker);
            xFilterManager.appendFilter("PNG - Portable Network Graphic", ".png");
            xFilterManager.appendFilter("JPEG - Joint Photographic Experts Group", ".jpg");
            xFilterManager.appendFilter("GIF - Graphics Interchange Format", ".gif");
            
            xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xFilePicker);
            XExecutableDialog xExecutable = (XExecutableDialog)UnoRuntime.queryInterface(XExecutableDialog.class, xFilePicker);
            short nResult = xExecutable.execute();
            if (nResult == com.sun.star.ui.dialogs.ExecutableDialogResults.OK){
                m_sWorkUrl = xFilePicker.getDisplayDirectory();
                String[] sPathList = xFilePicker.getFiles();
                if (sPathList.length > 0){
                    sStorePath = sPathList[0];
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        } finally {
            if (xComponent != null){
                xComponent.dispose();
            }
        }
        return sStorePath;
    }
/*    
    public String raiseFolderPicker() {
        String sReturnFolder = "";
        XComponent xComponent = null;
        try {
            Object oFolderPicker = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", m_xContext);
            XFolderPicker xFolderPicker = (XFolderPicker)UnoRuntime.queryInterface(XFolderPicker.class, oFolderPicker);
            XExecutableDialog xExecutable = (XExecutableDialog)UnoRuntime.queryInterface(XExecutableDialog.class, oFolderPicker);
            xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, oFolderPicker);
            if(m_sWorkUrl.equals("")){
                Object oPathSettings = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.util.PathSettings", m_xContext);
                XPropertySet xPropertySet = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, oPathSettings);
                m_sWorkUrl = (String) xPropertySet.getPropertyValue("Work");
            }
            xFolderPicker.setDisplayDirectory(m_sWorkUrl);
            xFolderPicker.setTitle("hello");
            short nResult = xExecutable.execute();
            if (nResult == com.sun.star.ui.dialogs.ExecutableDialogResults.OK){
                m_sWorkUrl = xFolderPicker.getDisplayDirectory();
                sReturnFolder = xFolderPicker.getDirectory();
            }
        }catch( Exception ex ) {
            System.err.println(ex.getLocalizedMessage());
        }finally{
            if (xComponent != null){
                xComponent.dispose();
            }
        }
        System.out.println(sReturnFolder);
        return sReturnFolder;
    }
*/

}