package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XRadioButton;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.FontSize;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;


public class GuiOfOrganigroupPropsDialogs extends GuiOfDialogs {
    
    public      ListenerOfPropsDialogs  m_oListenerOfPropsDialogs   = null;
    public      XDialog                 m_xPropsDialog              = null;
    
    public GuiOfOrganigroupPropsDialogs(){ }
    
    public GuiOfOrganigroupPropsDialogs(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
    }
    
    //Properties dialogs' methods ****************************************************************************
    
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setOrganigramPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != OrganizationChart.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");

            enableVisibleOrganigramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfOrganigroupPD(true);
            setOrganigramPropsControls(false);
            
            if(sListPos == OrganizationChart.DEFAULT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_default.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_default.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_default.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_default.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localMedium, localYes, localNo);
                enableColorFunctionFieldOfOrganigroupPD(false);
                setConnImageControlColor();
            }
            if(sListPos == OrganizationChart.WITHOUT_OUTLINE){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_withoutOutline.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_withoutOutline.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_withoutOutline.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_withoutOutline.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localMedium, localNo, localNo);
                enableColorFunctionFieldOfOrganigroupPD(false);
                setConnImageControlColor();
            }
            if(sListPos == OrganizationChart.NOT_ROUNDED){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_notRounded.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_notRounded.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_notRounded.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_notRounded.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localNo, localYes, localNo);
                enableColorFunctionFieldOfOrganigroupPD(false);
                setConnImageControlColor();
            }
            if(sListPos == OrganizationChart.WITH_SHADOW){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_withShadow.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_withShadow.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_withShadow.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_withShadow.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localNo, "", localMedium, localNo, localYes);
                enableColorFunctionFieldOfOrganigroupPD(false);
                setConnImageControlColor();
            }

            if(sListPos == OrganizationChart.GREEN_DARK){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_greenDark.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_greenDark.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_greenDark.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_greenDark.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOGREENS[1]);
            }
            if(sListPos == OrganizationChart.GREEN_BRIGHT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_greenBright.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_greenBright.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_greenBright.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_greenBright.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOGREENS[3]);
            }
            if(sListPos == OrganizationChart.BLUE_DARK){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_blueDark.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_blueDark.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_blueDark.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_blueDark.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOBLUES[1]);
            }
            if(sListPos == OrganizationChart.BLUE_BRIGHT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_blueBright.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_blueBright.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_blueBright.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_blueBright.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOBLUES[3]);
            }
            if(sListPos == OrganizationChart.PURPLE_DARK){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_purpleDark.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_purpleDark.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_purpleDark.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_purpleDark.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOPURPLES[1]);
            }
            if(sListPos == OrganizationChart.PURPLE_BRIGHT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_purpleBright.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_purpleBright.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_purpleBright.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_purpleBright.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOPURPLES[3]);
            }
            if(sListPos == OrganizationChart.ORANGE_DARK){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_orangeDark.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_orangeDark.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_orangeDark.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_orangeDark.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOORANGES[1]);
            }
            if(sListPos == OrganizationChart.ORANGE_BRIGHT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_orangeBright.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_orangeBright.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_orangeBright.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_orangeBright.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOORANGES[3]);
            }
            if(sListPos == OrganizationChart.YELLOW_DARK){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_yellowDark.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_yellowDark.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_yellowDark.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_yellowDark.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOYELLOWS[1]);
            }
            if(sListPos == OrganizationChart.YELLOW_BRIGHT){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_yellowBright.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_yellowBright.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_yellowBright.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_yellowBright.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorTheme, localNo, localYes, localNo);
                setConnImageControlColor(Diagram.aLOYELLOWS[3]);
            }

            if(sListPos == OrganizationChart.BLUE_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_blueGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_blueGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_blueGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_blueGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.AQUA_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_aquaGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_aquaGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_aquaGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_aquaGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.RED_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_redGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_redGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_redGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_redGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.FIRE_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_fireGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_fireGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_fireGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_fireGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.SUN_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_sunGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_sunGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_sunGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_sunGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.GREEN_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_greenGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_greenGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_greenGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_greenGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.OLIVE_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_oliveGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_oliveGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_oliveGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_oliveGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.PURPLE_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_purpleGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_purpleGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_purpleGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_purpleGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.PINK_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_pinkGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_pinkGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_pinkGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_pinkGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.INDIAN_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_indianGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_indianGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_indianGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_indianGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.MAROON_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_maroonGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_maroonGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_maroonGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_maroonGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
            if(sListPos == OrganizationChart.BROWN_SCHEME){
                if(getController().getDiagramType() == Controller.SIMPLEORGANIGRAM)
                    setGraphic(xImageControl, "/images/simpleOrgchart_brownGradients.png");
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGraphic(xImageControl, "/images/hororgchart_brownGradients.png");
                if(getController().getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    setGraphic(xImageControl, "/images/tableHierarchy_brownGradients.png");
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    setGraphic(xImageControl, "/images/orgchart_brownGradients.png");
                setDescriptionLabelOfOrganigroupPD(localAllShape, localYes, localColorScheme, localMedium, localYes, localNo);
                setConnImageControlColor(Diagram.DEFAULT_CONN_COLOR);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleOrganigramPropsControls(true);
            enablePropertiesFieldOfOrganigroupPD(false);
            setOrganigramPropsControls(true);
            setConnImageControlColor();
        }
    }
    
    public void enableVisibleOrganigramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        //color controls
        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("siblingsOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("branchOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);
        enableVisibleControl(xControlContainer.getControl("gradientsOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("startColorLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("startColorImageControl"), bool);
        enableVisibleControl(xControlContainer.getControl("endColorLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("endColorImageControl"), bool);
        enableVisibleControl(xControlContainer.getControl("gradientDirectionLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("gradientDirectionListBox"), bool);
        
        //rounded controls
        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("noRoundedOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("mediumRoundedOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("extraRoundedOptionButton"), bool);

        //shadow controls
        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

        //outline controls
        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
    }
    
    public void setOrganigramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        
        if(bool){
            //color controls
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
            ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
            if(getController().getDiagram().isSimpleColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
            if(getController().getDiagram().isColorThemeGradientMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
            if(getController().getDiagram().isColorSchemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).setState(true);
            if(getController().getDiagram().isGradientColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("gradientsOptionButton"))).setState(true);
            setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
               
            short pos = 0;
            if(getController().getDiagram().isColorThemeGradientMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
            pos = 0;
            if(getController().getDiagram().isColorSchemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).selectItemPos(pos, true);
            setStartAndEndColorControlsInPropsDialog();
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("gradientDirectionListBox"))).selectItemPos(getController().getDiagram().getGradientDirectionProp(), true);
            enableOrganigramColorControls(false);
      
            //rounded controls
            short rounded = getController().getDiagram().getRoundedProp();
            if(rounded == Diagram.NULL_ROUNDED)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noRoundedOptionButton"))).setState(true);
            if(rounded == Diagram.MEDIUM_ROUNDED)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumRoundedOptionButton"))).setState(true);
            if(rounded == Diagram.EXTRA_ROUNDED)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraRoundedOptionButton"))).setState(true);
        
            //shadow controls
            if(getController().getDiagram().isShadowProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noShadowOptionButton"))).setState(true);
      
            //outline controls
            pos = 0;
            int lineWidth = getController().getDiagram().getShapesLineWidhtProp();
            if(lineWidth != 0)
                pos = (short)(lineWidth / 100);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).selectItemPos(pos, true);
            if(getController().getDiagram().isOutlineProp()){
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).setState(true);
                setPropsDialogLineWidthControls(true);
            } else {
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noOutlineOptionButton"))).setState(true);
                setPropsDialogLineWidthControls(false);
            }
        }
        if(!isShownPropsDialog()){
            //text controls
            //check font properties in shape and store in Diagram properties
            //getController().getDiagram().setFontPropertyValues();
            boolean isFitText = getController().getDiagram().isTextFitProp();
            if(isFitText){
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("textFitOptionButton"))).setState(true);
            }else{
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("fontSizeOptionButton"))).setState(true);
                XListBox fontSizeLB = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("fontSizeListBox"));
                //short selectedItemPos = fontSizeLB.getSelectedItemPos();
                short index = FontSize._getIndexOfFontSize(getController().getDiagram().getFontSizeProp());
                fontSizeLB.selectItemPos(index, true);
                String label = fontSizeLB.getSelectedItem();
                if(!label.startsWith("*"))
                    label = "*" + label.substring(1);
                fontSizeLB.removeItems(index, (short)1);
                fontSizeLB.addItem(label, index);
                fontSizeLB.selectItemPos(index, true);

            }
            enableFontSizeListBox(!isFitText);
            ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).setState((short)0);
            enableTextColorImageControl(false);

            //connector controls
            boolean isShownConnectors = getController().getDiagram().isShownConnectorsProp();
            enablePropsDialogConnectorControls(isShownConnectors);
            if(isShownConnectors){
                short connType = getController().getDiagram().getConnectorTypeProp();
                if(connType == 0)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("standardConnOptionButton"))).setState(true);
                if(connType == 1)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("lineConnOptionButton"))).setState(true);
                if(connType == 2)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("straightConnOptionButton"))).setState(true);
                if(connType == 3)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("curvedConnOptionButton"))).setState(true);
                if(getController().getDiagram().isConnectorStartArrowProp())
                    ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("startArrowsCheckBox"))).setState((short)1);
                else
                    ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("startArrowsCheckBox"))).setState((short)0);
                if(getController().getDiagram().isConnectorEndArrowProp())
                    ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("endArrowsCheckBox"))).setState((short)1);
                else
                    ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("endArrowsCheckBox"))).setState((short)0);
                short pos = 0;
                int connLineWidth = getController().getDiagram().getConnectorsLineWidhtProp();
                if(connLineWidth != 0)
                    pos = (short)(connLineWidth / 100);
                ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("connLineWidthListBox"))).selectItemPos(pos, true);

                setConnImageControlColor();
            }

            setMainItemHiddenCheckBox();
            enableLastHorLevelControls(getController().getDiagramType() == Controller.ORGANIGRAM);
            setLastHorLevelListBoxPos();
        }
    }
    
    public void setConnImageControlColor(){
        setConnImageControlColor(getController().getDiagram().getConnectorColorProp());
    }
    
    public void setConnImageControlColor(int color){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog); 
            setImageColorOfControl(xControlContainer.getControl("connColorImageControl"), color);
        }
    }
 
    public int getConnImageControlColor(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return getImageColorOfControl(xControlContainer.getControl("connColorImageControl"));
        }
        return 0x000000;
    }
    
    public void enableLastHorLevelControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl9"), bool);
            enableControl(xControlContainer.getControl("lastHorLevelListBox"), bool);
        }
    }
    
    public void setLastHorLevelListBoxPos(){
        setLastHorLevelListBoxPos(OrgChartTree.LASTHORLEVEL);
    }
    
    public void setLastHorLevelListBoxPos(short pos){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lastHorLevelListBox"))).selectItemPos(pos, true); 
        }
    }
    
    public short getLastHorLevelListBoxPos(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lastHorLevelListBox"))).getSelectedItemPos();
        }
        return -1;
    }
    
    public void setMainItemHiddenCheckBox(){
        setMainItemHiddenCheckBox(((OrganizationChart)(getController().getDiagram())).isHiddenRootElementProp());
    }
    
    public void setMainItemHiddenCheckBox(boolean bool){
        if(m_xPropsDialog != null){
            short sBool = 0;
            if(bool)
                sBool = 1;
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("mainItemHiddenCheckBox"))).setState(sBool);
        }
    }

    //common method
    public boolean isStateOfPropsDialogModifyColorCheckBox(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return (((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).getState() == (short)1);
        }
        return false;
    }

    public void enablePropertiesFieldOfOrganigroupPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false)
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }

    public  void setDescriptionLabelOfOrganigroupPD(String area, String modifiesColor, String colorMode, String rounded, String outline, String shadow){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7"))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8"))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9"))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10"))).setText(rounded);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11"))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12"))).setText(shadow);
        }
    }

    public void enableColorFunctionFieldOfOrganigroupPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("propsLabel3"), bool);
        }
    }

    public void setStartAndEndColorControlsInPropsDialog(){
        setStartAndEndColorControlsInPropsDialog(getController().getDiagram().getStartColorProp(), getController().getDiagram().getEndColorProp());
    }

    public void setStartAndEndColorControlsInPropsDialog(int startColor, int endColor){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            setImageColorOfControl(xControlContainer.getControl("startColorImageControl"), startColor);
            setImageColorOfControl(xControlContainer.getControl("endColorImageControl"), endColor);
        }
    }

    public void enableOrganigramColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl1"), bool);
            enableControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableControl(xControlContainer.getControl("siblingsOptionButton"), bool);
            enableControl(xControlContainer.getControl("branchOptionButton"), bool);
            enableControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableControl(xControlContainer.getControl("colorImageControl"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);
            enableControl(xControlContainer.getControl("gradientsOptionButton"), bool);
            enableControl(xControlContainer.getControl("startColorLabel"), bool);
            enableControl(xControlContainer.getControl("startColorImageControl"), bool);
            enableControl(xControlContainer.getControl("endColorLabel"), bool);
            enableControl(xControlContainer.getControl("endColorImageControl"), bool);
            enableControl(xControlContainer.getControl("gradientDirectionLabel"), bool);
            enableControl(xControlContainer.getControl("gradientDirectionListBox"), bool);
            
            if(bool){
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).getState())
                    setOrganigramPDColorControlsInColorMode();
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState())
                    setOrganigramPDColorControlsInPreDefinedColorThemeGradinetMode();
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState())
                    setOrganigramPDColorControlsInPreDefinedSchemeMode();
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("gradientsOptionButton"))).getState())
                    setOrganigramPDColorControlsInGradientMode();
            }
        }
    }

    public void setOrganigramPDColorRadioButtons(){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        boolean isAllDiagramModify = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).getState();
        if(isAllDiagramModify){
            enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), true);
            enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), true);
        }else{
            boolean isPreDefinedThemeMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState();
            boolean isPreDefinedSchemeMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState();
            if(isPreDefinedThemeMode || isPreDefinedSchemeMode){
                enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
                enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
                enableControl(xControlContainer.getControl("colorImageControl"), true);
            }
            enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), false);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), false);
        }
    }

    public void setOrganigramPDColorControlsInColorMode(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("colorImageControl"), true);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
            enableControl(xControlContainer.getControl("startColorLabel"), false);
            enableControl(xControlContainer.getControl("startColorImageControl"), false);
            enableControl(xControlContainer.getControl("endColorLabel"), false);
            enableControl(xControlContainer.getControl("endColorImageControl"), false);
            enableControl(xControlContainer.getControl("gradientDirectionLabel"), false);
            enableControl(xControlContainer.getControl("gradientDirectionListBox"), false);
        }
    }

    public void setOrganigramPDColorControlsInPreDefinedColorThemeGradinetMode(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("colorImageControl"), false);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), true);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
            enableControl(xControlContainer.getControl("startColorLabel"), false);
            enableControl(xControlContainer.getControl("startColorImageControl"), false);
            enableControl(xControlContainer.getControl("endColorLabel"), false);
            enableControl(xControlContainer.getControl("endColorImageControl"), false);
            enableControl(xControlContainer.getControl("gradientDirectionLabel"), false);
            enableControl(xControlContainer.getControl("gradientDirectionListBox"), false);
        }
    }
    
    public void setOrganigramPDColorControlsInPreDefinedSchemeMode(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("colorImageControl"), false);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), true);
            enableControl(xControlContainer.getControl("startColorLabel"), false);
            enableControl(xControlContainer.getControl("startColorImageControl"), false);
            enableControl(xControlContainer.getControl("endColorLabel"), false);
            enableControl(xControlContainer.getControl("endColorImageControl"), false);
            enableControl(xControlContainer.getControl("gradientDirectionLabel"), false);
            enableControl(xControlContainer.getControl("gradientDirectionListBox"), false);
        }
    }

    public void setOrganigramPDColorControlsInGradientMode(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
            enableControl(xControlContainer.getControl("colorImageControl"), false);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
            enableControl(xControlContainer.getControl("startColorLabel"), true);
            enableControl(xControlContainer.getControl("startColorImageControl"), true);
            enableControl(xControlContainer.getControl("endColorLabel"), true);
            enableControl(xControlContainer.getControl("endColorImageControl"), true);
            enableControl(xControlContainer.getControl("gradientDirectionLabel"), true);
            enableControl(xControlContainer.getControl("gradientDirectionListBox"), true);
        }
    }

    public XControl getStartColorImageControl(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return xControlContainer.getControl("startColorImageControl");
        }
        return null;
    }

    public XControl getEndColorImageControl(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return xControlContainer.getControl("endColorImageControl");
        }
        return null;
    }

    //common method with organigroups PD and VennDiagram PD and CycleDiagram PD
    public void setPropsDialogLineWidthControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("lineWidthLabel"), bool);
            enableControl(xControlContainer.getControl("lineWidthListBox"), bool);
        }
    }
    
    //common method
    public void enableFontSizeListBox(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("fontSizeListBox"), bool);
        }
    }

    //common method
    public void enableTextColorImageControl(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("textColorImageControl"), bool);
        }
    }

    public XControl getTextColorImageControl(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            return xControlContainer.getControl("textColorImageControl");
        }
        return null;
    }
    
    public void enablePropsDialogConnectorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl7"), bool);
            enableControl(xControlContainer.getControl("standardConnOptionButton"), bool);
            enableControl(xControlContainer.getControl("lineConnOptionButton"), bool);
            enableControl(xControlContainer.getControl("straightConnOptionButton"), bool);
            enableControl(xControlContainer.getControl("curvedConnOptionButton"), bool);
            enableControl(xControlContainer.getControl("arrowsLabel"), bool);
            enableControl(xControlContainer.getControl("startArrowsCheckBox"), bool);
            enableControl(xControlContainer.getControl("endArrowsCheckBox"), bool);
            enableControl(xControlContainer.getControl("connLineWidthLabel"), bool);
            enableControl(xControlContainer.getControl("connLineWidthListBox"), bool);
        }
    }

    public void setPropertiesOfOrganingram(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != OrganizationChart.USER_DEFINE){
                if(sNewStyle == OrganizationChart.DEFAULT)                    //selectAllShape, modifyColors, rounded,                outline, lineWidth,          shadow
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, false,    Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                if(sNewStyle == OrganizationChart.WITHOUT_OUTLINE)
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, false, Diagram.MEDIUM_ROUNDED, false, Diagram.LINE_WIDTH100, false);
                if(sNewStyle == OrganizationChart.NOT_ROUNDED)
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, false, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                if(sNewStyle == OrganizationChart.WITH_SHADOW)
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, false, Diagram.MEDIUM_ROUNDED, false, Diagram.LINE_WIDTH100, true);
                if(getController().getDiagram().isColorThemeGradientStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeGradientStyle(sNewStyle);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, true, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getEndColorProp());
                    getController().getDiagram().setConnectorColorProp(getController().getDiagram().getEndColorProp());
                    if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                        getController().getDiagram().setGradientDirectionProp(Diagram.HORIZONTAL);
                    else
                        getController().getDiagram().setGradientDirectionProp(Diagram.VERTICAL);
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    getController().getDiagram().setColorModeProp(colorMode);
                    ((OrganizationChart)getController().getDiagram()).setDiagramPropertyValues(true, true, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    getController().getDiagram().setConnectorColorProp(getController().getDiagram().getDefalutConnectorColor());
                    if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                        getController().getDiagram().setGradientDirectionProp(Diagram.HORIZONTAL);
                    else
                        getController().getDiagram().setGradientDirectionProp(Diagram.VERTICAL);
                }
            } else {
                boolean modifyColors = false;
                if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).getState() == 1)
                    modifyColors = true;
                if(modifyColors){
                    getController().getDiagram().setModifyColorsProp(true);
                    boolean isSelectedAllDiagram = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).getState();
                    if(isSelectedAllDiagram){
                        getController().getDiagram().setSelectedAllShapesProp(true);
                    }else{
                        getController().getDiagram().setSelectedAllShapesProp(false);
                        if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("selectedItemsOptionButton"))).getState())
                            getController().getDiagram().setSelectedAreaProp(Diagram.SELECTED_SHAPES);
                        if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("siblingsOptionButton"))).getState())
                            getController().getDiagram().setSelectedAreaProp(Diagram.SIBLING_SHAPES);
                        if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("branchOptionButton"))).getState())
                            getController().getDiagram().setSelectedAreaProp(Diagram.BRANCH_SHAPES);
                    }

                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).getState()){
                        getController().getDiagram().setColorProp(getImageColorOfControl(xControlContainer.getControl("colorImageControl")));
                        getController().getDiagram().setColorModeProp(Diagram.SIMPLE_COLOR_MODE);
                        getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState()){
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).getSelectedItemPos();
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE));
                        getController().getDiagram().setColorThemeGradientColors();
                        getController().getDiagram().setLineColorProp(getController().getDiagram().getEndColorProp());
                        getController().getDiagram().setConnectorColorProp(getController().getDiagram().getEndColorProp());
                        if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                            getController().getDiagram().setGradientDirectionProp(Diagram.HORIZONTAL);
                        else
                            getController().getDiagram().setGradientDirectionProp(Diagram.VERTICAL);
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORSCHEME_MODE_VALUE));
                        getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("gradientsOptionButton"))).getState()){
                        getController().getDiagram().setStartColorProp(getImageColorOfControl(xControlContainer.getControl("startColorImageControl")));
                        getController().getDiagram().setEndColorProp(getImageColorOfControl(xControlContainer.getControl("endColorImageControl")));
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("gradientDirectionListBox"))).getSelectedItemPos();
                        getController().getDiagram().setGradientDirectionProp(selectedItemPos);
                    }
                }else{
                    getController().getDiagram().setModifyColorsProp(false);
                }

                boolean isShadow = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).getState();
                getController().getDiagram().setShadowProp(isShadow);
                boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
                getController().getDiagram().setOutlineProp(isOutline);
                if(isOutline){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                    getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
                }
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noRoundedOptionButton"))).getState())
                    getController().getDiagram().setRoundedProp(Diagram.NULL_ROUNDED);
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumRoundedOptionButton"))).getState())
                    getController().getDiagram().setRoundedProp(Diagram.MEDIUM_ROUNDED);
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraRoundedOptionButton"))).getState())
                    getController().getDiagram().setRoundedProp(Diagram.EXTRA_ROUNDED);
            }

            setTextProperties();

            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("standardConnOptionButton"))).getState())
                getController().getDiagram().setConnectorTypeProp(Diagram.CONN_STANDARD);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("lineConnOptionButton"))).getState())
                getController().getDiagram().setConnectorTypeProp(Diagram.CONN_LINE);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("straightConnOptionButton"))).getState())
                getController().getDiagram().setConnectorTypeProp(Diagram.CONN_STRAIGHT);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("curvedConnOptionButton"))).getState())
                getController().getDiagram().setConnectorTypeProp(Diagram.CONN_CURVED);
            boolean isStartConnectors = false;
            if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("startArrowsCheckBox"))).getState() == 1)
                isStartConnectors = true;
            getController().getDiagram().setConnectorStartArrowProp(isStartConnectors);
            boolean isEndConnectors = false;
            if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("endArrowsCheckBox"))).getState() == 1)
                isEndConnectors = true;
            getController().getDiagram().setConnectorEndArrowProp(isEndConnectors);

            short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("connLineWidthListBox"))).getSelectedItemPos();
            getController().getDiagram().setConnectorsLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));

            getController().getDiagram().setConnectorColorProp(getConnImageControlColor());

            //main item
            boolean isMainItemHidden = false;
            if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("mainItemHiddenCheckBox"))).getState() == 1)
                isMainItemHidden = true;
            ((OrganizationChart)(getController().getDiagram())).setHiddenRootElementProp(isMainItemHidden);

            //lastHorizontal level in OrgChart
            if(getController().getDiagramType() == Controller.ORGANIGRAM)  
                ((OrgChartTree)((OrgChart)(getController().getDiagram())).getDiagramTree()).setLastHorLevel(getLastHorLevelListBoxPos());
        }
    }

    public void setTextProperties(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XRadioButton textFitRB = (XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("textFitOptionButton"));
            XListBox fontSizeLB = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("fontSizeListBox"));
            boolean isFitText = textFitRB.getState();
            getController().getDiagram().setTextFitProp(isFitText);

            float newFontSize = 0;
            if(!isFitText)
                newFontSize = FontSize._getFontSize(fontSizeLB.getSelectedItemPos());
            //set no mark in FontSize ListBox
            short index = FontSize._getIndexOfFontSize(getController().getDiagram().getFontSizeProp());
            fontSizeLB.selectItemPos(index, true);
            String label = fontSizeLB.getSelectedItem();
            if(label.startsWith("*"))
                label = " " + label.substring(1);
            fontSizeLB.removeItems(index, (short)1);
            fontSizeLB.addItem(label, index);
            fontSizeLB.selectItemPos(index, true);
            if(!isFitText)
                getController().getDiagram().setFontSizeProp(newFontSize);

            boolean isTextColorChange = isCheckedModifyTextColorCheckBox();
            getController().getDiagram().setTextColorChange(isTextColorChange);
            if(isTextColorChange)
                getController().getDiagram().setTextColorProp(getImageColorOfControl(getTextColorImageControl()));
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

    //common method with organigroups PD and VennDiagram PD and CycleDiagram PD
    public void setTextColorToolsProps(){
        boolean state = isCheckedModifyTextColorCheckBox();
        enableTextColorImageControl(state);
    }

    public void disposePropertiesDialog(){
        if(m_xPropsDialog != null){
            m_xPropsDialog.endExecute();
            setShownPropsDialog(false);
            XComponent xComponent = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xPropsDialog);
            if(xComponent != null)
                xComponent.dispose();
            m_xPropsDialog = null;
            enableControlDialogWindow(true);
            setFocusControlDialog();
        }
    }

    //------------------------------------------------------------------------------------------
    
}
