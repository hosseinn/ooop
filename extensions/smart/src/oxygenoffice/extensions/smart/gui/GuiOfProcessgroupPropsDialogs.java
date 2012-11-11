package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XRadioButton;
import com.sun.star.frame.XFrame;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.FontSize;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.diagram.processes.bendingprocess.BendingProcess;
import oxygenoffice.extensions.smart.diagram.processes.continuousblockprocess.ContinuousBlockProcess;
import oxygenoffice.extensions.smart.diagram.processes.staggeredprocess.StaggeredProcess;
import oxygenoffice.extensions.smart.diagram.processes.upwardarrowprocess.UpwardArrowProcess;


public class GuiOfProcessgroupPropsDialogs extends GuiOfRelationgroupPropsDialogs {
    
    public GuiOfProcessgroupPropsDialogs(){ }
    
    public GuiOfProcessgroupPropsDialogs(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
    }
    
    //Common methods ****************************************************************************************************************************
    public void setColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool){
                ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
                if(getController().getDiagram().isBaseColorsMode())
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).setState(true);
                if(getController().getDiagram().isSimpleColorMode())
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
                if(getController().getDiagram().isColorThemeGradientMode())
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
                if(getController().getDiagram().isColorSchemeMode())
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).setState(true);
                setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
                short pos = 0;
                if(getController().getDiagram().isColorThemeGradientMode())
                    pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
                ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
                pos = 0;
                if(getController().getDiagram().isColorSchemeMode())
                    pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE);
                ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).selectItemPos(pos, true);
            }
        }
    }
    
    public void setOutlineControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool){
                short pos = 0;
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
        }
    }
    
    public void setShadowControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool){
                if(getController().getDiagram().isShadowProp())
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).setState(true);
                else
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noShadowOptionButton"))).setState(true);
            }
        }
    }
    
    public void setRoundedControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool){
                if(getController().getDiagram().getRoundedProp() == Diagram.NULL_ROUNDED)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noRoundedOptionButton"))).setState(true);
                if(getController().getDiagram().getRoundedProp() == Diagram.MEDIUM_ROUNDED)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumRoundedOptionButton"))).setState(true);
                if(getController().getDiagram().getRoundedProp() == Diagram.EXTRA_ROUNDED)
                    ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraRoundedOptionButton"))).setState(true);
            }
        }
    }
    
    public void setTextControls(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
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
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).setState((short)0);
            enableTextColorImageControl(false);
        }
    }
    
    public void setOutlineProperties(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
            getController().getDiagram().setOutlineProp(isOutline);
            if(isOutline){
                short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
            }
        }
    }
    
    public void setShadowProperties(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).getState())
                getController().getDiagram().setShadowProp(true);
            else
                getController().getDiagram().setShadowProp(false);
        }
    }
    
    public void setRoundedProperties(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noRoundedOptionButton"))).getState())
                getController().getDiagram().setRoundedProp(Diagram.NULL_ROUNDED);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumRoundedOptionButton"))).getState())
                getController().getDiagram().setRoundedProp(Diagram.MEDIUM_ROUNDED);
            if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraRoundedOptionButton"))).getState())
                getController().getDiagram().setRoundedProp(Diagram.EXTRA_ROUNDED);
        }
    }
    
    public void enableColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);         
            enableControl(xControlContainer.getControl("areaFrameControl"), bool);
            enableControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableControl(xControlContainer.getControl("colorImageControl"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);
        }
    }
    
    public boolean setColorProperties(){
        boolean modifyColors = false;
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).getState() == 1)
                modifyColors = true;
            if(modifyColors){
                getController().getDiagram().setModifyColorsProp(true);
                boolean isSelectedAllDiagram = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).getState();
                getController().getDiagram().setSelectedAllShapesProp(isSelectedAllDiagram);

                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).getState()){
                    getController().getDiagram().setColorModeProp(Diagram.BASE_COLORS_MODE);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }

                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).getState()){
                    getController().getDiagram().setColorProp(getImageColorOfControl(xControlContainer.getControl("colorImageControl")));
                    getController().getDiagram().setColorModeProp(Diagram.SIMPLE_COLOR_MODE);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState()){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).getSelectedItemPos();
                    getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE));
                    getController().getDiagram().setColorThemeColors();
                }
                if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                    getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORSCHEME_MODE_VALUE));
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
            }else{
                getController().getDiagram().setModifyColorsProp(false);
            }
        }
        return modifyColors;
    }
    
    public void enableDisableColorControls(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).getState())
                enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), true);
            else
                enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), false);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).getState())
                enableControl(xControlContainer.getControl("colorImageControl"), true);
            else
                enableControl(xControlContainer.getControl("colorImageControl"), false);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState())
                enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), true);
            else
                enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState())
                enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), true);
            else
                enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
        }
    }
    
    public void setColorRadioButtons(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            boolean isAllDiagramModify = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).getState();
            if(isAllDiagramModify){
                enableControl(xControlContainer.getControl("baseColorsOptionButton"), true);
                enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), true);
                enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), true);
            }else{
                boolean isBaseColorsMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).getState();
                if(isBaseColorsMode)
                    enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), false);
                boolean isPreDefinedColorThemeMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState();
                if(isPreDefinedColorThemeMode)
                    enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
                boolean isPreDefinedColorSchemeMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState();
                if(isPreDefinedColorSchemeMode)
                    enableControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), false);
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
                enableControl(xControlContainer.getControl("colorImageControl"), true);
                enableControl(xControlContainer.getControl("baseColorsOptionButton"), false);
                enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), false);
                enableControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), false);
            }
        }
    }
    
    public void setArrowColorControls(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            boolean setState = true;
            short selectedItemPos = 0;
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState()){
                ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).setState((short)1);
                selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).getSelectedItemPos();
                setImageColorOfControl(xControlContainer.getControl("arrowColorImageControl"), Diagram.aLOCOLORS[selectedItemPos][1]);
                setState = false;
            }
            if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).setState((short)1);
                selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                setImageColorOfControl(xControlContainer.getControl("arrowColorImageControl"), SchemeDefinitions.aColorSchemes[selectedItemPos][1]);
                setState = false;
            }
            if(setState)
                enableDisableArrowColorImageControl();
        }
    }
    
    public void enableDisableArrowColorImageControl(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).getState() == 0)
                enableControl(xControlContainer.getControl("arrowColorImageControl"), false);
            else
                enableControl(xControlContainer.getControl("arrowColorImageControl"), true);
        }
    }
    //*******************************************************************************************************************************************
   
    
    //ContinuousBlockProcess Properties Dialog **************************************************************************************************
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setContinuousBlockProcessPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != ContinuousBlockProcess.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            
            enableVisibleContinuousBlockProcessPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfContinuousBlockProcessPD(true);
            setContinuousBlockProcessPropsControls(false);
            
            if(sListPos == ContinuousBlockProcess.DEFAULT){
                setGraphic(xImageControl, "/images/continuousblock_default.png");
                //                                              area        rounded  modifiesColor   colorMode    outline   shadow
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localBaseColors, localYes, localNo);
            }
            if (sListPos == ContinuousBlockProcess.NOT_ROUNDED){
                setGraphic(xImageControl, "/images/continuousblock_notRounded.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localNo, localYes, localBaseColors, localYes, localNo);
            }
            if (sListPos == ContinuousBlockProcess.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/continuousblock_withoutOutline.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localBaseColors, localNo, localNo);
            }
            if (sListPos == ContinuousBlockProcess.WITH_SHADOW){
                setGraphic(xImageControl, "/images/continuousblock_withShadow.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localBaseColors, localYes, localYes);
            }
            if(sListPos == ContinuousBlockProcess.GREEN_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_greenDark.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.GREEN_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_greenBright.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.BLUE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_blueDark.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.BLUE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_blueBright.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.PURPLE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_purpleDark.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.PURPLE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_purpleBright.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.ORANGE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_orangeDark.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.ORANGE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_orangeBright.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.YELLOW_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_yellowDark.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.YELLOW_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/continuousblock_yellowBright.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_blueGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_aquaGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.RED_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_redGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_fireGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.SUN_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_sunGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_greenGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_oliveGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_purpleGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.PINK_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_pinkGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_indianGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_maroonGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == ContinuousBlockProcess.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/continuousblock_brownGradients.png");
                setDescriptionLabelOfContinuousBlockProcessPD(localAllShape, localMedium, localYes, localColorScheme, localYes, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleContinuousBlockProcessPropsControls(true);
            enablePropertiesFieldOfContinuousBlockProcessPD(false);
            setContinuousBlockProcessPropsControls(true);
        }
    }
    
    public  void setDescriptionLabelOfContinuousBlockProcessPD(String area, String rounded, String modifiesColor, String colorMode, String outline, String shadow) {
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(rounded);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(shadow);
        }
    }
    
    public void enableVisibleContinuousBlockProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

            enableVisibleControl(xControlContainer.getControl("colorsFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("areaFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("arrowFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyArrowShapeColorCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("arrowColorImageControl"), bool);

            enableVisibleControl(xControlContainer.getControl("outlineFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("shadowFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("roundedFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("noRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("mediumRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("extraRoundedOptionButton"), bool);
        }
    }
    
    public void enablePropertiesFieldOfContinuousBlockProcessPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("propsFrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }
        
    public void setContinuousBlockProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            if(bool){
                setColorControls(bool);
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
                ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).setState((short)0);
                setImageColorOfControl(xControlContainer.getControl("arrowColorImageControl"), ((ContinuousBlockProcess)getController().getDiagram()).getArrowShapeColorProp());
                enableContinuousBlockProcessColorControls(false);
                
                setOutlineControls(bool);
                setShadowControls(bool);
                setRoundedControls(bool);
            }
            if(!isShownPropsDialog())
                setTextControls();
        }
    }
        
    public void enableContinuousBlockProcessColorControls(boolean bool){
        if(m_xPropsDialog != null){
            enableColorControls(bool);
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);         
            enableControl(xControlContainer.getControl("arrowFrameControl"), bool);
            enableControl(xControlContainer.getControl("modifyArrowShapeColorCheckBox"), bool);
            enableControl(xControlContainer.getControl("arrowColorImageControl"), bool);
            if(bool)
                setContinuousBlockProcessPDColorControls();
        }
    }
    
    public void setContinuousBlockProcessPDColorControls(){
        enableDisableColorControls();
        setArrowColorControls();
    }
    
    public void setContinuousBlockProcessPDColorRadioButtons(){
        setColorRadioButtons();
    }
    
    public void setPropertiesOfContinuousBlockProcess(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != ContinuousBlockProcess.USER_DEFINE){
                if(sNewStyle == ContinuousBlockProcess.DEFAULT){                     //selectAllShape, modifyColors, colorMode      sRounded,              outline,    line width,       shadow
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(sNewStyle == ContinuousBlockProcess.NOT_ROUNDED){
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(sNewStyle == ContinuousBlockProcess.WITHOUT_OUTLINE){
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, false, Diagram.LINE_WIDTH000, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(sNewStyle == ContinuousBlockProcess.WITH_SHADOW){
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(getController().getDiagram().isColorThemeGradientStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeGradientStyle(sNewStyle);
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((ContinuousBlockProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
            } else {
                boolean isModifyColors = setColorProperties();
                if(isModifyColors)
                    if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).getState() == 1)
                        ((ContinuousBlockProcess)getController().getDiagram()).setArrowShapeColorProp(getImageColorOfControl(xControlContainer.getControl("arrowColorImageControl")));
                setOutlineProperties();
                setShadowProperties();
                setRoundedProperties();
            }
            setTextProperties();
        }
    }
    //*************************************************************************************************************************************
     
    
    //StaggeredProcess Properties Dialog **************************************************************************************************
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setStaggeredProcessPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != StaggeredProcess.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            
            enableVisibleStaggeredProcessPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfStaggeredProcessPD(true);
            setStaggeredProcessPropsControls(false);
            
            if(sListPos == StaggeredProcess.DEFAULT){
                setGraphic(xImageControl, "/images/staggered_default.png");
                //                                       area       modifiesColor   colorMode     outline   shadow   rounded
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localBaseColors, localYes, localNo, localNo);
            }
            if (sListPos == StaggeredProcess.ROUNDED){
                setGraphic(xImageControl, "/images/staggered_rounded.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localBaseColors, localYes, localNo, localMedium);
            }
            if (sListPos == StaggeredProcess.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/staggered_withoutOutline.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localBaseColors, localNo, localNo, localNo);
            }
            if (sListPos == StaggeredProcess.WITH_SHADOW){
                setGraphic(xImageControl, "/images/staggered_withShadow.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localBaseColors, localYes, localYes, localMedium);
            }
            if(sListPos == StaggeredProcess.GREEN_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_greenDark.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.GREEN_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_greenBright.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.BLUE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_blueDark.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.BLUE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_blueBright.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.PURPLE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_purpleDark.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.PURPLE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_purpleBright.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.ORANGE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_orangeDark.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.ORANGE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_orangeBright.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.YELLOW_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_yellowDark.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.YELLOW_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/staggered_yellowBright.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/staggered_blueGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/staggered_aquaGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.RED_SCHEME){
                setGraphic(xImageControl, "/images/staggered_redGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/staggered_fireGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.SUN_SCHEME){
                setGraphic(xImageControl, "/images/staggered_sunGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/staggered_greenGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/staggered_oliveGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/staggered_purpleGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.PINK_SCHEME){
                setGraphic(xImageControl, "/images/staggered_pinkGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/staggered_indianGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/staggered_maroonGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == StaggeredProcess.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/staggered_brownGradients.png");
                setDescriptionLabelOfStaggeredProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleStaggeredProcessPropsControls(true);
            enablePropertiesFieldOfStaggeredProcessPD(false);
            setStaggeredProcessPropsControls(true);
        }
    }
    
    public  void setDescriptionLabelOfStaggeredProcessPD(String area, String modifiesColor, String colorMode, String outline, String shadow, String rounded) {
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(shadow);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(rounded);
        }
    }
    
    public void enableVisibleStaggeredProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

            enableVisibleControl(xControlContainer.getControl("colorsFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("areaFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);

            enableVisibleControl(xControlContainer.getControl("outlineFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("shadowFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("roundedFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("noRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("mediumRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("extraRoundedOptionButton"), bool);
        }
    }
    
    public void enablePropertiesFieldOfStaggeredProcessPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("propsFrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }
        
    public void setStaggeredProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            if(bool){
                setColorControls(bool);
                enableStaggeredProcessColorControls(false);
                setOutlineControls(bool);
                setShadowControls(bool);
                setRoundedControls(bool);
            }
            if(!isShownPropsDialog())
                setTextControls();
        }
    }
        
    public void enableStaggeredProcessColorControls(boolean bool){
        enableColorControls(bool);
        if(bool)
            setStaggeredProcessPDColorControls();
    }
    
    public void setStaggeredProcessPDColorControls(){
        enableDisableColorControls();
    }
    
    public void setStaggeredProcessPDColorRadioButtons(){
        setColorRadioButtons();
    }
    
    public void setPropertiesOfStaggeredProcess(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != StaggeredProcess.USER_DEFINE){
                if(sNewStyle == StaggeredProcess.DEFAULT){                     //selectAllShape, modifyColors, colorMode      sRounded,              outline,    line width,       shadow
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == StaggeredProcess.ROUNDED){
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == StaggeredProcess.WITHOUT_OUTLINE){
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.NULL_ROUNDED, false, Diagram.LINE_WIDTH000, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == StaggeredProcess.WITH_SHADOW){
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeGradientStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeGradientStyle(sNewStyle);
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((StaggeredProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                }
            } else {
                setColorProperties();
                setOutlineProperties();
                setShadowProperties();
                setRoundedProperties();
            }
            
            setTextProperties();
        }
    }
    //*************************************************************************************************************************************


    //StaggeredProcess Properties Dialog **************************************************************************************************
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setBendingProcessPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != BendingProcess.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            
            enableVisibleBendingProcessPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfBendingProcessPD(true);
            setBendingProcessPropsControls(false);
            
            if(sListPos == BendingProcess.DEFAULT){
                setGraphic(xImageControl, "/images/bending_default.png");
                //                                       area       modifiesColor   colorMode     outline   shadow   rounded
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localBaseColors, localYes, localNo, localNo);
            }
            if (sListPos == BendingProcess.ROUNDED){
                setGraphic(xImageControl, "/images/bending_rounded.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localBaseColors, localYes, localNo, localMedium);
            }
            if (sListPos == BendingProcess.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/bending_withoutOutline.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localBaseColors, localNo, localNo, localNo);
            }
            if (sListPos == BendingProcess.WITH_SHADOW){
                setGraphic(xImageControl, "/images/bending_withShadow.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localBaseColors, localYes, localYes, localMedium);
            }
            if(sListPos == BendingProcess.GREEN_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/bending_greenDark.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.GREEN_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/bending_greenBright.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.BLUE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/bending_blueDark.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.BLUE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/bending_blueBright.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.PURPLE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/bending_purpleDark.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.PURPLE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/bending_purpleBright.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.ORANGE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/bending_orangeDark.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.ORANGE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/bending_orangeBright.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.YELLOW_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/bending_yellowDark.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.YELLOW_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/bending_yellowBright.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/bending_blueGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/bending_aquaGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.RED_SCHEME){
                setGraphic(xImageControl, "/images/bending_redGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/bending_fireGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.SUN_SCHEME){
                setGraphic(xImageControl, "/images/bending_sunGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/bending_greenGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/bending_oliveGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/bending_purpleGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.PINK_SCHEME){
                setGraphic(xImageControl, "/images/bending_pinkGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/bending_indianGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/bending_maroonGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == BendingProcess.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/bending_brownGradients.png");
                setDescriptionLabelOfBendingProcessPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleBendingProcessPropsControls(true);
            enablePropertiesFieldOfBendingProcessPD(false);
            setBendingProcessPropsControls(true);
        }
    }
    
    public  void setDescriptionLabelOfBendingProcessPD(String area, String modifiesColor, String colorMode, String outline, String shadow, String rounded) {
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(shadow);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(rounded);
        }
    }
    
    public void enableVisibleBendingProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

            enableVisibleControl(xControlContainer.getControl("colorsFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("areaFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);

            enableVisibleControl(xControlContainer.getControl("outlineFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("shadowFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

            enableVisibleControl(xControlContainer.getControl("roundedFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("noRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("mediumRoundedOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("extraRoundedOptionButton"), bool);
        }
    }
    
    public void enablePropertiesFieldOfBendingProcessPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("propsFrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }
        
    public void setBendingProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            if(bool){
                setColorControls(bool);
                enableBendingProcessColorControls(false);
                setOutlineControls(bool);
                setShadowControls(bool);
                setRoundedControls(bool);
            }
            if(!isShownPropsDialog())
                setTextControls();
        }
    }
        
    public void enableBendingProcessColorControls(boolean bool){
        enableColorControls(bool);
        if(bool)
            setBendingProcessPDColorControls();
    }
    
    public void setBendingProcessPDColorControls(){
        enableDisableColorControls();
    }
    
    public void setBendingProcessPDColorRadioButtons(){
        setColorRadioButtons();
    }
    
    public void setPropertiesOfBendingProcess(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != BendingProcess.USER_DEFINE){
                if(sNewStyle == BendingProcess.DEFAULT){                     //selectAllShape, modifyColors, colorMode      sRounded,              outline,    line width,       shadow
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == BendingProcess.ROUNDED){
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == BendingProcess.WITHOUT_OUTLINE){
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.NULL_ROUNDED, false, Diagram.LINE_WIDTH000, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == BendingProcess.WITH_SHADOW){
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_ROUNDED, true, Diagram.LINE_WIDTH100, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeGradientStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeGradientStyle(sNewStyle);
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((BendingProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, Diagram.NULL_ROUNDED, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                }
            } else {
                setColorProperties();
                setOutlineProperties();
                setShadowProperties();
                setRoundedProperties();
            }
            
            setTextProperties();
        }
    }
    //*************************************************************************************************************************************
    
    
    //UpwardArrowProcess Properties Dialog **************************************************************************************************
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setUpwardArrowProcessPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != UpwardArrowProcess.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localBaseColorsGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColorsGradients.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            
            enableVisibleUpwardArrowProcessPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfUpwardArrowProcessPD(true);
            setUpwardArrowProcessPropsControls(false);
            
            if(sListPos == UpwardArrowProcess.DEFAULT){
                setGraphic(xImageControl, "/images/upwardarrow_default.png");
                //                                              area  modifiesColor   colorMode    outline 
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localBaseColors, localYes);
            }
            if (sListPos == UpwardArrowProcess.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/upwardarrow_withoutOutline.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localBaseColors, localNo);
            }
            if (sListPos == UpwardArrowProcess.BASECOLORS_WITH_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_basecolorsWithGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localBaseColorsGradients, localYes);
            }
            if(sListPos == UpwardArrowProcess.GREEN_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_greenDark.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.GREEN_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_greenBright.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.BLUE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_blueDark.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.BLUE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_blueBright.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.PURPLE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_purpleDark.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.PURPLE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_purpleBright.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.ORANGE_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_orangeDark.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.ORANGE_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_orangeBright.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.YELLOW_DARK_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_yellowDark.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.YELLOW_BRIGHT_GRADIENT){
                setGraphic(xImageControl, "/images/upwardarrow_yellowBright.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorTheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_blueGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_aquaGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.RED_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_redGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_fireGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.SUN_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_sunGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_greenGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_oliveGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_purpleGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.PINK_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_pinkGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_indianGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_maroonGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
            if(sListPos == UpwardArrowProcess.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/upwardarrow_brownGradients.png");
                setDescriptionLabelOfUpwardArrowProcessPD(localAllShape, localYes, localColorScheme, localYes);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleUpwardArrowProcessPropsControls(true);
            enablePropertiesFieldOfUpwardArrowProcessPD(false);
            setUpwardArrowProcessPropsControls(true);
        }
    }
    
    public  void setDescriptionLabelOfUpwardArrowProcessPD(String area, String modifiesColor, String colorMode, String outline) {
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel5" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel6" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(outline);
        }
    }
    
    public void enableVisibleUpwardArrowProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

            enableVisibleControl(xControlContainer.getControl("colorsFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("areaFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("preDefinedColorSchemesListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("arrowFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("modifyArrowShapeColorCheckBox"), bool);
            enableVisibleControl(xControlContainer.getControl("arrowColorImageControl"), bool);

            enableVisibleControl(xControlContainer.getControl("outlineFrameControl"), bool);
            enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
            enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
            enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);
        }
    }
    
    public void enablePropertiesFieldOfUpwardArrowProcessPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 5; i <= 8; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("propsFrameControl"), bool);
            for(int i = 1; i <= 8; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }
        
    public void setUpwardArrowProcessPropsControls(boolean bool){
        if(m_xPropsDialog != null){
            if(bool){
                setColorControls(bool);
                XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
                ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).setState((short)0);
                setImageColorOfControl(xControlContainer.getControl("arrowColorImageControl"), ((UpwardArrowProcess)getController().getDiagram()).getArrowShapeColorProp());
                enableUpwardArrowProcessColorControls(false);
                
                setOutlineControls(bool);
            }
            if(!isShownPropsDialog())
                setTextControls();
        }
    }
        
    public void enableUpwardArrowProcessColorControls(boolean bool){
        if(m_xPropsDialog != null){
            enableColorControls(bool);
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);         
            enableControl(xControlContainer.getControl("arrowFrameControl"), bool);
            enableControl(xControlContainer.getControl("modifyArrowShapeColorCheckBox"), bool);
            enableControl(xControlContainer.getControl("arrowColorImageControl"), bool);
            if(bool)
                setUpwardArrowProcessPDColorControls();
        }
    }
    
    public void setUpwardArrowProcessPDColorControls(){
        enableDisableColorControls();
        setArrowColorControls();
    }
    
    public void setUpwardArrowProcessPDColorRadioButtons(){
        setColorRadioButtons();
    }
    
    public void setPropertiesOfUpwardArrowProcess(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != UpwardArrowProcess.USER_DEFINE){
                if(sNewStyle == UpwardArrowProcess.DEFAULT){                     //selectAllShape, modifyColors, colorMode     outline,    line width
                    ((UpwardArrowProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(sNewStyle == UpwardArrowProcess.WITHOUT_OUTLINE){
                    ((UpwardArrowProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, false, Diagram.LINE_WIDTH000);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(sNewStyle == UpwardArrowProcess.BASECOLORS_WITH_GRADIENT){
                    ((UpwardArrowProcess)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_WITH_GRADIENT_MODE, true, Diagram.LINE_WIDTH100);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(getController().getDiagram().isColorThemeGradientStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeGradientStyle(sNewStyle);
                    ((UpwardArrowProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                    ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((UpwardArrowProcess)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100);
                    getController().getDiagram().setColorModeProp(colorMode);
                    ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getController().getDiagram().getArrowColor());
                }
            } else {
                boolean isModifyColors = setColorProperties();
                if(isModifyColors)
                    if(((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyArrowShapeColorCheckBox"))).getState() == 1)
                        ((UpwardArrowProcess)getController().getDiagram()).setArrowShapeColorProp(getImageColorOfControl(xControlContainer.getControl("arrowColorImageControl")));
                setOutlineProperties();
            }
            setTextProperties();
        }
    }
    //*************************************************************************************************************************************
    
    
}
