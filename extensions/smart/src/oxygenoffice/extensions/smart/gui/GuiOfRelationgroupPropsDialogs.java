
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
import oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram.CycleDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram.PyramidDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.targetdiagram.TargetDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram.VennDiagram;

public class GuiOfRelationgroupPropsDialogs extends GuiOfOrganigroupPropsDialogs {
    
    public GuiOfRelationgroupPropsDialogs(){ }
    
    public GuiOfRelationgroupPropsDialogs(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
    }
    
    //VennDiagram Properties Dialog --------------------------------------------------------------------------------------------------
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setVennDiagramPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != VennDiagram.USER_DEFINE){
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localMedium = getDialogPropertyValue("Strings2", "Strings2.Common.LocalMedium.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            
            enableVisibleVennDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfVennDiagramPD(true);
            setVennDiagramPropsControls(false);
            
            if(sListPos == VennDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/venn_default.png");
                //                                      area        outline modifiesColor   colorMode       transp  TFrame   roundedTFrame
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localBaseColors, localMedium, localYes, localYes);
            }
            if (sListPos == VennDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/venn_withoutOutline.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localNo, localYes, localBaseColors, localMedium, localYes, localYes);
            }
            if (sListPos == VennDiagram.WITHOUT_FRAME){
                setGraphic(xImageControl, "/images/venn_withoutFrame.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localBaseColors, localMedium, localNo, "");
                enableRoundedTextFieldFieldOfVennDiagramPD(false);
            }
            if (sListPos == VennDiagram.NOT_ROUNDED){
                setGraphic(xImageControl, "/images/venn_notRounded.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localBaseColors, localMedium, localYes, localNo);
            }
            if(sListPos == VennDiagram.GREEN_DARK){
                setGraphic(xImageControl, "/images/venn_greenDark.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.GREEN_BRIGHT){
                setGraphic(xImageControl, "/images/venn_greenBright.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.BLUE_DARK){
                setGraphic(xImageControl, "/images/venn_blueDark.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.BLUE_BRIGHT){
                setGraphic(xImageControl, "/images/venn_blueBright.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.PURPLE_DARK){
                setGraphic(xImageControl, "/images/venn_purpleDark.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.PURPLE_BRIGHT){
                setGraphic(xImageControl, "/images/venn_purpleBright.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.ORANGE_DARK){
                setGraphic(xImageControl, "/images/venn_orangeDark.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.ORANGE_BRIGHT){
                setGraphic(xImageControl, "/images/venn_orangeBright.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.YELLOW_DARK){
                setGraphic(xImageControl, "/images/venn_yellowDark.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
            if(sListPos == VennDiagram.YELLOW_BRIGHT){
                setGraphic(xImageControl, "/images/venn_yellowBright.png");
                setDescriptionLabelOfVennDiagramPD(localAllShape, localYes, localYes, localColorTheme, localMedium,  localYes, localYes);
            }
        }else{
            enableVisibleControl(xImageControl, false);
            enableVisibleVennDiagramPropsControls(true);
            enablePropertiesFieldOfVennDiagramPD(false);
            setVennDiagramPropsControls(true);
        }
    }

    public void enablePropertiesFieldOfVennDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 8; i <= 14; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 14; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }

    public void enableRoundedTextFieldFieldOfVennDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("propsLabel7"), bool);
        }
    }

    public  void setDescriptionLabelOfVennDiagramPD(String area, String outline, String modifiesColor, String colorMode, String transp, String TFrame, String roundedTFrame){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(transp);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel13" ))).setText(TFrame);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel14" ))).setText(roundedTFrame);
        }
    }
    
    public void enableVisibleVennDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
        enableVisibleControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
        enableVisibleControl(xControlContainer.getControl("colorOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("colorImageControl"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
        
        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("noTransparencyOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("mediumTransparencyOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("extraTransparencyOptionButton"), bool);
   
        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesFrameOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noFrameOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl6"), bool);
        enableVisibleControl(xControlContainer.getControl("yesFrameRoundedOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noFrameRoundedOptionButton"), bool);
    }
    
    public void setVennDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        
        if(bool){
            //color controls
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
            ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
            //short colorMode = getController().getDiagram().getColorModeProps();
            
            if(getController().getDiagram().isBaseColorsMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).setState(true);
            if(getController().getDiagram().isSimpleColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
            if(getController().getDiagram().isColorThemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
            
            setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
            short pos = 0;
            if(getController().getDiagram().isColorThemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
            enableVennDiagramColorControls(false);
      
            //trancparency controls
            short transparency = getController().getDiagram().getTransparencyProp();
            if(transparency == Diagram.NULL_TRANSP)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noTransparencyOptionButton"))).setState(true);
            if(transparency == Diagram.MEDIUM_TRANSP)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumTransparencyOptionButton"))).setState(true);
            if(transparency == Diagram.EXTRA_TRANSP)
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraTransparencyOptionButton"))).setState(true);
            
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

            //textframe controls
            if(getController().getDiagram().isFrameProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).setState(true);
            else 
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noFrameOptionButton"))).setState(true);
            
            //rounded textframe controls
            if(getController().getDiagram().isRoundedFrameProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameRoundedOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noFrameRoundedOptionButton"))).setState(true);
            enableVennDiagramPDRoundedFrameControls(getController().getDiagram().isFrameProp());
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
        }
    }
    
    public void enableVennDiagramColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl1"), bool);
            enableControl(xControlContainer.getControl("allDiagramOptionButton"), bool);
            enableControl(xControlContainer.getControl("selectedItemsOptionButton"), bool);
            enableControl(xControlContainer.getControl("baseColorsOptionButton"), bool);
            enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), bool);
            enableControl(xControlContainer.getControl("colorOptionButton"), bool);
            enableControl(xControlContainer.getControl("colorImageControl"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), bool);
            enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), bool);
            if(bool)
                setVennDiagramPDColorControls();
        }
    }
    
    public void setVennDiagramPDColorControls(){
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
        }
    }
    
    public void setVennDiagramPDColorRadioButtons(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            boolean isAllDiagramModify = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).getState();
            if(isAllDiagramModify){
                enableControl(xControlContainer.getControl("baseColorsOptionButton"), true);
                enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), true);
            }else{
                boolean isBaseColorsMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).getState();
                if(isBaseColorsMode){
                    enableControl(xControlContainer.getControl("setBaseColorsCommandButton"), false);
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorImageControl"), true);
                }
                boolean isPreDefinedColorThemeMode = ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).getState();
                if(isPreDefinedColorThemeMode){
                    enableControl(xControlContainer.getControl("preDefinedColorThemesListBox"), false);
                    ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
                    enableControl(xControlContainer.getControl("colorImageControl"), true);
                }
                enableControl(xControlContainer.getControl("baseColorsOptionButton"), false);
                enableControl(xControlContainer.getControl("preDefinedColorThemesOptionButton"), false);
            }
        }
    }
    
    public void enableVennDiagramPDRoundedFrameControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl6"), bool);
            enableControl(xControlContainer.getControl("yesFrameRoundedOptionButton"), bool);
            enableControl(xControlContainer.getControl("noFrameRoundedOptionButton"), bool);
        }
    }
    
    public void setPropertiesOfVennDiagram(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle !=VennDiagram.USER_DEFINE){
                if(sNewStyle == VennDiagram.DEFAULT){                      //selectAllShape, modifyColors,   colorMode           transparency,       outline, frame, rounded frame
                    ((VennDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_TRANSP, true,  Diagram.LINE_WIDTH100, true, true);
                     getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == VennDiagram.WITHOUT_OUTLINE){
                    ((VennDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_TRANSP, false,  Diagram.LINE_WIDTH100, true, true);
                     getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == VennDiagram.WITHOUT_FRAME){
                    ((VennDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_TRANSP, true,  Diagram.LINE_WIDTH100, false, true);
                     getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == VennDiagram.NOT_ROUNDED){
                    ((VennDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, Diagram.MEDIUM_TRANSP, true,  Diagram.LINE_WIDTH100, true, false);
                     getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeStyle(sNewStyle);
                    ((VennDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   colorMode, Diagram.MEDIUM_TRANSP, true,  Diagram.LINE_WIDTH100, true, true);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeColors();
                }
            } else {
                boolean modifyColors = false;
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
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeColors();
                    }
                }else{
                    getController().getDiagram().setModifyColorsProp(false);
                }

                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noTransparencyOptionButton"))).getState())
                    getController().getDiagram().setTransparencyProp(Diagram.NULL_TRANSP);
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("mediumTransparencyOptionButton"))).getState())
                    getController().getDiagram().setTransparencyProp(Diagram.MEDIUM_TRANSP);
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("extraTransparencyOptionButton"))).getState())
                    getController().getDiagram().setTransparencyProp(Diagram.EXTRA_TRANSP);
                
                boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
                getController().getDiagram().setOutlineProp(isOutline);
                if(isOutline){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                    getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
                }
                
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).getState())
                    getController().getDiagram().setFrameProp(true);
                else
                    getController().getDiagram().setFrameProp(false);
                
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameRoundedOptionButton"))).getState())
                    getController().getDiagram().setRoundedFrameProp(true);
                else
                    getController().getDiagram().setRoundedFrameProp(false);
            }
            
            setTextProperties();
        }
    }

    //**************************************************************************************************
    
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setCycleDiagramPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != CycleDiagram.USER_DEFINE){
            
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            
            enableVisibleCycleDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfCycleDiagramPD(true);
            setCycleDiagramPropsControls(false);

            if(sListPos == CycleDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/ring_default.png");
                //                                      area   modifiesColor colorMode       outline   textFrame  shadow
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/ring_withoutOutline.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localNo, localNo, localNo);
            }
            if(sListPos == CycleDiagram.WITH_SHADOW){
                setGraphic(xImageControl, "/images/ring_withShadow.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo, localYes);
            }
            if(sListPos == CycleDiagram.WITH_FRAME){
                setGraphic(xImageControl, "/images/ring_withFrame.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localBaseColors, localYes, localYes, localNo);
            }
            if(sListPos == CycleDiagram.GREEN_DARK){
                setGraphic(xImageControl, "/images/ring_greenDark.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.GREEN_BRIGHT){
                setGraphic(xImageControl, "/images/ring_greenBright.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.BLUE_DARK){
                setGraphic(xImageControl, "/images/ring_blueDark.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.BLUE_BRIGHT){
                setGraphic(xImageControl, "/images/ring_blueBright.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.PURPLE_DARK){
                setGraphic(xImageControl, "/images/ring_purpleDark.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.PURPLE_BRIGHT){
                setGraphic(xImageControl, "/images/ring_purpleBright.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.ORANGE_DARK){
                setGraphic(xImageControl, "/images/ring_orangeDark.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.ORANGE_BRIGHT){
                setGraphic(xImageControl, "/images/ring_orangeBright.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.YELLOW_DARK){
                setGraphic(xImageControl, "/images/ring_yellowDark.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.YELLOW_BRIGHT){
                setGraphic(xImageControl, "/images/ring_yellowBright.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/ring_blueGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/ring_aquaGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.RED_SCHEME){
                setGraphic(xImageControl, "/images/ring_redGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/ring_fireGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.SUN_SCHEME){
                setGraphic(xImageControl, "/images/ring_sunGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/ring_greenGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/ring_oliveGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/ring_purpleGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.PINK_SCHEME){
                setGraphic(xImageControl, "/images/ring_pinkGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/ring_indianGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/ring_maroonGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
            if(sListPos == CycleDiagram.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/ring_brownGradients.png");
                setDescriptionLabelOfCycleDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo, localNo);
            }
        }
        if(sListPos == CycleDiagram.USER_DEFINE){
            enableVisibleControl(xImageControl, false);
            enableVisibleCycleDiagramPropsControls(true);
            enablePropertiesFieldOfCycleDiagramPD(false);
            setCycleDiagramPropsControls(true);
        }
    }

    public void enablePropertiesFieldOfCycleDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }

    public  void setDescriptionLabelOfCycleDiagramPD(String area, String modifiesColor, String colorMode, String outline, String textFrame, String shadow){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(textFrame);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(shadow);
        }
    }

    public void enableVisibleCycleDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
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

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);
        
        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesFrameOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noFrameOptionButton"), bool);
    }
      
    public void setCycleDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        if(bool){
            //color controls
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
            ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
            if(getController().getDiagram().isBaseColorsMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).setState(true);
            if(getController().getDiagram().isSimpleColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
            if(getController().getDiagram().isColorThemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
            if(getController().getDiagram().isColorSchemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).setState(true);
            setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
            short pos = 0;
            if(getController().getDiagram().isColorThemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
            pos = 0;
            if(getController().getDiagram().isColorSchemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).selectItemPos(pos, true);
            enableCycleDiagramColorControls(false);
            
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
            
            //shadow props
            if(getController().getDiagram().isShadowProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noShadowOptionButton"))).setState(true);
                
            //textFrame props
            if(getController().getDiagram().isFrameProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noFrameOptionButton"))).setState(true);
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
        }
    }
    
    public void enableCycleDiagramColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl1"), bool);
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
       
            if(bool)
                setCycleDiagramPDColorControls();
        }
    }
    
    public void setCycleDiagramPDColorControls(){
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
    
    public void setCycleDiagramPDColorRadioButtons(){
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

    public void setPropertiesOfCycleDiagram(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != CycleDiagram.USER_DEFINE){
                if(sNewStyle == CycleDiagram.DEFAULT){                     //selectAllShape, modifyColors,   colorMode       outline, line width,         shadow, frame
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100, false, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == CycleDiagram.WITHOUT_OUTLINE){
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, false, Diagram.LINE_WIDTH000, false, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == CycleDiagram.WITH_FRAME){
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100, false, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == CycleDiagram.WITH_SHADOW){
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100, true, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeStyle(sNewStyle);
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100, false, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeColors();
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((CycleDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100, false, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                }
            } else {
                boolean modifyColors = false;
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
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeColors();
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORSCHEME_MODE_VALUE));
                        //getController().getDiagram().setColorThemeColors();
                        getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                    }
                }else{
                    getController().getDiagram().setModifyColorsProp(false);
                }

                boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
                getController().getDiagram().setOutlineProp(isOutline);
                if(isOutline){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                    getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
                }
                
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).getState())
                    getController().getDiagram().setFrameProp(true);
                else
                    getController().getDiagram().setFrameProp(false);
                
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).getState())
                    getController().getDiagram().setShadowProp(true);
                else
                    getController().getDiagram().setShadowProp(false);
            }
            
            setTextProperties();
        }
    }
    
    //*************************************************************************************************
    
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setPyramidDiagramPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != PyramidDiagram.USER_DEFINE){
            
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localBaseColorsGradients = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColorsGradients.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            
            enableVisiblePyramidDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfPyramidDiagramPD(true);
            setPyramidDiagramPropsControls(false);

            if(sListPos == PyramidDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/pyramid_default.png");
                //                                      area  modifiesColor  colorMode   outline    shadow
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/pyramid_withoutOutline.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localNo, localNo);
            }
            if(sListPos == PyramidDiagram.WITH_SHADOW){
                setGraphic(xImageControl, "/images/pyramid_withShadow.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColors, localYes, localYes);
            }
            if(sListPos == PyramidDiagram.BC_WITH_GRADIENTS){
                setGraphic(xImageControl, "/images/pyramid_basecolorsWithGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localBaseColorsGradients, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.GREEN_DARK){
                setGraphic(xImageControl, "/images/pyramid_greenDark.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.GREEN_BRIGHT){
                setGraphic(xImageControl, "/images/pyramid_greenBright.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.BLUE_DARK){
                setGraphic(xImageControl, "/images/pyramid_blueDark.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.BLUE_BRIGHT){
                setGraphic(xImageControl, "/images/pyramid_blueBright.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.PURPLE_DARK){
                setGraphic(xImageControl, "/images/pyramid_purpleDark.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.PURPLE_BRIGHT){
                setGraphic(xImageControl, "/images/pyramid_purpleBright.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.ORANGE_DARK){
                setGraphic(xImageControl, "/images/pyramid_orangeDark.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.ORANGE_BRIGHT){
                setGraphic(xImageControl, "/images/pyramid_orangeBright.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.YELLOW_DARK){
                setGraphic(xImageControl, "/images/pyramid_yellowDark.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.YELLOW_BRIGHT){
                setGraphic(xImageControl, "/images/pyramid_yellowBright.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorTheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_blueGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_aquaGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.RED_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_redGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_fireGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.SUN_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_sunGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_greenGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_oliveGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_purpleGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.PINK_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_pinkGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_indianGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_maroonGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
            if(sListPos == PyramidDiagram.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/pyramid_brownGradients.png");
                setDescriptionLabelOfPyramidDiagramPD(localAllShape, localYes, localColorScheme, localYes, localNo);
            }
        }
        if(sListPos == PyramidDiagram.USER_DEFINE){
            enableVisibleControl(xImageControl, false);
            enableVisiblePyramidDiagramPropsControls(true);
            enablePropertiesFieldOfPyramidDiagramPD(false);
            setPyramidDiagramPropsControls(true);
        }
    }

    public void enablePropertiesFieldOfPyramidDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 6; i <= 10; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");
               
            }
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 10; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }

    public  void setDescriptionLabelOfPyramidDiagramPD(String area, String modifiesColor, String colorMode, String outline, String shadow){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel6" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(shadow);
        }
    }

    public void enableVisiblePyramidDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
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

        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesShadowOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noShadowOptionButton"), bool);
    }
    
    public void setPyramidDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        if(bool){
            //color controls
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
            ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
            if(getController().getDiagram().isBaseColorsMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).setState(true);
            if(getController().getDiagram().isSimpleColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
            if(getController().getDiagram().isColorThemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
            if(getController().getDiagram().isColorSchemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).setState(true);
            setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
            short pos = 0;
            if(getController().getDiagram().isColorThemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
            pos = 0;
            if(getController().getDiagram().isColorSchemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).selectItemPos(pos, true);
            enablePyramidDiagramColorControls(false);
            
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
            
            //shadow props
            if(getController().getDiagram().isShadowProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noShadowOptionButton"))).setState(true);        
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
        }
    }
    
    public void enablePyramidDiagramColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl1"), bool);
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
       
            if(bool)
                setPyramidDiagramPDColorControls();
        }
    }
    
    public void setPyramidDiagramPDColorControls(){
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
    
    public void setPyramidDiagramPDColorRadioButtons(){
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
    
    public void setPropertiesOfPyramidDiagram(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);
            if(sNewStyle != PyramidDiagram.USER_DEFINE){
                if(sNewStyle == PyramidDiagram.DEFAULT){                     //selectAllShape, modifyColors,   colorMode       outline, line width,           shadow
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == PyramidDiagram.WITHOUT_OUTLINE){
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, false, Diagram.LINE_WIDTH000, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == PyramidDiagram.WITH_SHADOW){
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, true, Diagram.LINE_WIDTH100, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == PyramidDiagram.BC_WITH_GRADIENTS){
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_WITH_GRADIENT_MODE, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeStyle(sNewStyle);
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeColors();
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((PyramidDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                }
            } else {
                boolean modifyColors = false;
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
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeColors();
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORSCHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeGradientColors();
                    }
                }else{
                    getController().getDiagram().setModifyColorsProp(false);
                }

                boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
                getController().getDiagram().setOutlineProp(isOutline);
                if(isOutline){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                    getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
                }
     
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesShadowOptionButton"))).getState())
                    getController().getDiagram().setShadowProp(true);
                else
                    getController().getDiagram().setShadowProp(false);
            }
            
            setTextProperties();
        }
    }
    
    //*************************************************************************************************
    
    //runing after user push Properties button on ControlDialog and when user change style in styleListBox in PropsDialog
    public void setTargetDiagramPropsDialog(short sListPos){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        XControl xImageControl = xControlContainer.getControl("imageControl");
        
        if(sListPos != TargetDiagram.USER_DEFINE){
            
            String localAllShape = getDialogPropertyValue("Strings2", "Strings2.Common.LocalDiagramScope.Label");
            String localYes = getDialogPropertyValue("Strings2", "Strings2.Common.LocalYes.Label");
            String localNo = getDialogPropertyValue("Strings2", "Strings2.Common.LocalNo.Label");
            String localBaseColors = getDialogPropertyValue("Strings2", "Strings2.Common.LocalBaseColors.Label");
            String localColorScheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorScheme.Label");
            String localColorTheme = getDialogPropertyValue("Strings2", "Strings2.Common.LocalColorTheme.Label");
            String localCenter = getDialogPropertyValue("Strings2", "Strings2.Common.LocalCentert.Label");
            String localLeft = getDialogPropertyValue("Strings2", "Strings2.Common.LocalLeft.Label");
        
            enableVisibleTargetDiagramPropsControls(false);
            enableVisibleControl(xImageControl, true);
            enablePropertiesFieldOfTargetDiagramPD(true);
            setTargetDiagramPropsControls(false);

            setTextImageControlColor(Diagram.DEFAULT_TEXT_COLOR);
            
            if(sListPos == TargetDiagram.DEFAULT){
                setGraphic(xImageControl, "/images/target_default.png");
                //                                      area   modifiesColor  colorMode        layout       outline   frame
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.WITHOUT_OUTLINE){
                setGraphic(xImageControl, "/images/target_withoutOutline.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localNo, localNo);
            }
            if(sListPos == TargetDiagram.WITH_FRAME){
                setGraphic(xImageControl, "/images/target_withTextFrame.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localCenter, localYes, localYes);
            }
            if(sListPos == TargetDiagram.LEFT_LAYOUT){
                setGraphic(xImageControl, "/images/target_leftLayout.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localBaseColors, localLeft, localYes, localNo);
            }
            if(sListPos == TargetDiagram.GREEN_DARK){
                setGraphic(xImageControl, "/images/target_greenDark.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOGREENS[1]);
            }
            if(sListPos == TargetDiagram.GREEN_BRIGHT){
                setGraphic(xImageControl, "/images/target_greenBright.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOGREENS[3]);
            }
            if(sListPos == TargetDiagram.BLUE_DARK){
                setGraphic(xImageControl, "/images/target_blueDark.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOBLUES[1]);
            }
            if(sListPos == TargetDiagram.BLUE_BRIGHT){
                setGraphic(xImageControl, "/images/target_blueBright.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOBLUES[3]);
            }
            if(sListPos == TargetDiagram.PURPLE_DARK){
                setGraphic(xImageControl, "/images/target_purpleDark.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOPURPLES[1]);
            }
            if(sListPos == TargetDiagram.PURPLE_BRIGHT){
                setGraphic(xImageControl, "/images/target_purpleBright.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOPURPLES[3]);
            }
            if(sListPos == TargetDiagram.ORANGE_DARK){
                setGraphic(xImageControl, "/images/target_orangeDark.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOORANGES[1]);
            }
            if(sListPos == TargetDiagram.ORANGE_BRIGHT){
                setGraphic(xImageControl, "/images/target_orangeBright.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOORANGES[3]);
            }
            if(sListPos == TargetDiagram.YELLOW_DARK){
                setGraphic(xImageControl, "/images/target_yellowDark.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOYELLOWS[1]);
            }
            if(sListPos == TargetDiagram.YELLOW_BRIGHT){
                setGraphic(xImageControl, "/images/target_yellowBright.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorTheme, localCenter, localYes, localNo);
                setTextImageControlColor(Diagram.aLOYELLOWS[3]);
            }
            if(sListPos == TargetDiagram.BLUE_SCHEME){
                setGraphic(xImageControl, "/images/target_blueGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.AQUA_SCHEME){
                setGraphic(xImageControl, "/images/target_aquaGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.RED_SCHEME){
                setGraphic(xImageControl, "/images/target_redGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.FIRE_SCHEME){
                setGraphic(xImageControl, "/images/target_fireGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.SUN_SCHEME){
                setGraphic(xImageControl, "/images/target_sunGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.GREEN_SCHEME){
                setGraphic(xImageControl, "/images/target_greenGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.OLIVE_SCHEME){
                setGraphic(xImageControl, "/images/target_oliveGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.PURPLE_SCHEME){
                setGraphic(xImageControl, "/images/target_purpleGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.PINK_SCHEME){
                setGraphic(xImageControl, "/images/target_pinkGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.INDIAN_SCHEME){
                setGraphic(xImageControl, "/images/target_indianGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.MAROON_SCHEME){
                setGraphic(xImageControl, "/images/target_maroonGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }
            if(sListPos == TargetDiagram.BROWN_SCHEME){
                setGraphic(xImageControl, "/images/target_brownGradients.png");
                setDescriptionLabelOfTargetDiagramPD(localAllShape, localYes, localColorScheme, localCenter, localYes, localNo);
            }     
        } else {
            enableVisibleControl(xImageControl, false);
            enableVisibleTargetDiagramPropsControls(true);
            enablePropertiesFieldOfTargetDiagramPD(false);
            setTargetDiagramPropsControls(true);
        }
    }
    
    public void setTextImageControlColor(int color){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog); 
            short sState = ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).getState();
            if(sState == 0)
                setImageColorOfControl(xControlContainer.getControl("textColorImageControl"), color);
        }
    }

    public void enablePropertiesFieldOfTargetDiagramPD(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            if(bool == false){
                for(int i = 7; i <= 12; i++)
                    ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel" + i))).setText("");

            }
            enableControl(xControlContainer.getControl("FrameControl"), bool);
            for(int i = 1; i <= 12; i++)
                enableControl(xControlContainer.getControl("propsLabel" + i), bool);
        }
    }

    public  void setDescriptionLabelOfTargetDiagramPD(String area, String modifiesColor, String colorMode, String layout, String outline, String tFrame){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel7" ))).setText(area);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel8" ))).setText(modifiesColor);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel9" ))).setText(colorMode);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel10" ))).setText(layout);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel11" ))).setText(outline);
            ((XFixedText)UnoRuntime.queryInterface(XFixedText.class, xControlContainer.getControl("propsLabel12" ))).setText(tFrame);
        }
    }
    
    public void enableVisibleTargetDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);

        enableVisibleControl(xControlContainer.getControl("frameControl2"), bool);
        enableVisibleControl(xControlContainer.getControl("modifyColorsCheckBox"), bool);
        enableVisibleControl(xControlContainer.getControl("frameControl1"), bool);
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

        enableVisibleControl(xControlContainer.getControl("frameControl3"), bool);
        enableVisibleControl(xControlContainer.getControl("centerLayoutOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("leftLayoutOptionButton"), bool);
        
        enableVisibleControl(xControlContainer.getControl("frameControl4"), bool);
        enableVisibleControl(xControlContainer.getControl("yesOutlineOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthLabel"), bool);
        enableVisibleControl(xControlContainer.getControl("lineWidthListBox"), bool);
        enableVisibleControl(xControlContainer.getControl("noOutlineOptionButton"), bool);

        enableVisibleControl(xControlContainer.getControl("frameControl5"), bool);
        enableVisibleControl(xControlContainer.getControl("yesFrameOptionButton"), bool);
        enableVisibleControl(xControlContainer.getControl("noFrameOptionButton"), bool);
    }

    public void setTargetDiagramPropsControls(boolean bool){
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
        if(bool){
            //color controls
            ((XCheckBox)UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyColorsCheckBox"))).setState((short)0);
            ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("allDiagramOptionButton"))).setState(true);
            if(getController().getDiagram().isBaseColorsMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("baseColorsOptionButton"))).setState(true);
            if(getController().getDiagram().isSimpleColorMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("colorOptionButton"))).setState(true);
            if(getController().getDiagram().isColorThemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorThemesOptionButton"))).setState(true);
            if(getController().getDiagram().isColorSchemeMode())
                ((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).setState(true);
            setImageColorOfControl(xControlContainer.getControl("colorImageControl"), getController().getDiagram().getColorProp());
            short pos = 0;
            if(getController().getDiagram().isColorThemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORTHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorThemesListBox"))).selectItemPos(pos, true);
            pos = 0;
            if(getController().getDiagram().isColorSchemeMode())
                pos = (short)(getController().getDiagram().getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE);
            ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).selectItemPos(pos, true);
            enableTargetDiagramColorControls(false);
            
            //layout controls
            if(((TargetDiagram)getController().getDiagram()).isLeftLayoutProperty())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("leftLayoutOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("centerLayoutOptionButton"))).setState(true);

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
 
            //textFrame props
            if(getController().getDiagram().isFrameProp())
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).setState(true);
            else
                ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("noFrameOptionButton"))).setState(true);
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
            if(getController().getDiagram().isTextColorChange())
                ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).setState((short)1);
            else
                ((XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, xControlContainer.getControl("modifyTextColorCheckBox"))).setState((short)0);
            enableTextColorImageControl(getController().getDiagram().isTextColorChange());
        }    
    }
    
    public void enableTargetDiagramColorControls(boolean bool){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            enableControl(xControlContainer.getControl("frameControl1"), bool);
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
       
            if(bool)
                setTargetDiagramPDColorControls();
        }
    }
    
    public void setTargetDiagramPDColorControls(){
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
    
    public void setTargetDiagramPDColorRadioButtons(){
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

    public void setPropertiesOfTargetDiagram(){
        if(m_xPropsDialog != null){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            short sNewStyle = styleListBox.getSelectedItemPos();
            getController().getDiagram().setStyleProp(sNewStyle);

            if(sNewStyle != TargetDiagram.USER_DEFINE){
                if(sNewStyle == CycleDiagram.DEFAULT){                     //selectAllShape, modifyColors,   colorMode,    isLeftLayout, outline, line width,        frame
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true,   true,   Diagram.BASE_COLORS_MODE, false,     true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == TargetDiagram.WITHOUT_OUTLINE){
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, false, false, Diagram.LINE_WIDTH000, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == TargetDiagram.WITH_FRAME){
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, false, true, Diagram.LINE_WIDTH100, true);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(sNewStyle == TargetDiagram.LEFT_LAYOUT){
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true, true, Diagram.BASE_COLORS_MODE, true, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setLineColorProp(getController().getDiagram().getDefaultLineColor());
                }
                if(getController().getDiagram().isColorThemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfThemeStyle(sNewStyle);
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, false, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeColors();
                }
                if(getController().getDiagram().isColorSchemeStyle(sNewStyle)){
                    short colorMode = getController().getDiagram().getColorModeOfSchemeStyle(sNewStyle);
                    ((TargetDiagram)getController().getDiagram()).setPropertiesValues(true, true, colorMode, false, true, Diagram.LINE_WIDTH100, false);
                    getController().getDiagram().setColorModeProp(colorMode);
                    getController().getDiagram().setColorThemeGradientColors();
                }
            } else {
                boolean modifyColors = false;
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
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORTHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeColors();
                    }
                    if(((XRadioButton)UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("preDefinedColorSchemesOptionButton"))).getState()){
                        short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("preDefinedColorSchemesListBox"))).getSelectedItemPos();
                        getController().getDiagram().setColorModeProp((short)(selectedItemPos + Diagram.FIRST_COLORSCHEME_MODE_VALUE));
                        getController().getDiagram().setColorThemeGradientColors();
                    }
                }else{
                    getController().getDiagram().setModifyColorsProp(false);
                }

                boolean isLeftLayout = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("leftLayoutOptionButton"))).getState();
                ((TargetDiagram)getController().getDiagram()).setLeftLayoutProperty(isLeftLayout);
                
                boolean isOutline = ((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesOutlineOptionButton"))).getState();
                getController().getDiagram().setOutlineProp(isOutline);
                if(isOutline){
                    short selectedItemPos = ((XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("lineWidthListBox"))).getSelectedItemPos();
                    getController().getDiagram().setShapesLineWidthProp(getController().getDiagram().getLineWidthValue(selectedItemPos));
                }
                
                if(((XRadioButton) UnoRuntime.queryInterface(XRadioButton.class, xControlContainer.getControl("yesFrameOptionButton"))).getState())
                    getController().getDiagram().setFrameProp(true);
                else
                    getController().getDiagram().setFrameProp(false);
            }
            
            setTextProperties();
        }
    }
/*
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
*/
    //********************************************************************************************************


    
}
