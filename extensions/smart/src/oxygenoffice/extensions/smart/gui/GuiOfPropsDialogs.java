package oxygenoffice.extensions.smart.gui;


import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XWindow;
import com.sun.star.frame.XFrame;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;


public class GuiOfPropsDialogs extends  GuiOfProcessgroupPropsDialogs {

    public GuiOfPropsDialogs(){ }
    
    public GuiOfPropsDialogs(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
    }
    
    public short executePropertiesDialog() {
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
                if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                    diagramDefine = "ContinuousBlockProcessPropsDialog.xdl";
                if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                    diagramDefine = "BendingProcessPropsDialog.xdl";
                if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                    diagramDefine = "StaggeredProcessPropsDialog.xdl";
                if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                    diagramDefine = "UpwardArrowProcessPropsDialog.xdl";
                String sDialogURL = sPackageURL + "/dialogs/" + diagramDefine;
                if(getController().getGroupType() == Controller.ORGANIGROUP)
                    m_oListenerOfPropsDialogs = new ListenerOfOrganigroupPropsDialogs(this, m_Controller);
                if(getController().getGroupType() == Controller.RELATIONGROUP)
                    m_oListenerOfPropsDialogs = new ListenerOfRelationgruopPropsDialogs(this, m_Controller);
                if(getController().getGroupType() == Controller.PROCESSGROUP)
                    m_oListenerOfPropsDialogs = new ListenerOfProcessgroupPropsDialogs(this, m_Controller);
                m_xPropsDialog = xDialogProv.createDialogWithHandler(sDialogURL, m_oListenerOfPropsDialogs);

                if (m_xPropsDialog != null) {
                    //can set focus with XWindow interface in GuiOfDialogs.java
                    m_xPropsDialogWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xPropsDialog);
                    XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
                    setControlScaleImageProp(xControlContainer.getControl("imageControl"), false);
                }
            } catch (com.sun.star.lang.IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        if(m_xPropsDialog != null){
            short style = getController().getDiagram().getStyleProp();
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xPropsDialog);
            XListBox styleListBox = (XListBox)UnoRuntime.queryInterface(XListBox.class, xControlContainer.getControl("listBox"));
            styleListBox.selectItemPos(style, true);
            
            if(getController().getGroupType() == Controller.ORGANIGROUP)
                setOrganigramPropsDialog(style);
            if (getController().getDiagramType() == Controller.VENNDIAGRAM)
                setVennDiagramPropsDialog(style);
            if(getController().getDiagramType() == Controller.CYCLEDIAGRAM)
                setCycleDiagramPropsDialog(style);
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM)
                setPyramidDiagramPropsDialog(style);
            if(getController().getDiagramType() == Controller.TARGETDIAGRAM)
                setTargetDiagramPropsDialog(style);
            if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                setContinuousBlockProcessPropsDialog(style);
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS)
                setStaggeredProcessPropsDialog(style);
            if(getController().getDiagramType() == Controller.BENDINGPROCESS)
                setBendingProcessPropsDialog(style);
            if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                setUpwardArrowProcessPropsDialog(style);
            setShownPropsDialog(true);
            short exec = m_xPropsDialog.execute();
            setShownPropsDialog(false);
            return exec;
        }
        return -1;
    }

}
