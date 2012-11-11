package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.lang.WrappedTargetException;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;

public class ListenerOfPropsDialogs implements  XDialogEventHandler {

    private GuiOfPropsDialogs   m_GuiOfPropsDialog  = null;
    private Controller          m_Controller        = null;
    
    ListenerOfPropsDialogs(GuiOfPropsDialogs gui, Controller controller){
        m_GuiOfPropsDialog = gui;
        m_Controller = controller;
    }

    public GuiOfPropsDialogs getGui(){
        return m_GuiOfPropsDialog;
    }

    public Controller getController(){
        return m_Controller;
    }

    @Override
    public boolean callHandlerMethod(XDialog arg0, Object arg1, String arg2) throws WrappedTargetException {
        return false;
    }

    @Override
    public String[] getSupportedMethodNames() {
        return null;
    }
 
    public void setNextColorOnControlDialog(){
        if(getController().getDiagram().isBaseColorsMode()){
            RelationDiagram diagram = ((RelationDiagram)getController().getDiagram());
            diagram.setColorProp(diagram.getNextColor());
        }                
    }
}
