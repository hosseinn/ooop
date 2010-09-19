package org.openoffice.extensions.diagrams;

import com.sun.star.frame.XFrame;

// store pairs of XFrame and Controller
public class FrameObject {

    private XFrame      m_xFrame        = null;
    private Controller  m_Controller    = null;

    FrameObject(XFrame xFrame, Controller controller){
        m_xFrame = xFrame;
        m_Controller = controller;
    }

    public XFrame getXFrame(){
        return m_xFrame;
    }

    public void setXFrame(XFrame xFrame){
        m_xFrame = xFrame ;
    }

    public Controller getController(){
        return m_Controller;
    }

    public void setController(Controller controller){
        m_Controller = controller ;
    }

}
