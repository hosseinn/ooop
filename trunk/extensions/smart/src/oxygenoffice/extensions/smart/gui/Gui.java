package oxygenoffice.extensions.smart.gui;

import com.sun.star.frame.XFrame;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;

// Gui class is an extended class. Had to split because of its huge code
// class hierarchy :
// Gui <- GuiOfPropsDialogs <- GuiOfProcessgroupPropsDialogs <- GuiOfRelationgroupPropsDialogs
// <- GuiOfOrganigroupPropsDialogs <- GuiOfDialogs <- GuiBaseClass  // <- = extends

public class Gui extends GuiOfPropsDialogs {

    public Gui() { }

    public Gui(Controller controller, XComponentContext xContext, XFrame xFrame){
        super(controller, xContext, xFrame);
    }

}
