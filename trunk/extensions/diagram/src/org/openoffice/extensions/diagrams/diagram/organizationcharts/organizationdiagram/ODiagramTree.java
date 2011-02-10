package org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class ODiagramTree extends DiagramTree{

    public static short LASTHORLEVEL = 2;

    public ODiagramTree(OrganizationDiagram organigram){
        super(organigram);
    }

    public ODiagramTree(OrganizationDiagram organigram, DiagramTree diagramTree) {
        super(organigram, diagramTree);
        OTreeItem.initStaticMembers();
        m_RootItem = new OTreeItem(this, null, diagramTree.getRootItem());
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.convertTreeItems(diagramTree.getRootItem());
    }

    @Override
    public void initTreeItems(){
        OTreeItem.initStaticMembers();
        m_RootItem = new OTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public XShape getFirstChildShape(XShape xDadShape){
        // the struct of diagram change below second level that's why we need level of shape
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                int endOfConn = getEndGluePointIndex(xConnShape);
                if(ODiagramTree.LASTHORLEVEL == -1 && endOfConn == 3)
                    ODiagramTree.LASTHORLEVEL = (short)(level - 1);
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= ODiagramTree.LASTHORLEVEL){
                    if( xPos == -1 || xChildeShape.getPosition().X < xPos){
                        xPos = xChildeShape.getPosition().X;
                        xFirstChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y < yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xFirstChildShape = xChildeShape;
                    }
                }
            }
        }
        return xFirstChildShape;
    }

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= ODiagramTree.LASTHORLEVEL){
                    if( xPos == -1 || xChildeShape.getPosition().X > xPos){
                        xPos = xChildeShape.getPosition().X;
                        xLastChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y > yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xLastChildShape = xChildeShape;
                    }

                }
            }
        }
        return xLastChildShape;
    }

    @Override
    public XShape getFirstSiblingShape(XShape xBaseShape, TreeItem dad){
        if(dad == null)
            return null;
        if(dad.getRectangleShape() == null)
            return null;
        int    level                = dad.getLevel() + 1;
        XShape xDadShape            = dad.getRectangleShape();
        XShape xSiblingShape        = null;
        XShape xFirstSiblingShape   = null;
        Point  baseShapePos         = xBaseShape.getPosition();
        int    xPos                 = -1;
        int    yPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if(level <= ODiagramTree.LASTHORLEVEL){
                    if( xSiblingShape.getPosition().X > baseShapePos.X){
                        if( xPos == -1 || xSiblingShape.getPosition().X < xPos){
                            xPos = xSiblingShape.getPosition().X;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }else{
                    if( xSiblingShape.getPosition().Y > baseShapePos.Y){
                        if( yPos == -1 || xSiblingShape.getPosition().Y < yPos){
                            yPos = xSiblingShape.getPosition().Y;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }
            }
        }
        return xFirstSiblingShape;
    }

    @Override
    public void refresh(){
        OTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
        m_RootItem.display();
    }

    @Override
    public void refreshConnectorProps(){
        for(XShape xConnShape : connectorList){
            XShape xShape = getEndShapeOfConnector(xConnShape);
            short level = getTreeItem(xShape).getLevel();
            Integer start, end;
            start = new Integer(2);
            if(level <= ODiagramTree.LASTHORLEVEL){
                end = new Integer(0);
            } else{
                end = new Integer(3);
            }
            getOrgChart().setConnectorShapeProps(xConnShape, start, end);
        }
    }

    public int getEndGluePointIndex(XShape xConnShape){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
                return AnyConverter.toInt(xProps.getPropertyValue("EndGluePointIndex"));
        } catch (UnknownPropertyException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.out.println(ex.getLocalizedMessage());
        }
        return -1;
    }
}