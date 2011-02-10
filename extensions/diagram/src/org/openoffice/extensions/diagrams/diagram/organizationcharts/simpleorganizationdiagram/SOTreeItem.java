package org.openoffice.extensions.diagrams.diagram.organizationcharts.simpleorganizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class SOTreeItem extends TreeItem{

    protected static double[]        _maxPositions;
    protected static double          _maxPos                = -1.0;

    private static int              _horSpace               = 0;
    private static int              _verSpace               = 0;
    private static int              _shapeWidth             = 0;
    private static int              _shapeHeight            = 0;
    private static int              _groupPosX              = 0;
    private static int              _groupPosY              = 0;


    public SOTreeItem(DiagramTree diagramTree, TreeItem dad, TreeItem item) {
        super(diagramTree, dad, item);
    }

    public SOTreeItem(DiagramTree diagramTree, XShape xShape, TreeItem dad, short level, double num){
        super(diagramTree, xShape, dad);
        setLevel(level);
        setPos(num);
    }

    public static void initStaticMembers(){
        _maxLevel = -1;
        _maxPos = -1.0;
        _maxPositions = new double[100];
        for(int i=0; i<_maxPositions.length; i++)
            _maxPositions[i] = -1.0;
    }

    @Override
    public final void setPos(double pos){
        m_Pos = pos;
        if(m_Pos > _maxPositions[m_Level])
            _maxPositions[m_Level] = m_Pos;
        if(m_Pos > _maxPos)
            _maxPos = m_Pos;
        //XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
        //xText.setString(m_Level +":"+m_Pos);
    }

    @Override
    public void initTreeItems(){
        XShape xFirstChildShape = getDiagramTree().getFirstChildShape(m_xRectangleShape);
        if(xFirstChildShape != null){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = _maxPositions[firstChildLevel] + 1.0;
            m_FirstChild = new SOTreeItem(getDiagramTree(), xFirstChildShape, this, firstChildLevel , firstChildPos);
            m_FirstChild.initTreeItems();
        }

        XShape xFirstSiblingShape = getDiagramTree().getFirstSiblingShape(m_xRectangleShape, m_Dad);
        if(xFirstSiblingShape != null){
            short firstSiblingLevel = m_Level;
            double firstSiblingNum = m_Pos + 1.0;
            m_FirstSibling = new SOTreeItem(getDiagramTree(), xFirstSiblingShape, m_Dad, firstSiblingLevel , firstSiblingNum);
            m_FirstSibling.initTreeItems();
        }

        if(m_Dad != null && m_Dad.getFirstChild().equals(this)){
            double newPos = 0.0;
            if(isFirstSibling())
                newPos = (_maxPositions[m_Level] + m_Pos) / 2;
            else
                newPos = m_Pos;
            if(newPos > m_Dad.getPos())
                m_Dad.setPos(newPos);
            if(newPos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - newPos);
        }
    }

    @Override
    public void setPositionsOfItems(){
        if(isFirstChild()){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = _maxPositions[firstChildLevel] + 1.0;
            m_FirstChild.setLevel(firstChildLevel);
            m_FirstChild.setPos(firstChildPos);

            m_FirstChild.setPositionsOfItems();
        }

        if(isFirstSibling()){
            short firstSiblingLevel = m_Level;
            double firstSiblingNum = m_Pos + 1.0;
            m_FirstSibling.setLevel(firstSiblingLevel);
            m_FirstSibling.setPos(firstSiblingNum);

            m_FirstSibling.setPositionsOfItems();
        }

        if(m_Dad != null && m_Dad.getFirstChild().equals(this)){
            double newPos = 0.0;
            if(isFirstSibling())
                newPos = (_maxPositions[m_Level] + m_Pos) / 2;
            else
                newPos = m_Pos;
            if(newPos > m_Dad.getPos())
                m_Dad.setPos(newPos);
            if(newPos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - newPos);

        }
    }

    // set _shapeWidth, , horSpace, _shapeHeight, verSpace, _groupPosX, _groupPosY
    @Override
    public void setProps(){
        int baseShapeWidth = _shapeWidth = getDiagramTree().getControlShapeSize().Width;
        int baseShapeHeight = _shapeHeight = getDiagramTree().getControlShapeSize().Height;
        _horSpace = _verSpace = 0;
        if(_maxPos > 0){
            int horUnit = (int)(baseShapeWidth / ( _maxPos * (getDiagramTree().getOrgChart().getWIDTH() + getDiagramTree().getOrgChart().getHORSPACE()) + getDiagramTree().getOrgChart().getWIDTH()));
            _shapeWidth = horUnit * getDiagramTree().getOrgChart().getWIDTH();
            _horSpace = horUnit * getDiagramTree().getOrgChart().getHORSPACE();
        }
        if(_maxLevel > 0){
            int verUnit = (baseShapeHeight) / ( _maxLevel * (getDiagramTree().getOrgChart().getHEIGHT() + getDiagramTree().getOrgChart().getVERSPACE()) + getDiagramTree().getOrgChart().getHEIGHT());
            _shapeHeight = verUnit * getDiagramTree().getOrgChart().getHEIGHT();
            _verSpace = verUnit * getDiagramTree().getOrgChart().getVERSPACE();
        }
        _groupPosX = getDiagramTree().getControlShapePos().X;
        _groupPosY = getDiagramTree().getControlShapePos().Y;
    }

    @Override
    public void setPosOfRect(){
        int xCoord = _groupPosX + (int)((_shapeWidth + _horSpace) * getPos());
        int yCoord = _groupPosY + (_shapeHeight + _verSpace) * getLevel();
        setPosition(new Point(xCoord, yCoord));
        setSize(new Size(_shapeWidth, _shapeHeight));
    }

    @Override
    public void convertTreeItems(TreeItem treeItem){
        if(treeItem.isFirstChild()){
            m_FirstChild = new SOTreeItem(getDiagramTree(), this, treeItem.getFirstChild());
            m_FirstChild.convertTreeItems(treeItem.getFirstChild());
        }
        if(treeItem.isFirstSibling()){
            m_FirstSibling = new SOTreeItem(getDiagramTree(), getDad(), treeItem.getFirstSibling());
            m_FirstSibling.convertTreeItems(treeItem.getFirstSibling());
        }
    }
}
