package org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class OTreeItem extends TreeItem{


    protected static double          _level1MaxPos           = -1;
    protected static double          _level2MaxPos           = -1;
    protected static double          _maxPos                 = -1;

    protected static int            _borderTop              = 0;
    protected static int            _borderLeft             = 0;
    protected static int            _groupSizeWidth         = 0;
    protected static int            _groupSizeHeight        = 0;
    protected static int            _horSpace               = 0;
    protected static int            _verSpace               = 0;
    protected static int            _shapeWidth             = 0;
    protected static int            _shapeHeight            = 0;
    protected static int            _horUnit                = 0;

    
    public OTreeItem(DiagramTree diagramTree, XShape xShape, TreeItem dad, short level, double num){
        super(diagramTree, xShape, dad);
        setLevel(level);
        setPos(num);
    }

    public static void initStaticMembers(){
        TreeItem._maxLevel = -1;
        _maxPos = _level1MaxPos = _level2MaxPos = -1.0;
    }

    @Override
    public final void setPos(double pos){
        m_Pos = pos;
        //XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
        //xText.setString(m_Level +":"+m_Pos);
        if(m_Level == 1)
           if(m_Pos >_level1MaxPos)
                _level1MaxPos = m_Pos;
        if(m_Level == 2)
            if(m_Pos >_level2MaxPos)
                _level2MaxPos = m_Pos;
        if(_maxPos < m_Pos)
            _maxPos = m_Pos;
        if(m_Level > 3)
            if(_level2MaxPos < (m_Pos - 1))
                _level2MaxPos = (short)(m_Pos - 1);
    }

    // set the treeItems in the tree (set dad-firstChild-firstSibling relations of TreeItems)
    // and set m_level and m_PosNum
    @Override
    public void initTreeItems(){
        XShape xFirstChildShape = getDiagramTree().getFirstChildShape(m_xRectangleShape);
        if(xFirstChildShape != null){
            short firstChildLevel = (short)(m_Level + 1);
            short firstChildNum = (short)0;
            if(firstChildLevel == 1)
                if(_level1MaxPos > -1)
                    firstChildNum = (short)(_level1MaxPos + 2);
            if(firstChildLevel == 2)
                if(_level2MaxPos > -1)
                    firstChildNum = (short)(_level2MaxPos + 2);
            if(firstChildLevel == 3)
                firstChildNum = (short)(_level2MaxPos + 1);
            if(firstChildLevel > 3)
                firstChildNum = (short)(getPos() + 1);
                //firstChildNum = (short)(++_level2MaxPos + 1);
                
            m_FirstChild = new OTreeItem(getDiagramTree(), xFirstChildShape, this, firstChildLevel , firstChildNum);
            m_FirstChild.initTreeItems();
        }

        if(m_FirstChild != null){
            TreeItem lastChildItem = getDiagramTree().getTreeItem(getDiagramTree().getLastChildShape(m_xRectangleShape));
            short newPos = (short)(  ( m_FirstChild.getPos() + lastChildItem.getPos() ) / 2 );
            if(m_Level == 0)
                setPos( newPos );
            if(m_Level == 1) {
                if(newPos > getPos()){
                    setPos( newPos );
                }else if(newPos < getPos()){
                    short diff = (short)(getPos() - newPos);
                    increaseDescendantsPosNum(diff);
                }
            }
        }

        XShape xFirstSiblingShape = getDiagramTree().getFirstSiblingShape(m_xRectangleShape, m_Dad);
        if(xFirstSiblingShape != null){
            short firstSiblingLevel = m_Level;
            double firstSiblingNum = 0.0;
            if(m_Level == 1 || m_Level == 2)
                firstSiblingNum = (short)(getPos() + 2);
            if(m_Level == 2){
                int deep = getDeepOfTreeBranch(this);
                if(deep > 1)
                    firstSiblingNum = (short)(getPos() + deep + 1);
            }
            if(m_Level > 2) {
                firstSiblingNum = getPos();
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }

            m_FirstSibling = new OTreeItem(getDiagramTree(), xFirstSiblingShape, m_Dad, firstSiblingLevel , firstSiblingNum);
            m_FirstSibling.initTreeItems();
        }
    }

    @Override
    public void setPositionsOfItems(){
        if(m_FirstChild != null){
            short firstChildLevel = (short)(m_Level + 1);
            short firstChildNum = (short)0;
            if(firstChildLevel == 1)
                if(_level1MaxPos > -1)
                    firstChildNum = (short)(_level1MaxPos + 2);
            if(firstChildLevel == 2)
                if(_level2MaxPos > -1)
                    firstChildNum = (short)(_level2MaxPos + 2);
            if(firstChildLevel == 3)
                firstChildNum = (short)(_level2MaxPos + 1);
            if(firstChildLevel > 3)
                firstChildNum = (short)( getPos() + 1);
           
            m_FirstChild.setLevel(firstChildLevel);
            m_FirstChild.setPos(firstChildNum);
            m_FirstChild.setPositionsOfItems();

            TreeItem lastChildItem = getDiagramTree().getTreeItem(getDiagramTree().getLastChildShape(m_xRectangleShape));
            short newPos = (short)(  ( m_FirstChild.getPos() + lastChildItem.getPos() ) / 2 );
            if(m_Level == 0)
                setPos( newPos );
            if(m_Level == 1) {
                if(newPos > getPos()){
                    setPos( newPos );
                }else if(newPos < getPos()){
                    short diff = (short)(getPos() - newPos);
                    increaseDescendantsPosNum(diff);
                }
            }
        }

        if(m_FirstSibling != null){
 
            short firstSiblingLevel = m_Level;
            short firstSiblingNum = (short)0;
            if(m_Level == 1 || m_Level == 2)
                firstSiblingNum = (short)(getPos() + 2);
            if(m_Level == 2){
                int deep = getDeepOfTreeBranch(this);
                if(deep > 1){
                    firstSiblingNum = (short)(getPos() + deep + 1);
                }
            }
            if(m_Level > 2) {
                firstSiblingNum = (short)getPos();
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }

            m_FirstSibling.setLevel(firstSiblingLevel);
            m_FirstSibling.setPos(firstSiblingNum);

            m_FirstSibling.setPositionsOfItems();
        }
    }

    // set _horSpace, _shapeWidth, _horUnit, _verSpace, _shapeHeight and controlShape size
    @Override
    public void setProps(){
        Size size = getDiagramTree().getControlShapeSize();
        Point point = getDiagramTree().getControlShapePos();
        _groupSizeWidth = size.Width;
        _groupSizeHeight = size.Height;
        _borderLeft = point.X;
        _borderTop =  point.Y;

        if(_maxPos > 0){
            _horSpace =  (short)(_groupSizeWidth * 2 * getDiagramTree().getOrgChart().getHORSPACE() / ( _maxPos * (getDiagramTree().getOrgChart().getWIDTH() + getDiagramTree().getOrgChart().getHORSPACE()) + 2 * getDiagramTree().getOrgChart().getWIDTH() ) );
            _shapeWidth = _horSpace * getDiagramTree().getOrgChart().getWIDTH() / getDiagramTree().getOrgChart().getHORSPACE();
        }else{
            _horSpace = 0;
            _shapeWidth = _groupSizeWidth;
        }

        _horUnit = (_shapeWidth + _horSpace ) / 2;

        int verBaseUnit = _groupSizeHeight / ( (_maxLevel+1)*getDiagramTree().getOrgChart().getHEIGHT() + _maxLevel*getDiagramTree().getOrgChart().getVERSPACE() );
        if(_maxLevel > 0)
            _verSpace = verBaseUnit * getDiagramTree().getOrgChart().getVERSPACE();
        else
            _verSpace = 0;

        _shapeHeight = verBaseUnit * getDiagramTree().getOrgChart().getHEIGHT();
    }

    @Override
    public void setPosOfRect(){
        int shapeWidth  = _shapeWidth;
        int shapeHeight = _shapeHeight;
        int xCoord;
        if(getLevel() == 0){
            double dFirstChildPos = (double)getFirstChild().getPos();
            double dLevel1MaxPos = (double)_level1MaxPos;
            double dPos = dFirstChildPos + ( (dLevel1MaxPos - dFirstChildPos) / 2 );
            //System.out.println(dPos);
            xCoord = (int)(_borderLeft + dPos * _horUnit);
        }
        else
            xCoord = (short)(_borderLeft + getPos() * _horUnit);//(shapeWidth + _horSpace) / 2;
        int yCoord = _borderTop + m_Level * ( shapeHeight + _verSpace );

        setPosition(new Point(xCoord, yCoord));
        setSize(new Size(shapeWidth, shapeHeight));
    }
    
/*
    @Override
    public String toString(){
        return "toString: " + m_sRectangleName + " id: " + m_ID + " level: " + m_Level + " num: " + getPosNum();
    }
*/
}
