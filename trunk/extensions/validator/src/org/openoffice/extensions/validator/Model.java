package org.openoffice.extensions.validator;

import java.util.List;
import java.util.ArrayList;


public class Model {

    Controller m_Controller = null;
    List<Cell> m_list = null;

    public Model( Controller controller ){
        m_Controller = controller;
        m_list = new ArrayList<Cell>();
    }

    public Controller getController(){
        return m_Controller;
    }

    public void addItemToList(Cell cell){
        m_list.add(cell);
    }

    public void clearList(){
        if(m_list.size() > 0) {
            for(int i = 0; i< m_list.size(); i++)
                m_list.get(i).removeListener();
            m_list.clear();
        }
    }

    public Cell getItemFromList(short itemPos){
        return m_list.get(itemPos);
    }

    public void deleteItemFromList(int itemPos){
        m_list.get(itemPos).removeListener();
        for(int i = itemPos + 1; i< m_list.size(); i++)
            m_list.get(i).decreaseNumber();
        m_list.remove(itemPos);
    }

/*
    // need to listen error3 cell always becouse there are cells not formulas without listeners
    // runCorrectedError3 is a semaphor to close other functions in Gui.callHandlerMethod() itemchanged component
    public void refreshError3Cells(){
        for(int i = 0; i< m_list.size(); i++){
            Cell cell = m_list.get(i);
            if( cell.getErrorType() == 3 ) {
                cell.setPrecedentsCells();
                if(cell.getNotValidPrecCells() == null || cell.getNotValidPrecCells().isEmpty() )
                        cell.clearPrecedentsCellsAndType();
                if( cell.getErrorType() == 0 ){
                    short itemPos = getController().getGui().getSelectedItemPosFromListBox();
                    short cellNum = cell.getNumber();
                    if( itemPos != cellNum ){
                        getController().getGui().runCorrectedError3 = true;
                        getController().getGui().removeItemFromListBox( cellNum );
                        getController().getGui().addItemToListBox( cell.toString(), cellNum );
                        getController().getGui().selectItemPosInListBox( itemPos );
                        getController().getGui().runCorrectedError3 = false;
                    }
                }
            }
        }
    }

    public List<Cell> getList(){
        return m_list;
    }  
 */

}
