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

    public void stop(){
        if(m_list.size() > 0)
            for(int i = 0; i< m_list.size(); i++)
                m_list.get(i).stop();
    }

    public void clearList(){
        if(m_list.size() > 0) {
            for(int i = 0; i< m_list.size(); i++){
                m_list.get(i).stop();
                m_list.get(i).removeListener();
            }
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

}
