/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package conversorexcel;

import javax.swing.UIManager;

/**
 *
 * @author movfab03
 */
public class ConversorExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
          try { 
         //Cambia el aspecto de todo el programa a estilo de Windows                  
    UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel"); 
} catch (Exception ex) {  
    ex.printStackTrace(); 
}
        
          //Inicia el formulario principal
        frmPrincipal p = new frmPrincipal();
        p.setVisible(true);
        
    }
    
}
