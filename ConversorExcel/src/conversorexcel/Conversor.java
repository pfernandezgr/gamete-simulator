/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package conversorexcel;

//import java.io.*; 
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author movfab03
 */
public class Conversor {

    public void convertirExcel(String path, int eventos, String vacio) {
        //contador de filas
        int fila = 1;

        //primera fila de un marker
        int filaActual = 1;
        FileInputStream fis = null;
        try {
            //Abre el archivo
            File myFile = new File(path);
            fis = new FileInputStream(myFile);
            // Genera la instancia de un workbook del XLSX
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
            // Devuelve la primera página del libro excel
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            //Creo el archivo nuevo
            XSSFWorkbook archivo = new XSSFWorkbook();
//Creo el segundo archivo
            XSSFWorkbook SegundoArchivo = new XSSFWorkbook();
            XSSFSheet hojaSegundoArchivo = SegundoArchivo.createSheet("Sheet 1");
            //contador auxiliar para el segundo archivo
            int contadorColumnas = 2;
            //Recorro todas las filas del excel
            while (mySheet.getRow(fila) != null && mySheet.getRow(fila).getCell(0) != null) {

                filaActual = fila;

                //Nombre del marker
                String unMarker = mySheet.getRow(fila).getCell(0).getStringCellValue();
                System.out.println("Marker " + unMarker);
                //Creo una hoja del nuevo archivo excel con el nombre del marker
                XSSFSheet hojaActual = archivo.createSheet(unMarker);

                boolean salir = false;
                do {

                    fila += 1;
                    //chequea que la fila y la celda existan
                    if (mySheet.getRow(fila) != null && mySheet.getRow(fila).getCell(0) != null) {
                        // Obtiene el nombre del marker desde la planilla
                        String marker = mySheet.getRow(fila).getCell(0).getStringCellValue();
                        //chequea hasta donde se repite el mismo nombre de marker, cuando cambia sale del loop
                        if (!marker.equals(unMarker)) {
                            salir = true;
                        }
                    } else {
                        salir = true;
                    }
                } while (!salir);

                //Cantidad de filas que tienen el mismo nombre de marker actual
                int filasDelMarker = fila - filaActual;

                int contadorFilasTotal = 0;

                int columna = 2;
                ArrayList<celda> array = new ArrayList();
                //Recorro todas las columnas
                while (mySheet.getRow(0).getCell(columna) != null) {
//Obtiene el nombre del marker  
                    String marcador = mySheet.getRow(0).getCell(columna).getStringCellValue();
//variable auxiliar para contar los eventos
                    int eventosResto = eventos;
//celdas con el mismo marker en la columna actual
                    array.clear();

                    //recorro todas las filas que tengan el mismo nombre de marker
                    for (int i = 0; i < filasDelMarker; i++) {
                        Double numero;
                        String nombre = "";
                        String alelo = String.valueOf(mySheet.getRow(i + filaActual).getCell(1).getNumericCellValue());
                        if (mySheet.getRow(i + filaActual).getCell(columna) != null) {
                            numero = mySheet.getRow(i + filaActual).getCell(columna).getNumericCellValue();
                            //Genera el nombre de la celda, ej 0010000
                            if (numero > 0) {
                                for (int j = 0; j < filasDelMarker; j++) {
                                    if (j == i) {
                                        nombre += "1";
                                    } else {
                                        nombre += "0";
                                    }
                                }
                            }

                        } else {
                            numero = 0.0;
                        }
                        int cantidad = (int) Math.floor(numero * eventos);
                        eventosResto -= cantidad;
                        double resto = (numero * eventos) - cantidad;
                        array.add(new celda(nombre, cantidad, resto, alelo));

                    }

//Esta parte se hace cargo del redondeo
//eventosResto son la cantidad de eventos que no entraron enteros en ninguna celda
//y se reparten de acuerdo al que tenga mayor resto
                    while (eventosResto > 0) {
                        if (array.size() > 1) {
                            celda celdaMayor = array.get(0);

                            for (celda unaCelda : array) {
                                if (unaCelda.getResto() > celdaMayor.getResto()) {
                                    celdaMayor = unaCelda;
                                }
                            }
                            if (celdaMayor.getResto() == 0) {
                                array.add(new celda(vacio, eventosResto, 0.0, vacio));
                                eventosResto = 0;
                            } else {
                                celdaMayor.sumarUnoACantidad();
                                celdaMayor.setResto(0.0);
                            }
                        } else if (array.size() == 1) {
                            if (array.get(0).getResto() == 0) {
                                //Genera un nombre de celda ej. 99999 y escribe esto en la cantidad de eventos definidos

                                array.add(new celda(vacio, eventosResto, 0.0, vacio));
                                eventosResto = 0;
                            } else {
                                array.get(0).sumarUnoACantidad();
                            }
                        }

                        eventosResto -= 1;
                    }

                    List<String> solution = new ArrayList<>();
                    int contadorEventos = 1;
//Escribe en el nuevo archivo Excel las celdas para una columna de un marker
//escribirá tantas celdas como eventos se hayan indicado
                    for (celda unaCelda : array) {

//agrega el nombre del marker al encabezado del segundo archivo
escribirDatoEnHoja(hojaSegundoArchivo, 0, contadorColumnas, unMarker);


                        for (int l = 0; l < unaCelda.getCantidad(); l++) {
                            //Escribe los datos en la hoja que corresponde del primer archivo
                            escribirDatoEnHoja(hojaActual, contadorFilasTotal, 0, marcador);
                            escribirDatoEnHoja(hojaActual, contadorFilasTotal, 1, contadorEventos);
                            escribirDatoEnHoja(hojaActual, contadorFilasTotal, 2, unaCelda.getNombre());
                            escribirDatoEnHoja(hojaActual, contadorFilasTotal, 3, unaCelda.getAlelo());
                            solution.add(unaCelda.getAlelo());

                            //Si es el primer marker crea los encabezados y las primeras 2 columnas del segundo archivo
                            if (contadorColumnas == 2) {
                                escribirDatoEnHoja(hojaSegundoArchivo, 0, 0, "POP ID");
                                escribirDatoEnHoja(hojaSegundoArchivo, 0, 1, "POP NUMBER CODE");
                                escribirDatoEnHoja(hojaSegundoArchivo, 0, 2, "GAMETES");
                                escribirDatoEnHoja(hojaSegundoArchivo, contadorFilasTotal + 1, 0, marcador);
                                escribirDatoEnHoja(hojaSegundoArchivo, contadorFilasTotal + 1, 1, columna -1);
                                escribirDatoEnHoja(hojaSegundoArchivo, contadorFilasTotal + 1, 2, contadorEventos);
                            }
                           
                            contadorEventos += 1;
                            contadorFilasTotal += 1;
                        }

                    }
                    Collections.shuffle(solution);
                    for (int l = 0; l < solution.size(); l++) {
                        
                       String sol = solution.get(l).replace(".0", "");
                       if(sol.matches("\\d+")){
                           int solucion = Integer.parseInt(sol);
                            //Escribo la solución en el primer archivo
                              escribirDatoEnHoja(hojaActual, contadorFilasTotal - l - 1, 4, solucion);
                        //Escribo la solución en el segundo archivo
                         escribirDatoEnHoja(hojaSegundoArchivo, contadorFilasTotal - l, contadorColumnas, solucion);
                       }else{
                             //Escribo la solución en el primer archivo
                              escribirDatoEnHoja(hojaActual, contadorFilasTotal - l - 1, 4, solution.get(l));
                        //Escribo la solución en el segundo archivo
                         escribirDatoEnHoja(hojaSegundoArchivo, contadorFilasTotal - l, contadorColumnas, solution.get(l));
                       }
                      
                        
                        
                        }
                    columna += 1;
                    
                }
contadorColumnas += 1;
            }

//Escribir el primer archivo
            JFileChooser chooser = new JFileChooser();
            chooser.setFileFilter(new XLSXFilter());
            chooser.setDialogTitle("Save the first file as...");
            chooser.setCurrentDirectory(new File("/home/me/Documents"));
            int retrival = chooser.showSaveDialog(null);
            File myOtherFile;
            if (retrival == JFileChooser.APPROVE_OPTION) {

                String ruta = chooser.getSelectedFile().getAbsolutePath();
                if (!ruta.endsWith(".xlsx") && !ruta.endsWith(".XLSX")) {
                    myOtherFile = new File(ruta + ".XLSX");
                } else {
                    myOtherFile = new File(ruta);
                }

                FileOutputStream f = new FileOutputStream(myOtherFile);
                archivo.write(f);
                f.close();
                Desktop dt = Desktop.getDesktop();
                 //Abre el archivo que se creó
                dt.open(myOtherFile);

            }
           

             JFileChooser chooser2 = new JFileChooser();
            chooser2.setFileFilter(new XLSXFilter());
            chooser2.setDialogTitle("Save the second file as...");
            chooser2.setCurrentDirectory(new File("/home/me/Documents"));
            int retrival2 = chooser2.showSaveDialog(null);
            File myOtherFile2;
            if (retrival2 == JFileChooser.APPROVE_OPTION) {

                String ruta2 = chooser2.getSelectedFile().getAbsolutePath();
                if (!ruta2.endsWith(".xlsx") && !ruta2.endsWith(".XLSX")) {
                    myOtherFile2 = new File(ruta2 + ".XLSX");
                } else {
                    myOtherFile2 = new File(ruta2);
                }

                FileOutputStream f2 = new FileOutputStream(myOtherFile2);
                SegundoArchivo.write(f2);
                f2.close();
                Desktop dt = Desktop.getDesktop();
                dt.open(myOtherFile2);

            }
            
            
            
            
            
            
            
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "An error occurred \n" + ex.getLocalizedMessage(), "Error", JOptionPane.ERROR_MESSAGE);

            ex.printStackTrace();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "An error occurred \n" + ex.getLocalizedMessage(), "Error", JOptionPane.ERROR_MESSAGE);

            ex.printStackTrace();

        } finally {
            try {
                fis.close();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "An error occurred \n" + ex.getLocalizedMessage(), "Error", JOptionPane.ERROR_MESSAGE);

                ex.printStackTrace();
            }
        }

    }

    public String[] Randomize(String[] arr) {
        String[] randomizedArray = new String[arr.length];
        System.arraycopy(arr, 0, randomizedArray, 0, arr.length);
        Random rgen = new Random();

        for (int i = 0; i < randomizedArray.length; i++) {
            int randPos = rgen.nextInt(randomizedArray.length);
            String tmp = randomizedArray[i];
            randomizedArray[i] = randomizedArray[randPos];
            randomizedArray[randPos] = tmp;
        }

        return randomizedArray;
    }

    /**
     * Fisher–Yates shuffle.
     */
    public void escribirDatoEnHoja(XSSFSheet hoja, int fila, int columna, Object dato) {

        Row unaFila = hoja.getRow(fila);
        if (unaFila == null) {
            unaFila = hoja.createRow(fila);
        }
        Cell cell = hoja.getRow(fila).getCell(columna);
        if (cell == null) {
            cell = hoja.getRow(fila).createCell(columna);
        }
        if (dato instanceof Integer) {
            Integer i = (Integer) dato;

            cell.setCellValue(i);
        }
        if (dato instanceof String) {

            String i = (String) dato;
            cell.setCellValue(i);

        }
        if (dato instanceof Double) {
            Double i = (Double) dato;
            cell.setCellValue(i);
        }

    }
}
