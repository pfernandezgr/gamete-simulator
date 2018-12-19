/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package conversorexcel;

/**
 *
 * @author movfab03
 */
public class celda {
    //ej. "0001000"
    private String nombre;
    private int cantidad;
    private double resto;
    private String alelo;

    public String getAlelo() {
        return alelo;
    }

    public void setAlelo(String alelo) {
        this.alelo = alelo;
    }

    public celda(String nombre, int cantidad, double resto, String alelo) {
        this.nombre = nombre;
        this.cantidad = cantidad;
        this.resto = resto;
        this.alelo = alelo;
        
    }
    @Override
    public String toString(){
    return "Nombre: "+ nombre + " Cantidad: "+ cantidad+" Resto: "+resto;
}
    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public int getCantidad() {
        return cantidad;
    }
public void sumarUnoACantidad(){
    cantidad += 1;
}
    public void setCantidad(int cantidad) {
        this.cantidad = cantidad;
    }

    public double getResto() {
        return resto;
    }

    public void setResto(double resto) {
        this.resto = resto;
    }
    
}
