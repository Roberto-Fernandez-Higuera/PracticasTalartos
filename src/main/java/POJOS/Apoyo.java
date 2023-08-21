/**
 * @author Roberto Fernández Higuera
 */

package POJOS;

public class Apoyo implements java.io.Serializable {
    private double idApoyo;
    private String numApoyo;
    private double longitudMantenimineto;
    private double longitudLimpieza;
    private double longitudApertura;
    private double numAnomalia;
    private double limpiezaBase;
    private double podaCalle;
    private double fijoSalida;
    private double dia;
    private String capataz;
    private String observaciones;

    public Apoyo() {
    }

    public Apoyo(String numApoyo, double longitudMantenimineto, double longitudLimpieza, double longitudApertura, double numAnomalia, double limpiezaBase, double podaCalle, double fijoSalida, double dia, String capataz, String observaciones) {
        this.numApoyo = numApoyo;
        this.longitudMantenimineto = longitudMantenimineto;
        this.longitudLimpieza = longitudLimpieza;
        this.longitudApertura = longitudApertura;
        this.numAnomalia = numAnomalia;
        this.limpiezaBase = limpiezaBase;
        this.podaCalle = podaCalle;
        this.fijoSalida = fijoSalida;
        this.dia = dia;
        this.capataz = capataz;
        this.observaciones = observaciones;
    }

    public double getIdApoyo() {
        return this.idApoyo;
    }

    public void setIdApoyo(double idApoyo) {
        this.idApoyo = idApoyo;
    }

    public String getNumApoyo() {
        return numApoyo;
    }

    public void setNumApoyo(String numApoyo) {
        this.numApoyo = numApoyo;
    }

    public double getLongitudMantenimineto() {
        return this.longitudMantenimineto;
    }

    public void setLongitudMantenimiento(double longitudMantenimineto) {
        this.longitudMantenimineto = longitudMantenimineto;
    }

    public double getLongitudLimpieza() {
        return this.longitudLimpieza;
    }

    public void setLongitudLimpieza(double longitudLimpieza) {
        this.longitudLimpieza = longitudLimpieza;
    }

    public double getLongitudApertura() {
        return this.longitudApertura;
    }

    public void setLongitudApertura(double longitudApertura) {
        this.longitudApertura = longitudApertura;
    }

    public double getNumAnomalia() {
        return this.numAnomalia;
    }

    public void setNumAnomalia(double numAnomalia) {
        this.numAnomalia = numAnomalia;
    }

    public double getLimpiezaBase() {
        return this.limpiezaBase;
    }

    public void setLimpiezaBase(double limpiezaBase) {
        this.limpiezaBase = limpiezaBase;
    }

    public double getPodaCalle() {
        return this.podaCalle;
    }

    public void setPodaCalle(double podaCalle) {this.podaCalle = podaCalle;
    }

    public double getFijoSalida() {
        return this.fijoSalida;
    }

    public void setFijoSalida(double fijoSalida) {
        this.fijoSalida = fijoSalida;
    }

    public double getDia() {
        return this.dia;
    }

    public void setDia(double dia) {
        this.dia = dia;
    }

    public String getCapataz() {
        return this.capataz;
    }

    public void setCapataz(String capataz) {
        this.capataz = capataz;
    }

    public String getObservaciones() {
        return this.observaciones;
    }

    public void setObservaciones(String observaciones) {
        this.observaciones = observaciones;
    }

}