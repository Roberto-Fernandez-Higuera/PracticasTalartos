/**
 * @author Roberto Fernández Higuera
 */

package POJOS;

import java.util.Date;

public class Capataz implements java.io.Serializable {

    private double idCapataz;
    private double dia;
    private double numApoyos;
    private double fijoSalida;
    private double longMantenimiento;
    private double anomalia;
    private double longApertura;
    private double talasFuera;
    private double longitudLimpieza;
    private double limpiezaBase;
    private double km;
    private double importeMedios;
    private double importeCoeficiente;
    private String zona;
    private String observaciones;
    private String codLinea;
    private double podaCalle;
    private String nombreApoyo;
    private double numDiasTrabajados;
    private String pendienteTractor;
    private String trabajoRematado;

    public Capataz() {

    }

    public Capataz(double dia, double numApoyos, double fijoSalida, double longMantenimiento, double anomalia, double longApertura, double talasFuera, double longitudLimpieza, double limpiezaBase, double km, double importeMedios, double importeCoeficiente, String zona, String observaciones, String codLinea, double podaCalle, String nombreApoyo, double numDiasTrabajados, String pendienteTractor, String trabajoRematado) {
        this.dia = dia;
        this.numApoyos = numApoyos;
        this.fijoSalida = fijoSalida;
        this.longMantenimiento = longMantenimiento;
        this.anomalia = anomalia;
        this.longApertura = longApertura;
        this.talasFuera = talasFuera;
        this.longitudLimpieza = longitudLimpieza;
        this.limpiezaBase = limpiezaBase;
        this.km = km;
        this.importeMedios = importeMedios;
        this.importeCoeficiente = importeCoeficiente;
        this.zona = zona;
        this.observaciones = observaciones;
        this.codLinea = codLinea;
        this.podaCalle = podaCalle;
        this.nombreApoyo = nombreApoyo;
        this.numDiasTrabajados = numDiasTrabajados;
        this.pendienteTractor = pendienteTractor;
        this.trabajoRematado = trabajoRematado;
    }

    public double getIdCapataz() {
        return idCapataz;
    }

    public void setIdCapataz(double idCapataz) {
        this.idCapataz = idCapataz;
    }

    public double getDia() {
        return this.dia;
    }

    public void setDia(double dia) {
        this.dia = dia;
    }

    public double getNumApoyos() {
        return this.numApoyos;
    }

    public void setNumApoyos(double numApoyos) {
        this.numApoyos = numApoyos;
    }

    public double getFijoSalida() {
        return this.fijoSalida;
    }

    public void setFijoSalida(double fijoSalida) {
        this.fijoSalida = fijoSalida;
    }

    public double getLongMantenimiento() {
        return this.longMantenimiento;
    }

    public void setLongMantenimiento(double longMantenimiento) {
        this.longMantenimiento = longMantenimiento;
    }

    public double getAnomalia() {
        return this.anomalia;
    }

    public void setAnomalia(double anomalia) {
        this.anomalia = anomalia;
    }

    public double getLongApertura() {
        return this.longApertura;
    }

    public void setLongApertura(double longApertura) {
        this.longApertura = longApertura;
    }

    public double getTalasFuera() {
        return this.talasFuera;
    }

    public void setTalasFuera(double talasFuera) {
        this.talasFuera = talasFuera;
    }

    public double getLongitudLimpieza() {
        return this.longitudLimpieza;
    }

    public void setLongitudLimpieza(double longitudLimpieza) {
        this.longitudLimpieza = longitudLimpieza;
    }

    public double getLimpiezaBase() {
        return this.limpiezaBase;
    }

    public void setLimpiezaBase(double limpiezaBase) {
        this.limpiezaBase = limpiezaBase;
    }

    public double getKm() {
        return this.km;
    }

    public void setKm(double km) {
        this.km = km;
    }

    public double getImporteMedios() {
        return this.importeMedios;
    }

    public void setImporteMedios(double importeMedios) {
        this.importeMedios = importeMedios;
    }

    public double getImporteCoeficiente() {
        return this.importeCoeficiente;
    }

    public void setImporteCoeficiente(double importeCoeficiente) {
        this.importeCoeficiente = importeCoeficiente;
    }

    public String getZona() {
        return this.zona;
    }

    public void setZona(String zona) {
        this.zona = zona;
    }

    public String getObservaciones() {
        return this.observaciones;
    }

    public void setObservaciones(String observaciones) {
        this.observaciones = observaciones;
    }

    public String getCodLinea() {
        return this.codLinea;
    }

    public void setCodLinea(String codLinea) {
        this.codLinea = codLinea;
    }

    public double getPodaCalle() {
        return this.podaCalle;
    }

    public void setPodaCalle(double podaCalle) {
        this.podaCalle = podaCalle;
    }

    public String getNombreApoyo() {
        return this.nombreApoyo;
    }

    public void setNombreApoyo(String nombreApoyo) {
        this.nombreApoyo = nombreApoyo;
    }

    public double getNumDiasTrabajados() {
        return this.numDiasTrabajados;
    }

    public void setNumDiasTrabajados(double numDiasTrabajados) {
        this.numDiasTrabajados = numDiasTrabajados;
    }

    public String getPendienteTractor() {
        return this.pendienteTractor;
    }

    public void setPendienteTractor(String pendienteTractor) {
        this.pendienteTractor = pendienteTractor;
    }

    public String getTrabajoRematado() {
        return this.trabajoRematado;
    }

    public void setTrabajoRematado(String trabajoRematado) {
        this.trabajoRematado = trabajoRematado;
    }
}