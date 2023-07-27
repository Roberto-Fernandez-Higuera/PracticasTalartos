/**
 * @author Roberto Fernández Higuera
 */

package POJOS;

import java.util.Date;

public class Capataz() {

    private Integer idCapataz;
    private Date dia;
    private Integer numApoyos;
    private Integer fijoSalida;
    private Integer longMantenimiento;
    private Integer anomalia;
    private Integer longApertura;
    private Integer talasFuera;
    private Integer longitudCopa;
    private Integer limpiezaBase;
    private Integer km;
    private Integer importeMedios;
    private Integer importeCoeficiente;
    private String zona;
    private String observaciones;
    private String codLinea;

    public Capataz() {

    }

    public Capataz(Date dia, Integer numApoyos, Integer fijoSalida, Integer longMantenimiento, Integer anomalia, Integer longApertura, Integer talasFuera, Integer longitudCopa, Integer limpiezaBase, Integer km, Integer importeMedios, Integer importeCoeficiente, String zona, String observaciones, String codLinea) {
        this.dia = dia;
        this.numApoyos = numApoyos;
        this.fijoSalida = fijoSalida;
        this.longMantenimiento = longMantenimiento;
        this.anomalia = anomalia;
        this.longApertura = longApertura;
        this.talasFuera = talasFuera;
        this.longitudCopa = longitudCopa;
        this.limpiezaBase = limpiezaBase;
        this.km = km;
        this.importeMedios = importeMedios;
        this.importeCoeficiente = importeCoeficiente;
        this.zona = zona;
        this.observaciones = observaciones;
        this.codLinea = codLinea;
    }

    public Integer getIdCapataz() {
        return idCapataz;
    }

    public void setIdCapataz(Integer idCapataz) {
        this.idCapataz = idCapataz;
    }

    public Date getDia() {
        return this.dia;
    }

    public void setDia(Date dia) {
        this.dia = dia;
    }

    public Integer getNumApoyos() {
        return this.numApoyos;
    }

    public void setNumApoyos(Integer numApoyos) {
        this.numApoyos = numApoyos;
    }

    public Integer getFijoSalida() {
        return this.fijoSalida;
    }

    public void setFijoSalida(Integer fijoSalida) {
        this.fijoSalida = fijoSalida;
    }

    public Integer getLongMantenimiento() {
        return this.LongMantenimiento;
    }

    public void setLongMantenimiento(Integer longMantenimiento) {
        this.longMantenimiento = longMantenimiento;
    }

    public Integer getAnomalia() {
        return this.anomalia;
    }

    public void setAnomalia(Integer anomalia) {
        this.anomalia = anomalia;
    }

    public Integer getLongApertura() {
        return this.longApertura;
    }

    public void setLongApertura(Integer longApertura) {
        this.longApertura = longApertura;
    }

    public Integer getTalasFuera() {
        return this.talasFuera;
    }

    public void setTalasFuera(Integer talasFuera) {
        this.talasFuera = talasFuera;
    }

    public Integer getLongitudCopa() {
        return this.longitudCopa;
    }

    public void setLongitudCopa(Integer longitudCopa) {
        this.longitudCopa = longitudCopa;
    }

    public Integer getLimpiezaBase() {
        return this.limpiezaBase;
    }

    public void setLimpiezaBase(Integer limpiezaBase) {
        this.limpiezaBase = limpiezaBase;
    }

    public Integer getKm() {
        return this.km;
    }

    public void setKm(Integer km) {
        this.km = km;
    }

    public Integer getImporteMedios() {
        return this.importeMedios;
    }

    public void setImporteMedios(Integer importeMedios) {
        this.importeMedios = importeMedios;
    }

    public Integer getImporteCoeficiente() {
        return this.importeCoeficiente;
    }

    public void setImporteCoeficiente(Integer importeCoeficiente) {
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
}