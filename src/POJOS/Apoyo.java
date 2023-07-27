/**
 * @author Roberto Fernández Higuera
 */

package src.POJOS;
import java.util.Date;


public class Apoyo{
    private Integer idApoyo;
    private Integer numApoyo;
    private Integer longitudMantenimineto;
    private Integer longitudLimpieza;
    private Integer longitudApertura;
    private Integer numAnomalia;
    private Integer limpiezaBase;
    private Integer podaCalle;
    private Integer fijoSalida;
    private Date dia;
    private String capataz;
    private String observaciones;

    public Apoyo(){
    }

    public Apoyo(Integer numApoyo, Integer longitudMantenimineto, Integer longitudLimpieza, Integer longitudApertura, Integer numAnomalia, Integer limpiezaBase, Integer podaCalle, Integer fijoSalida, Date dia, String capataz, String observaciones){
        this.numApoyo = numApoyo;
        this.longitudMantenimineto = longitudMantenimineto;
        this.longitudLimpieza = longitudLimpieza;
        this.longitudApertura = longitudApertura;
        this.numAnomalia = numAnomalia;
        this.limpiezaBase = limpiezaBase;
        this.podaCalle = talasFueraCalle;
        this.fijoSalida;
        this.dia = dia;
        this.capataz = capataz;
        this.observaciones = observaciones;
    }

    public Integer getIdApoyo() {
        return this.idApoyo;
    }
    public void setIdApoyo(Integer idApoyo) {
        this.idApoyo = idApoyo;
    }

    public Integer getNumApoyo() {
        return numApoyo;
    }

    public void setNumApoyo(Integer numApoyo) {
        this.numApoyo = numApoyo;
    }

    public Integer getLongitudMantenimineto() {
        return this.longitudMantenimineto;
    }

    public void setLongitudMantenimineto(Integer longitudMantenimineto) {
        this.longitudMantenimineto = longitudMantenimineto;
    }

    public Integer getLongitudLimpieza() {
        return this.longitudLimpieza;
    }

    public void setLongitudLimpieza(Integer longitudLimpieza) {
        this.longitudLimpieza = longitudLimpieza;
    }

    public Integer getLongitudApertura() {
        return this.longitudApertura;
    }

    public void setLongitudApertura(Integer longitudApertura) {
        this.longitudApertura = longitudApertura;
    }

    public Integer getNumAnomalia() {
        return this.numAnomalia;
    }

    public void setNumAnomalia(Integer numAnomalia) {
        this.numAnomalia = numAnomalia;
    }

    public Integer getLimpiezaBase() {
        return this.limpiezaBase;
    }

    public void setLimpiezaBase(Integer limpiezaBase) {
        this.limpiezaBase = limpiezaBase;
    }

    public Integer getPodaCalle() {
        return this.talasFueraCalle;
    }

    public void setPodaCalle(Integer talasFueraCalle) {
        this.talasFueraCalle = talasFueraCalle;
    }

    public Integer getFijoSalida() {
        return this.fijoSalida;
    }

    public void setFijoSalida(Integer fijoSalida) {
        this.fijoSalida = fijoSalida;
    }

    public Date getDia() {
        return this.dia;
    }

    public void setDia(Date dia) {
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
