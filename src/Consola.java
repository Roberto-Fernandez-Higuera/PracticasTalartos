/**
 * @author Roberto Fernández Higuera
 */

public class Consola {

    public Consola(){

    }

    private ExcelManager excelManager = new ExcelManager();

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma(){
        excelManager.creacionExcelApoyosRealizados();
        excelManager.creacionExcelControlCapataces();
    }

}