/**
 * @author Roberto Fernández Higuera
 */
//package src;
public class Consola {

    public Consola() {

    }

    private ExcelManager excelManager = new ExcelManager();
    private ExcelManagerCapataces excelManagerCapataces = new ExcelManagerCapataces();

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma() {
        excelManager.creacionExcelApoyosRealizados();
        excelManagerCapataces.creacionExcelControlCapataces();
    }

}