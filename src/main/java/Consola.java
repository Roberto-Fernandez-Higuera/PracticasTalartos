/**
 * @author Roberto Fernández Higuera
 */

import java.util.Scanner;

public class Consola {

    public String nombreHojaApoyos = "";
    public String nombreHojaCapataces = "";

    public Consola() {

    }

    private ExcelManager excelManager = new ExcelManager();
    private ExcelManagerCapataces excelManagerCapataces = new ExcelManagerCapataces();

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma() {

        System.out.print("*******************************************\n");
        System.out.print("**PROGRAMA MEDICIONES PARTES TALARTOS S.L**\n");
        System.out.print("*******************************************\n");

        System.out.print(" \n");

        Scanner scanner = new Scanner(System.in);

        System.out.print("Introduce el nombre de la hoja(línea sobre la que quieres realizar cambios): \n");
        String nombreHojaApoyos = scanner.nextLine();

        scanner.close();

        excelManager.creacionExcelApoyosRealizados(nombreHojaApoyos);

        Scanner sc = new Scanner(System.in);

        System.out.print("Introduce el nombre de la hoja(capataz sobre el que quieres realizar cambios): \n");
        String nombreHojaCapataces = scanner.nextLine();

        scanner.close();

        excelManagerCapataces.creacionExcelControlCapataces(nombreHojaCapataces);
    }

}