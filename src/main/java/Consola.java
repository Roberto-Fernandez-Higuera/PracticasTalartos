/**
 * @author Roberto Fernández Higuera
 */

import java.util.Scanner;

public class Consola {

    public String nombreHojaApoyos = "";
    public String nombreHojaCapataces = "";

    public Consola() {

    }

    private String nombreExcel;
    private ExcelManager excelManager = new ExcelManager(nombreExcel);
    private String nombreExcelApoyos;
    private ExcelManagerCapataces excelManagerCapataces = new ExcelManagerCapataces(nombreExcelApoyos);

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma() {

        System.out.print("*******************************************\n");
        System.out.print("**PROGRAMA MEDICIONES PARTES TALARTOS S.L**\n");
        System.out.print("*******************************************\n");

        System.out.println(" \n");

        System.out.print("Es importante que introduzcas correctamente los nombres.\n");

        System.out.println(" \n");

        /**
         * PARTE EXCEL APOYOS
         */

        Scanner scanner = new Scanner(System.in);

        System.out.print("Introduce el nombre de la línea (línea sobre la que quieres realizar cambios): \n");
        String nombreHojaApoyos = scanner.nextLine();

        scanner.close();

        Scanner scCodigo = new Scanner(System.in);

        System.out.print("Introduce el código de dicha línea (sobre la que vamos a trabajar): \n");
        String codigoHoja = scCodigo.nextLine();

        scanner.close();

        Scanner scNombreExcel = new Scanner(System.in);

        System.out.print("Introduce el nombre del Excel sobre el que vas a realizar cambios (si existe accedes a él, de lo contrario creará uno nuevo): \n");
        String nombreExcel = scNombreExcel.nextLine();

        scanner.close();

        excelManager.creacionExcelApoyosRealizados(nombreHojaApoyos, codigoHoja, nombreExcel);


        /**
         * PARTE EXCEL CAPATACES
         */

        Scanner scSiONo = new Scanner(System.in);

        System.out.print("EXCEL DE APOYOS CREADO, ¿QUIÉRES GENERAR EL DE CAPATACES  Y/N? (Y para Sí, N para No)\n");
        String siONo = scSiONo.nextLine();

        scanner.close();

        if (siONo.equals("Y")) {
            Scanner sc = new Scanner(System.in);

            System.out.print("Introduce el nombre de la hoja (capataz sobre la que quieres realizar cambios): \n");
            String nombreHojaCapataces = scanner.nextLine();

            scanner.close();

            Scanner scNombreExcelCapataces = new Scanner(System.in);

            System.out.print("Introduce el nombre del Excel sobre el que quieres trabajar (si existe accedes a él, de lo contrario creará uno nuevo): \n");
            String nombreExcelCapataces = scNombreExcelCapataces.nextLine();

            scanner.close();

            excelManagerCapataces.creacionExcelControlCapataces(nombreHojaCapataces, nombreExcelCapataces);
        } else {

        }
    }

}