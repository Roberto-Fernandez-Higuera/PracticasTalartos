/**
 * @author Roberto Fernández Higuera
 */

import java.io.IOException;
import java.util.Scanner;

public class Consola {

    public String nombreHojaApoyos = "";
    public String nombreHojaCapataces = "";

    public Consola() {

    }

    private String nombreExcel;
    private ExcelManager excelManager = new ExcelManager();
    private String nombreExcelApoyos;

    //No creamos el objeto aún para realizar una (lazy initialization)
    private ExcelManagerCapataces excelManagerCapataces;

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma() throws IOException {

        /**
         * PARTE EXCEL APOYOS
         */

        Scanner scanner = new Scanner(System.in);

        System.out.print("Introduce el nombre de la línea (línea sobre la que quieres realizar cambios): \n");
        String nombreHojaApoyos = scanner.nextLine();


        Scanner scCodigo = new Scanner(System.in);

        System.out.print("Introduce el código de dicha línea (sobre la que vamos a trabajar): \n");
        String codigoHoja = scCodigo.nextLine();


        Scanner scNombreExcel = new Scanner(System.in);

        System.out.print("Introduce el nombre del Excel sobre el que vas a realizar cambios (si existe accedes a él, de lo contrario creará uno nuevo): \n");
        String nombreExcel = scNombreExcel.nextLine();

        excelManager.creacionExcelApoyosRealizados(nombreHojaApoyos, codigoHoja, nombreExcel);


        /**
         * PARTE EXCEL CAPATACES
         */


        Scanner scSiONo = new Scanner(System.in);

        System.out.print("EXCEL DE APOYOS CREADO, ¿QUIÉRES GENERAR EL DE CAPATACES  Y/N? (Y para Sí, N para No)\n");
        String siONo = scSiONo.nextLine();

        //Si decide crear el CONTROL DE CAPATACES, ENTRO EN EL IF Y LO GENERO, DE LO CONTRARIO, SALGO
        if (siONo.equals("Y")) {

            Scanner scNombreExcelCapataces = new Scanner(System.in);

            System.out.print("Introduce el nombre del Excel sobre el que quieres trabajar (si existe accedes a él, de lo contrario creará uno nuevo): \n");
            String nombreExcelCapataces = scNombreExcelCapataces.nextLine();

            excelManagerCapataces = new ExcelManagerCapataces();
            excelManagerCapataces.creacionExcelControlCapataces(nombreExcelCapataces, nombreHojaApoyos ,codigoHoja);
        } else {

        }


    }

}