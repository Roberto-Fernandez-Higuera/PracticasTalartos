/**
 * @author Roberto Fernández Higuera
 */

import java.util.Scanner;

public class Consola {

    static String provincias;
    static String zona;

    public Consola() {

    }

    private ExcelManager excelManager = new ExcelManager();

    //No creamos el objeto aún para realizar una (lazy initialization)
    private ExcelManagerCapataces excelManagerCapataces;

    /**
     * Método encargado de la creación del Excel de apoyos realizados
     */
    public void ejecucionPrograma(){

        /**
         * PARTE EXCEL APOYOS
         */

        Scanner scCodigo = new Scanner(System.in);

        System.out.print("Introduce el código de dicha línea (sobre la que vamos a trabajar): \n");
        String codigoHoja = scCodigo.nextLine();


        Scanner scNombreExcel = new Scanner(System.in);

        System.out.print("Introduce el nombre del Excel APOYOS sobre el que vas a realizar cambios (si existe accedes a él, de lo contrario creará uno nuevo): \n");
        String nombreExcel = scNombreExcel.nextLine();

        excelManager.creacionExcelApoyosRealizados(ExcelManager.linea, codigoHoja, nombreExcel);


        /**
         * PARTE EXCEL CAPATACES
         */


        Scanner scSiONo = new Scanner(System.in);

        System.out.print("EXCEL DE APOYOS CREADO, ¿QUIÉRES GENERAR EL DE CAPATACES  Y/N? (Y para Sí, N para No)\n");
        String siONo = scSiONo.nextLine();

        //Si decide crear el CONTROL DE CAPATACES, ENTRO EN EL IF Y LO GENERO, DE LO CONTRARIO, SALGO
        if (siONo.equals("Y")) {

            Scanner scNombreExcelCapataces = new Scanner(System.in);

            System.out.print("Introduce el nombre del Excel de CONTROL DE CAPATACES sobre el que quieres trabajar (si existe accedes a él, de lo contrario creará uno nuevo): \n");
            String nombreExcelCapataces = scNombreExcelCapataces.nextLine();

            Scanner scProvincias = new Scanner(System.in);

            System.out.print("Introduce el nombre de la carpeta de PROVINCIAS con la que quieres trabajar: \n");
            provincias = scProvincias.nextLine();

            Scanner scZona = new Scanner(System.in);

            System.out.print("Introduce el nombre de la zona de "+provincias+" con la que quieres trabajar: \n");
            zona = scZona.nextLine();

            Scanner scNombreExcelParaCapataces = new Scanner(System.in);

            System.out.print("Introduce el nombre del Excel de APOYOS sobre el que quieres generar CONTROL CAPATACES: \n");
            String nombreExcelParaCapataces = scNombreExcelParaCapataces.nextLine();

            excelManagerCapataces = new ExcelManagerCapataces(nombreExcelParaCapataces);
            excelManagerCapataces.creacionExcelControlCapataces(nombreExcelCapataces, ExcelManager.linea ,codigoHoja);

        } else {

        }
    }
}