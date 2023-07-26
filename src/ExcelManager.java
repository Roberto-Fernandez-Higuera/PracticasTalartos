/**
 * @author Roberto Fernández Higuera
 */
package src;

import POJOS.Apoyo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class ExcelManager(){

    private static FileInputStream file;
    private static XSSFWorkbook wb;
    private XSSFSheet hojaIberdrola;

    //Arraylist con los valores de todos los campos
    private ArrayList<Apoyo> listaApoyos = new ArrayList<>();

    //MAPAS A UTILIZAR
    private HashMap<Integer, Mediciones> mapaMediciones = new HashMap<>();

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEES LAS PARTES DEL EXCEL
     */
    public ExcelManager(){
        try {
            this.file = new FileInputStream("src/main/resources/RIVALAMORA-ESTADIO 5 MAYO.xlsx");
            this.wb = new XSSFWorkbook(file);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel 1");
            System.exit(-1);
        }
        hojaIberdrola = wb.getSheetAt(0);
        this.mapaMediciones = leerDatosMedicionesPartes();
    }

    /**
     * Setteo de los valores que introducimos a cada campo de las mediciones
     * @return
     */
    public HashMap leerDatosMedicionesPartes(){
        int numFilas = hojaIberdrola.getLastRowNum() - 22;

        for(int i = 5; i < numFilas; i++){
            Row fila = hojaIberdrola.getRow(i);
            if (fila != null && fila.getCell(0) != null){
               Apoyo apoyoAnyadir = new Apoyo();

               /**
                * ID FILA APOYO
                */
               Integer id = fila.getRowNum() + 1;
               apoyoAnyadir.setIdApoyo(id);

               /**
                * NUM APOYO
                */
                apoyoAnyadir.setNumApoyo(fila.getCell(1).getNumericCellValue());

               /**
                * LONGITUD MANTENIMINETO
                */
               if (fila.getCell(2) == null){
                   apoyoAnyadir.setLongitudMantenimineto(0);
               } else {
                   apoyoAnyadir.setLongitudMantenimineto(fila.getCell(2).getNumericCellValue());
               }

               /**
                * LONGITU LIMPIEZA
                */
               if (fila.getCell(3) == null){
                   apoyoAnyadir.setLongitudLimpieza(0);
               } else {
                   apoyoAnyadir.setLongitudLimpieza(fila.getCell(3).getNumericCellValue());
               }

               /**
                * LONGITUD APERTURA
                */
               if(fila.getCell(4) == null){
                   apoyoAnyadir.setLongitudApertura(0);
               } else {
                   apoyoAnyadir.setLongitudApertura(fila.getCell(4).getNumericCellValue());
               }

               /**
                * ANOMALÍA VEGETACIÓN
                */
               if (fila.getCell(5) == null){
                   apoyoAnyadir.setNumAnomalia(0);
               } else {
                   apoyoAnyadir.setNumAnomalia(fila.getCell(5).getNumericCellValue());
               }

                /**
                 * LIMPIEZA BASE
                 */
               if (fila.getCell(6) == null){
                   apoyoAnyadir.setLimpiezaBase(0);
               } else {
                   apoyoAnyadir.setLimpiezaBase(fila.getCell(6).getNumericCellValue());
               }

               /**
                * PODA CALLE
                */
               if (fila.getCell(7) == null){
                   apoyoAnyadir.setPodaCalle(0);
               } else {
                   apoyoAnyadir.setPodaCalle(fila.getCell(7).getNumericCellValue());
               }

               /**
                * FIJO SALIDA
                */
               if (fila.getCell(8) == null){
                   apoyoAnyadir.setFijoSalida(0);
               } else {
                   apoyoAnyadir.setFijoSalida(fila.getCell(8).getNumericCellValue());
               }

                /**
                 * DÍA
                 */
               apoyoAnyadir.setDia(fila.getCell(9).getDateCellValue());

               /**
                * CAPATAZ
                */
               apoyoAnyadir.setCapataz(fila.getCell(10).getStringCellValue());

               /**
                * OBSERVACIONES
                */
               if (fila.getCell(11) == null){
                   apoyoAnyadir.setObservaciones("");
               } else {
                   apoyoAnyadir.setObservaciones(fila.getCell(11).getStringCellValue());
               }
               listaApoyos.add(apoyoAnyadir);
               mapaMediciones.put(id, apoyoAnyadir);
            }
        }
        return mapaMediciones;
    }

    public void ejecucionPrograma(){
        creacionExcelApoyosRealizados();
        //creacionExcelControlCapataces();
    }

    /**
     * PARTE EXCEL APOYOS REALIZADOS
     */
    public void creacionExcelApoyosRealizados(){
        FileOutputStream fileMod = null;
        try{
            fileMod = new FileOutputStream("EXCELS FINALES/APOYOS REALIZADOS NOMBRE.xlsx");
        } FileNotFoundException e) {
            System.out.println("Error al crear EXCEL DE APOYOS\n");
            System.exit(-1);
        }

        //Método que va a crear y rellenar mi excel de apoyos
        introducirValoresApoyos();

        try {
            wb.write(fileMod);
        } catch (IOException e) {
            System.out.println("Error al escribir EXCELL APOYOS\n");
            System.exit(-1);
        }

        /**
         * DUDA, PREGUNTAR SI SE UTILIZA EL SIGUIENTE TRY AQUÍ CUANDO CREEMOS LOS DOS EXCELS
         */
        try {
            file.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar fichero");
            System.exit(-1);
        }

        try {
            fileMod.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar EXCEL");
            System.exit(-1);
        }
    }

    private void introducirValoresApoyos(){
        int numApoyo = 0;
        int longitudMantenimineto = 0;
        int longitudLimpieza = 0;
        int longitudApertura = 0;
        int anomaliaVegetacion = 0;
        int longitudCopa = 0;
        int limpiezaBase = 0;
        int podaCalle = 0;
        int fijoSalida = 0;
        Date dia = null;
        String capataz = "";
        int numDiasTrabajados = 0;
        String pendienteTractor = "";
        String trabajoRematado = "";
        String observaciones = "";

        for (int i = 0; i < listaApoyos.size(); i++){
            Row fila = hoja.createRow(i);

            numApoyo = listaApoyos.get(i).getNumApoyo();
            Cell celdaNumApoyo = fila.createCell(0);
            celdaNumApoyo.setCellValue(numApoyo);

            longitudMantenimineto = listaApoyos.get(i).getLongitudMantenimineto();
            Cell celdaLongitudMantenimiento = fila.createCell(1);
            celdaLongitudMantenimiento.setCellValue(longitudMantenimineto);

            longitudLimpieza = listaApoyos.get(i).getLongitudLimpieza();
            Cell celdaLongitudLimpeza = fila.createCell(2);
            celdaLongitudLimpeza.setCellValue(longitudLimpieza);

            longitudApertura = listaApoyos.get(i).getLongitudApertura();
            Cell celdaLongitudApertura = fila.createCell(3);
            celdaLongitudApertura.setCellValue(longitudApertura);

            anomaliaVegetacion = listaApoyos.get(i).getnumAnomalia();
            Cell celdaAnomaliaVegetacion = fila.createCell(4);
            celdaAnomaliaVegetacion.setCellValue(anomaliaVegetacion);

            /**
             * TODO Hablar con Inés sobre los campos que no están (CASOS ESPECIALES)
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaLongitudCopa = fila.createCell(5);
            celdaLongitudCopa.setCellValue(longitudCopa);

            limpiezaBase = listaApoyos.get(i).getLimpiezaBase();
            Cell celdaLimpiezaBase = fila.createCell(6);
            celdaLimpiezaBase.setCellValue(limpiezaBase);

            /**
             * CASO ESPECIAL
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaPodaCalle = fila.createCell(7);
            celdaPodaCalle.setCellValue(podaCalle);

            /**
             * CASO ESPECIAL
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaFijoSalida = fila.createCell(8);
            celdaFijoSalida.setCellValue(fijoSalida);

            dia = listaApoyos.get(i).getDia();
            Cell celdaDia = fila.createCell(9);
            celdaDia.setCellValue(dia);

            capataz = listaApoyos.get(i).getCapataz();
            Cell celdaCapataz = fila.createCell(10);
            celdaCapataz.setCellValue(capataz);

            /**
             * CASO ESPECIAL
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaNumDiasTrabajados = fila.createCell(11);
            celdaNumDiasTrabajados.setCellValue(numDiasTrabajados);

            /**
             * CASO ESPECIAL
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaPendienteTractor = fila.createCell(12);
            celdaPendienteTractor.setCellValue(pendienteTractor);

            /**
             * CASO ESPECIAL
             * Valores inicializados a 0 para ser cambiados posteriormente a mano
             */
            Cell celdaTrabajoRematado = fila.createCell(13);
            celdaTrabajoRematado.setCellValue(trabajoRematado);

            observaciones = listaApoyos.get(i).getObservaciones();
            Cell celdaObservaciones = fila.createCell(14);
            celdaObservaciones.setCellValue(observaciones);
        }
    }

    /**
     * PARTE EXCEL CONTROL CAPATACES
     */
    public void creacionExcelControlCapataces(){
        FileOutputStream fileMod2 = null;
        try{
            fileMod2 = new FileOutputStream("EXCELS FINALES/CONTROL CAPATACES NOMBRE.xlsx");
        } FileNotFoundException e) {
            System.out.println("Error al crear EXCEL DE CONTROL CAPATACES\n");
            System.exit(-1);
        }

        //Método que va a crear y rellenar mi excel de capataces
        introducirValoresCapataces();

        try {
            wb.write(fileMod2);
        } catch (IOException e) {
            System.out.println("Error al escribir EXCELL CAPATACES\n");
            System.exit(-1);
        }

        try {
            file.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar fichero");
            System.exit(-1);
        }

        try {
            fileMod2.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar EXCEL");
            System.exit(-1);
        }
    }

    public void introducirValoresCapataces(){
        /**
         * TODO Rellenado del excel capataces
         * INTRODUCIR TODOS LOS VALORES
         */
    }

}