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
                * APOYO
                */
               Integer id = fila.getRowNum() + 1;
               apoyoAnyadir.setIdApoyo(id);
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
               mapaMediciones.put(id, apoyoAnyadir);
            }
        }
        return mapaMediciones;
    }

    public void creacionExcelApoyosRealizados(){

    }

    public void creacionExcelControlCapataces(){

    }

}