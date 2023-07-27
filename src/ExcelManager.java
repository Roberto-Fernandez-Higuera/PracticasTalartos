/**
 * @author Roberto Fernández Higuera
 */
//package src;

import POJOS.Apoyo;
import org.apache.poi.ss.usermodel.*;
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

    //Arraylist con los valores de todos los campos del Excel apoyos
    private ArrayList<Apoyo> listaApoyos = new ArrayList<>();


    //MAPAS A UTILIZAR
    private HashMap<Integer, Apoyo> mapaApoyos = new HashMap<>();

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEES LAS PARTES DEL EXCEL
     */
    public ExcelManager(){
        try {
            this.file = new FileInputStream("src/main/resources/NOMBRE_DEL_EXCEL.xlsx");
            this.wb = new XSSFWorkbook(file);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel 1");
            System.exit(-1);
        }
        hojaIberdrola = wb.getSheetAt(0);
        this.mapaApoyos = leerDatosMedicionesPartes();
    }

    /**
     * Setteo de los valores que introducimos a cada campo de las mediciones
     * @return MAPA MEDICIONES
     */
    public HashMap leerDatosMedicionesPartes(){
        int numFilas = hojaIberdrola.getLastRowNum() - 21;

        for(int i = 4; i < numFilas; i++){
            Row fila = hojaIberdrola.getRow(i);
            if (fila != null && fila.getCell(0) != null){
               Apoyo apoyoAnyadir = new Apoyo();

               /**
                * ID FILA APOYO
                */
               Integer id = fila.getRowNum() - 1;
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
               mapaApoyos.put(id, apoyoAnyadir);
            }
        }
        return mapaApoyos;
    }

    /**
     * PARTE EXCEL APOYOS REALIZADOS
     */
    public void creacionExcelApoyosRealizados(){
        FileOutputStream fileMod = null;
        try{
            fileMod = new FileOutputStream("EXCELS_FINALES/EXCELS_APOYO/NOMBRE_EXCEL_QUE_QUEREMOS.xlsx");
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
        int contadorLongMant = 0;
        int contadorLongLimp = 0;
        int contadorLongApertura = 0;
        int contadorAnomalia = 0;
        int contadorLongitudCopa = 0;
        int contadorLongitudLimpiezaBase = 0;
        int contadorPodaCalle = 0;
        int contadorFijoSalida = 0;
        int contadorNumeroDiasTrabajados = 0;

        /**
         * Dar estilo de color y alineado para el título
         */
        CellStyle estiloCeldaTitulo = wb.createCellStyle();
        //COLOR
        estiloCeldaTitulo.setFillForegroundColor(Indexed.Colors.GREEN.getIndex());
        estiloCeldaTitulo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //NEGRITA
        Font font = wb.createFont();
        font.setBold(true);
        estiloCeldaTitulo.setFont(font);
        //ALINEADO HORIZONTAL Y VERTICAL
        estiloCeldaTitulo.setAlignment(HorizontalAlignment.CENTER);
        estiloCeldaTitulo.setVerticalAlignment(VerticalAlignment.CENTER);
        //BORDE DE LA CELDA EN NEGRITA
        estiloCeldaTitulo.setBorderTop(BorderStyle.THIN);
        estiloCeldaTitulo.setBorderBottom(BorderStyle.THIN);
        estiloCeldaTitulo.setBorderLeft(BorderStyle.THIN);
        estiloCeldaTitulo.setBorderRight(BorderStyle.THIN);

        /**
         * Dar estilo de negrita y alineado para celdas con información
         */
        CellStyle estiloCeldaInfo = wbCapataces.createCellStyle();
        //NEGRITA
        Font font = wbCapataces.createFont();
        font.setBold(true);
        estiloCeldaInfo.setFont(font);
        //ALINEADO HORIZONTAL Y VERTICAL
        estiloCeldaInfo.setAlignment(HorizontalAlignment.CENTER);
        estiloCeldaInfo.setVerticalAlignment(VerticalAlignment.CENTER);
        //BORDE DE LA CELDA EN NEGRITA
        estiloCeldaInfo.setBorderTop(BorderStyle.THIN);
        estiloCeldaInfo.setBorderBottom(BorderStyle.THIN);
        estiloCeldaInfo.setBorderLeft(BorderStyle.THIN);
        estiloCeldaInfo.setBorderRight(BorderStyle.THIN);

        for (int i = 0; i < listaApoyos.size() + 2; i++){
            Row fila = hoja.createRow(i);

            if (i == 0) {
                Cell celdaTitulo = fila.createCell(0);
                celdaTitulo.setCellValue("Código y nombre de la línea que deseas.");
            } else if (i == 1){

               Cell celdaColumnaApoyo = fila.createCell(0);
               celdaColumnaApoyo.setCellValue("APOYO");
               celdaColumnaApoyo.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaLongitudMantenimineto = fila.createCell(1);
               celdaColumnaLongitudMantenimineto.setCellValue("LONG\nMANT");
               celdaColumnaLongitudMantenimineto.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaLongitudLimpieza = fila.createCell(2);
               celdaColumnaLongitudLimpieza.setCellValue("LONG\nLIMPIEZA");
               celdaColumnaLongitudLimpieza.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaLongitudApertura = fila.createCell(3);
               celdaColumnaLongitudApertura.setCellValue("LONG\nAPERTURA");
               celdaColumnaLongitudApertura.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaAnomalia = fila.createCell(4);
               celdaColumnaAnomalia.setCellValue("ANOMALIA");
               celdaColumnaAnomalia.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaLongitudCopa = fila.createCell(5);
               celdaColumnaLongitudCopa.setCellValue("LONGITUD\nCOPA");
               celdaColumnaLongitudCopa.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaLimpiezaBase = fila.createCell(6);
               celdaColumnaLimpiezaBase.setCellValue("LIMPIEZA\nBASE\nAPOYOS");
               celdaColumnaLimpiezaBase.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaPodaCalle = fila.createCell(7);
               celdaColumnaPodaCalle.setCellValue("PODA\nCALLE");
               celdaColumnaPodaCalle.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaFijoSalida = fila.createCell(8);
               celdaColumnaFijoSalida.setCellValue("FIJO\nSALIDA");
               celdaColumnaFijoSalida.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaDia = fila.createCell(9);
               celdaColumnaDia.setCellValue("FECHA");
               celdaColumnaDia.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaCapataz = fila.createCell(10);
               celdaColumnaCapataz.setCellValue("CAPATAZ");
               celdaColumnaCapataz.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaCapataz = fila.createCell(11);
               celdaColumnaCapataz.setCellValue("Nº DIAS\nTRABAJADOS");
               celdaColumnaCapataz.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaTractor = fila.createCell(12);
               celdaColumnaTractor.setCellValue("PENDIENTE\nTRACTOR");
               celdaColumnaTractor.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaTrabajoRematado = fila.createCell(13);
               celdaColumnaTrabajoRematado.setCellValue("TRABAJO\nREMATADO");
               celdaColumnaTrabajoRematado.setCellStyle(estiloCeldaTitulo);

               Cell celdaColumnaObservaciones = fila.createCell(14);
               celdaColumnaObservaciones.setCellValue("OBSERVACIONES");
               celdaColumnaObservaciones.setCellStyle(estiloCeldaTitulo);

            } else {

                numApoyo = listaApoyos.get(i).getNumApoyo();
                Cell celdaNumApoyo = fila.createCell(0);
                celdaNumApoyo.setCellValue(numApoyo);
                celdaNumApoyo.setCellStyle(estiloCeldaInfo);

                longitudMantenimineto = listaApoyos.get(i).getLongitudMantenimineto();
                contadorLongMant += longitudMantenimineto;
                Cell celdaLongitudMantenimiento = fila.createCell(1);
                celdaLongitudMantenimiento.setCellValue(longitudMantenimineto);
                celdaLongitudMantenimiento.setCellStyle(estiloCeldaInfo);

                longitudLimpieza = listaApoyos.get(i).getLongitudLimpieza();
                contadorLongLimp += longitudLimpieza;
                Cell celdaLongitudLimpeza = fila.createCell(2);
                celdaLongitudLimpeza.setCellValue(longitudLimpieza);
                celdaLongitudLimpeza.setCellStyle(estiloCeldaInfo);

                longitudApertura = listaApoyos.get(i).getLongitudApertura();
                contadorLongApertura += longitudApertura;
                Cell celdaLongitudApertura = fila.createCell(3);
                celdaLongitudApertura.setCellValue(longitudApertura);
                celdaLongitudApertura.setCellStyle(estiloCeldaInfo);

                anomaliaVegetacion = listaApoyos.get(i).getnumAnomalia();
                contadorAnomalia += anomaliaVegetacion;
                Cell celdaAnomaliaVegetacion = fila.createCell(4);
                celdaAnomaliaVegetacion.setCellValue(anomaliaVegetacion);
                celdaAnomaliaVegetacion.setCellStyle(estiloCeldaInfo);

                /**
                 * TODO Hablar con Inés sobre los campos que no están (CASOS ESPECIALES)
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaLongitudCopa = fila.createCell(5);
                contadorLongitudCopa += longitudCopa;
                celdaLongitudCopa.setCellValue(longitudCopa);
                celdaLongitudCopa.setCellStyle(estiloCeldaInfo);

                limpiezaBase = listaApoyos.get(i).getLimpiezaBase();
                contadorLongitudLimpiezaBase += limpiezaBase;
                Cell celdaLimpiezaBase = fila.createCell(6);
                celdaLimpiezaBase.setCellValue(limpiezaBase);
                celdaLimpiezaBase.setCellStyle(estiloCeldaInfo);

                /**
                 * CASO ESPECIAL
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaPodaCalle = fila.createCell(7);
                contadorPodaCalle += podaCalle;
                celdaPodaCalle.setCellValue(podaCalle);
                celdaPodaCalle.setCellStyle(estiloCeldaInfo);

                /**
                 * CASO ESPECIAL
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaFijoSalida = fila.createCell(8);
                contadorFijoSalida += fijoSalida;
                celdaFijoSalida.setCellValue(fijoSalida);
                celdaFijoSalida.setCellStyle(estiloCeldaInfo);

                dia = listaApoyos.get(i).getDia();
                Cell celdaDia = fila.createCell(9);
                celdaDia.setCellValue(dia);
                celdaDia.setCellStyle(estiloCeldaInfo);

                capataz = listaApoyos.get(i).getCapataz();
                Cell celdaCapataz = fila.createCell(10);
                celdaCapataz.setCellValue(capataz);
                celdaCapataz.setCellStyle(estiloCeldaInfo);

                /**
                 * CASO ESPECIAL
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaNumDiasTrabajados = fila.createCell(11);
                contadorNumeroDiasTrabajados += numDiasTrabajados
                celdaNumDiasTrabajados.setCellValue(numDiasTrabajados);
                celdaNumDiasTrabajados.setCellStyle(estiloCeldaInfo);

                /**
                 * CASO ESPECIAL
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaPendienteTractor = fila.createCell(12);
                celdaPendienteTractor.setCellValue(pendienteTractor);
                celdaPendienteTractor.setCellStyle(estiloCeldaInfo);

                /**
                 * CASO ESPECIAL
                 * Valores inicializados a 0 para ser cambiados posteriormente a mano
                 */
                Cell celdaTrabajoRematado = fila.createCell(13);
                celdaTrabajoRematado.setCellValue(trabajoRematado);
                celdaTrabajoRematado.setCellStyle(estiloCeldaInfo);

                observaciones = listaApoyos.get(i).getObservaciones();
                Cell celdaObservaciones = fila.createCell(14);
                celdaObservaciones.setCellValue(observaciones);
                celdaObservaciones.setCellStyle(estiloCeldaInfo);

            }
        }
        /**
         * CELDAS DE OPERACIONES FINALES
         */
        Row filaSumas = hoja.createRow(listaApoyos.size()+2);

        Cell celdaColumnaSumaTotalApoyos = filaSumas.createCell(0);
        int totalApoyos = listaApoyos.size();
        celdaColumnaSumaTotalApoyos.setCellValue(totalApoyos);
        celdaColumnaSumaTotalApoyos.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudMantenimiento = filaSumas.createCell(1);
        celdaColumnaSumaTotalLongitudMantenimiento.setCellValue(contadorLongMant);
        celdaColumnaSumaTotalLongitudMantenimiento.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudLimpieza = filaSumas.createCell(2);
        celdaColumnaSumaTotalLongitudLimpieza.setCellValue(contadorLongLimp);
        celdaColumnaSumaTotalLongitudLimpieza.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudApertura = filaSumas.createCell(3);
        celdaColumnaSumaTotalLongitudApertura.setCellValue(contadorLongApertura);
        celdaColumnaSumaTotalLongitudApertura.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalAnomalia = filaSumas.createCell(4);
        celdaColumnaSumaTotalAnomalia.setCellValue(contadorAnomalia);
        celdaColumnaSumaTotalAnomalia.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudCopa = filaSumas.createCell(5);
        celdaColumnaSumaTotalLongitudCopa.setCellValue(contadorLongitudCopa);
        celdaColumnaSumaTotalLongitudCopa.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLimpiezaBase = filaSumas.createCell(6);
        celdaColumnaSumaTotalLimpiezaBase.setCellValue(contadorLongitudLimpiezaBase);
        celdaColumnaSumaTotalLimpiezaBase.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalPodaCalle = filaSumas.createCell(7);
        celdaColumnaSumaTotalPodaCalle.setCellValue(contadorPodaCalle);
        celdaColumnaSumaTotalPodaCalle.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalNumDiasTrabaj = filaSumas.createCell(10);
        celdaColumnaSumaTotalNumDiasTrabaj.setCellValue(contadorNumeroDiasTrabajados);
        celdaColumnaSumaTotalNumDiasTrabaj.setCellStyle(estiloCeldaTitulo);

        /**
         * CELDAS OPERACIONES FINALES CON RESPECTIVAS DIVISIONES
         */

        Row filaSumasDivisiones = hoja.createRow(listaApoyos.size()+3);

        Cell celdaColumnaSumaTotalApoyosDivision = filaSumasDivisiones.createCell(0);
        celdaColumnaSumaTotalApoyosDivision.setCellValue(totalApoyos);
        celdaColumnaSumaTotalApoyosDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudMantenimientoDivision = filaSumasDivisiones.createCell(1);
        celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellValue(contadorLongMant/1000);
        celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudLimpiezaDivision = filaSumasDivisiones.createCell(2);
        celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellValue(contadorLongLimp);
        celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudAperturaDivision = filaSumasDivisiones.createCell(3);
        celdaColumnaSumaTotalLongitudAperturaDivision.setCellValue(contadorLongApertura/1000);
        celdaColumnaSumaTotalLongitudAperturaDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalAnomaliaDivision = filaSumasDivisiones.createCell(4);
        celdaColumnaSumaTotalAnomaliaDivision.setCellValue(contadorAnomalia);
        celdaColumnaSumaTotalAnomaliaDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLongitudCopaDivision = filaSumasDivisiones.createCell(5);
        celdaColumnaSumaTotalLongitudCopaDivision.setCellValue(contadorLongitudCopa/1000);
        celdaColumnaSumaTotalLongitudCopaDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalLimpiezaBaseDivision = filaSumasDivisiones.createCell(6);
        celdaColumnaSumaTotalLimpiezaBaseDivision.setCellValue(contadorLongitudLimpiezaBase);
        celdaColumnaSumaTotalLimpiezaBaseDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalPodaCalleDivision = filaSumasDivisiones.createCell(7);
        celdaColumnaSumaTotalPodaCalleDivision.setCellValue(contadorPodaCalle);
        celdaColumnaSumaTotalPodaCalleDivision.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaSumaTotalNumDiasTrabajDivision = filaSumasDivisiones.createCell(10);
        celdaColumnaSumaTotalNumDiasTrabajDivision.setCellValue(contadorNumeroDiasTrabajados);
        celdaColumnaSumaTotalNumDiasTrabajDivision.setCellStyle(estiloCeldaTitulo);

    }
}