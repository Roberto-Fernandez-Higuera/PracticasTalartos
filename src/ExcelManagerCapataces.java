package src;

import POJOS.Capataz;
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


public class ExcelManagerCapataces() {

    private static FileInputStream fileCapataces;
    private static XSSFWorkbook wbCapataces;
    private XSSFSheet hojaApoyos;

    //Arraylist con los valores de todos los campos del Excel capataces
    private ArrayList<Apoyo> listaCapataces = new ArrayList<>();

    //MAPAS A UTILIZAR
    private HashMap<Integer, Capataz> mapaCapataces = new HashMap<>();

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEES LAS PARTES DEL EXCEL
     */
    public ExcelManagerCapataces() {
        try {
            this.fileCapataces = new FileInputStream("EXCELS_FINALES/EXCELS_APOYO/NOMBRE_DEL_EXCEL_DE_APOYOS.xlsx");
            this.wbCapataces = new XSSFWorkbook(fileCapataces);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel 1");
            System.exit(-1);
        }
        hojaApoyos = wbCapataces.getSheetAt(0);
        this.mapaCapataces = leerDatosCapataces();
    }

    /**
     * TODO NECESITO SABER DE DONDE SALE CADA VALOR Y SABER SOBRE QUÉ HOJA EXCEL LOS TOMO
     * @return MAPA CAPATACES
     */
    public HashMap leerDatosCapataces(){
        int numFilas = hojaApoyos.getLastRowNum() - 1;

        for(int i = 2; i < numFilas; i++){
            Row fila = hojaApoyos.getRow(i);
            if (fila != null && fila.getCell(0) != null){
                Capataz capatazAnyadir = new Apoyo();

                /**
                 * ID DÍA CAPATAZ
                 */
                Integer id = fila.getRowNum();
                capatazAnyadir.setIdCapataz(id);

                /**
                 * DÍA APOYO
                 */
                capatazAnyadir.setDia();

                /**
                 * NUM APOYOS CAPATAZ
                 */
                capatazAnyadir.setNumApoyos();

                /**
                 * FIJO SALIDA
                 */
                capatazAnyadir.setFijoSalida();

                /**
                 * LONGITUD MANTENIMIENTO
                 */
                capatazAnyadir.setLongitudMantenimiento();

                /**
                 * NUM ANOMALIA
                 */
                capatazAnyadir.setNumAnomalia();

                /**
                 * LONGIUTD APERTURA
                 */
                capatazAnyadir.setLongitudApertura();

                /**
                 * TALAS FUERA
                 */
                capatazAnyadir.setTalasFuera();

                /**
                 * LONGITUD COPA
                 */
                capatazAnyadir.setLongitudCopa();

                /**
                 * LIMPIEZA BASE
                 */
                capatazAnyadir.setLimpiezaBase();

                /**
                 * KM
                 */
                capatazAnyadir.setKm();

                /**
                 * IMPORTE MEDIO
                 */
                capatazAnyadir.setImporteMedios();

                /**
                 * IMPORTE COEFICIENTE
                 */
                capatazAnyadir.setImporteCoeficiente();

                /**
                 * ZONA
                 */
                capatazAnyadir.setZona();

                /**
                 * OBSERVACIONES
                 */
                capatazAnyadir.setObservaciones();

                /**
                 * CÓDIGO LINEA
                 */
                capatazAnyadir.setCodLinea();

                listaCapataces.add(capatazAnyadir);
                mapaCapataces.put(id, capatazAnyadir);
            }
        }
        return mapaCapataces;
    }

    public void creacionExcelControlCapataces(){
        FileOutputStream fileModCapataces = null;
        try{
            fileModCapataces = new FileOutputStream("EXCELS_FINALES/EXCELS_APOYO/NOMBRE_EXCEL_QUE_QUEREMOS.xlsx");
        } FileNotFoundException e) {
            System.out.println("Error al crear EXCEL DE CAPATACES\n");
            System.exit(-1);
        }

        //Método que va a crear y rellenar mi excel de apoyos
        introducirValoresCapataz();

        try {
            wbCapataces.write(fileModCapataces);
        } catch (IOException e) {
            System.out.println("Error al escribir EXCELL CAPATACES\n");
            System.exit(-1);
        }

        try {
            fileCapataces.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar fichero");
            System.exit(-1);
        }

        try {
            fileModCapataces.close();
        } catch (IOException e) {
            System.out.println("Error al cerrar EXCEL");
            System.exit(-1);
        }
    }

    public void introducirValoresCapataz(){
        /**
         * TODO Rellenado del excel capataces
         * INTRODUCIR TODOS LOS VALORES
         */
        Date fecha;
        int numApoyos = 0;
        int fijoSalida = 0;
        int longMantenimiento = 0;
        int anomalia = 0;
        int longApertura = 0;
        int talasFuera = 0;
        int longCopa = 0;
        int limpiezaBase = 0;
        int km = 0;
        int importeMedios = 0;
        int importeCoeficiente = 0;
        String zona = "";
        String observaciones = "";
        String codLinea = "";

        /**
         * Dar estilo de color y alineado para el título
         */
        CellStyle estiloCeldaTitulo = wbCapataces.createCellStyle();
        //COLOR
        estiloCeldaTitulo.setFillForegroundColor(Indexed.Colors.GREEN.getIndex());
        estiloCeldaTitulo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //NEGRITA
        Font font = wbCapataces.createFont();
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

        for (int i = 0; i < listaCapataces.size() + 1; i++){
            Row fila = hojaApoyos.createRow(i);

            if (i == 0){

                Cell celdaColumnaDia = fila.createCell(0);
                celdaColumnaDia.setCellValue("DÍA");
                celdaColumnaDia.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaApoyos = fila.createCell(1);
                celdaColumnaApoyos.setCellValue("APOYOS");
                celdaColumnaApoyos.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaFijoSalida = fila.createCell(2);
                celdaColumnaFijoSalida.setCellValue("FIJO\nSALIDA");
                celdaColumnaFijoSalida.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLongitudMantenimineto = fila.createCell(3);
                celdaColumnaLongitudMantenimineto.setCellValue("LONG\nMANT");
                celdaColumnaLongitudMantenimineto.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaAnomalia = fila.createCell(4);
                celdaColumnaAnomalia.setCellValue("ANOMALIAS");
                celdaColumnaAnomalia.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLongitudApertura = fila.createCell(5);
                celdaColumnaLongitudApertura.setCellValue("LONGITUD\nAPERTURA");
                celdaColumnaLongitudApertura.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaTalasFueraCalle = fila.createCell(6);
                celdaColumnaTalasFueraCalle.setCellValue("TALAS FUERA\nDE LA\nCALLE");
                celdaColumnaTalasFueraCalle.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLongitudCopa = fila.createCell(7);
                celdaColumnaLongitudCopa.setCellValue("LONGITUD\nCOPA");
                celdaColumnaLongitudCopa.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLimpiezaBase = fila.createCell(8);
                celdaColumnaLimpiezaBase.setCellValue("LIMPIEZA\nBASE");
                celdaColumnaLimpiezaBase.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaIdentZonasNuecas = fila.createCell(9);
                celdaColumnaIdentZonasNuecas.setCellValue("KM");
                celdaColumnaIdentZonasNuecas.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaImporte = fila.createCell(10);
                celdaColumnaImporte.setCellValue("IMPORTE MEDIOS");
                celdaColumnaImporte.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaImporteCoeficiente = fila.createCell(11);
                celdaColumnaImporteCoeficiente.setCellValue("IMPORTE\nCOEFICIENTE");
                celdaColumnaImporteCoeficiente.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaZONA = fila.createCell(12);
                celdaColumnaZONA.setCellValue("ZONA");
                celdaColumnaZONA.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaOBSERVACIONES = fila.createCell(13);
                celdaColumnaOBSERVACIONES.setCellValue("OBSERVACIONES");
                celdaColumnaOBSERVACIONES.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaCODLINEA = fila.createCell(14);
                celdaColumnaCODLINEA.setCellValue("COD LINEA");
                celdaColumnaCODLINEA.setCellStyle(estiloCeldaTitulo);

            } else {

                fecha = listaCapataces.get(i).getDia();
                Cell celdaFecha = fila.createCell(0);
                celdaFecha.setCellValue(fecha);
                celdaFecha.setCellStyle(estiloCeldaInfo);

                numApoyos = listaCapataces.get(i).getNumApoyos();
                Cell celdaNumApoyos = fila.createCell(1);
                celdaNumApoyos.setCellValue(numApoyos);
                celdaNumApoyos.setCellStyle(estiloCeldaInfo);

                fijoSalida = listaCapataces.get(i).getFijoSalida();
                Cell celdaFijoSalida = fila.createCell(2);
                celdaFijoSalida.setCellValue(fijoSalida);
                celdaFijoSalida.setCellStyle(estiloCeldaInfo);

                longMantenimiento = listaCapataces.get(i).getLongMantenimiento();
                Cell celdaLongMantenimiento = fila.createCell(3);
                celdaLongMantenimiento.setCellValue(longMantenimiento);
                celdaLongMantenimiento.setCellStyle(estiloCeldaInfo);

                anomalia = listaCapataces.get(i).getAnomalia();
                Cell celdaAnomalia = fila.createCell(4);
                celdaAnomalia.setCellValue(anomalia);
                celdaAnomalia.setCellStyle(estiloCeldaInfo);

                longApertura = listaCapataces.get(i).getLongApertura();
                Cell celdaLongApertura = fila.createCell(5);
                celdaLongApertura.setCellValue(longApertura);
                celdaLongApertura.setCellStyle(estiloCeldaInfo);

                talasFuera = listaCapataces.get(i).getTalasFuera();
                Cell celdaTalasFuera = fila.createCell(6);
                celdaTalasFuera.setCellValue(talasFuera);
                celdaTalasFuera.setCellStyle(estiloCeldaInfo);

                longCopa = listaCapataces.get(i).getLongitudCopa();
                Cell celdaLongCopa = fila.createCell(7);
                celdaLongCopa.setCellValue(longCopa);
                celdaLongCopa.setCellStyle(estiloCeldaInfo);

                limpiezaBase = listaCapataces.get(i).getLimpiezaBase();
                Cell celdaLimpiezaBase = fila.createCell(8);
                celdaLimpiezaBase.setCellValue(limpiezaBase);
                celdaLimpiezaBase.setCellStyle(estiloCeldaInfo);

                km = listaCapataces.get(i).getKm();
                Cell celdaKm = fila.createCell(9);
                celdaKm.setCellValue(km);
                celdaKm.setCellStyle(estiloCeldaInfo);

                importeMedios = listaCapataces.get(i).getImporteMedios();
                Cell celdaImporteMedios = fila.createCell(10);
                celdaImporteMedios.setCellValue(importeMedios);
                celdaImporteMedios.setCellStyle(estiloCeldaInfo);

                importeCoeficiente = listaCapataces.get(i).getImporteCoeficiente();
                Cell celdaImporteCoeficiente = fila.createCell(11);
                celdaImporteCoeficiente.setCellValue(importeCoeficiente);
                celdaImporteCoeficiente.setCellStyle(estiloCeldaInfo);

                zona = listaCapataces.get(i).getZona();
                Cell celdaZona = fila.createCell(12);
                celdaZona.setCellValue(zona);
                celdaZona.setCellStyle(estiloCeldaInfo);

                observaciones = listaCapataces.get(i).getObservaciones();
                Cell celdaObservaciones = fila.createCell(13);
                celdaObservaciones.setCellValue(observaciones);
                celdaObservaciones.setCellStlye(estiloCeldaInfo);

                codLinea = listaCapataces.get(i).getCodLinea();
                Cell celdaCodLinea = fila.createCell(14);
                celdaCodLinea.setCellValue(codLinea);
                celdaCodLinea.setCellStyle(estiloCeldaInfo);
            }
        }

        /**
         * CELDAS DE OPERACIONES FINALES, PREGUNTAR A INÉS SI SE NECESITAN MÁS
         */
        Row filaSumas = hojaApoyos.createRow(listaCapataces.size()+1);



    }
}