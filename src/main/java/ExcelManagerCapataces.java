/**
 * @author Roberto Fernández Higuera
 */

import POJOS.Apoyo;
import POJOS.Capataz;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;


public class ExcelManagerCapataces {

    private static FileInputStream fileCapataces;
    private static XSSFWorkbook wbCapataces;
    private XSSFSheet hojaApoyos;

    //Arraylist con los valores de todos los campos del Excel capataces
    private ArrayList<Capataz> listaCapataces = new ArrayList<>();
    private ArrayList<String> listaCapatacesLinea = new ArrayList<>();

    //MAPAS A UTILIZAR
    private HashMap<Integer, Capataz> mapaCapataces = new HashMap<>();

    private String nombreHoja;
    private String nombreExcel;

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEER LAS PARTES DEL EXCEL
     */

    public ExcelManagerCapataces() {

        Scanner scanner = new Scanner(System.in);
        System.out.print("Introduce el nombre del Excel de APOYOS con el que quieres trabajar:\n");
        nombreExcel = scanner.nextLine();

        String rutaExcel = "EXCELS_FINALES/EXCELS_APOYO/"+nombreExcel+".xlsx";
        try {
            this.fileCapataces = new FileInputStream(rutaExcel);
            this.wbCapataces = new XSSFWorkbook(fileCapataces);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel: "+nombreExcel);
            System.exit(-1);
        }

        Scanner sc = new Scanner(System.in);
        System.out.print("Introduce el nombre de la hoja del Excel de APOYOS con la que quieres trabajar (línea de la que quieres que se cree el control de capataces):\n");
        nombreHoja = sc.nextLine();

        hojaApoyos = wbCapataces.getSheet(nombreHoja);
        this.mapaCapataces = leerDatosCapataces();
    }

    /**
     * TODO NECESITO SABER DE DONDE SALE CADA VALOR Y SABER SOBRE QUÉ HOJA EXCEL LOS TOMO
     *
     * @return MAPA CAPATACES
     */
    private HashMap leerDatosCapataces() {
        int numFilas = hojaApoyos.getLastRowNum() - 1;

        for (int i = 2; i < numFilas; i++) {
            Row fila = hojaApoyos.getRow(i);
            if (fila != null && fila.getCell(0) != null) {
                Capataz capatazAnyadir = new Capataz();

                /**
                 * ID CAPATAZ
                 */
                Integer id = fila.getRowNum();
                capatazAnyadir.setIdCapataz(id);

                /**
                 * NUM APOYOS CAPATAZ
                 */
                capatazAnyadir.setNumApoyos(0);

                /**
                 * LONGITUD MANTENIMIENTO
                 */
                capatazAnyadir.setLongMantenimiento(fila.getCell(1).getNumericCellValue());

                /**
                 * LONGITUD LIMPIEZA
                 */
                capatazAnyadir.setLongitudLimpieza(fila.getCell(2).getNumericCellValue());

                /**
                 * LONGIUTD APERTURA
                 */
                capatazAnyadir.setLongApertura(fila.getCell(3).getNumericCellValue());

                /**
                 * NUM ANOMALIA
                 */
                capatazAnyadir.setAnomalia(fila.getCell(4).getNumericCellValue());

                /**
                 * LIMPIEZA BASE
                 */
                capatazAnyadir.setLimpiezaBase(fila.getCell(5).getNumericCellValue());

                /**
                 * PODA CALLE
                 */
                capatazAnyadir.setPodaCalle(fila.getCell(6).getNumericCellValue());

                /**
                 * FIJO SALIDA
                 */
                capatazAnyadir.setFijoSalida(fila.getCell(7).getNumericCellValue());

                /**
                 * DÍA APOYO
                 */
                capatazAnyadir.setDia(fila.getCell(7).getNumericCellValue());

                /**
                 * NOMBRE APOYO
                 */
                capatazAnyadir.setNombreApoyo(fila.getCell(9).getStringCellValue());

                /**
                 * N DIAS TRABAJADOS
                 */
                capatazAnyadir.setNumDiasTrabajados(0);

                /**
                 * PENDIENTE TRACTOR
                 */
                capatazAnyadir.setPendienteTractor(fila.getCell(11).getStringCellValue());

                /**
                 * TRABAJO REMATADO
                 */
                capatazAnyadir.setTrabajoRematado(fila.getCell(12).getStringCellValue());

                /**
                 * OBSERVACIONES
                 */
                capatazAnyadir.setObservaciones(fila.getCell(13).getStringCellValue());


                listaCapataces.add(capatazAnyadir);
                mapaCapataces.put(id, capatazAnyadir);
            }
        }
        return mapaCapataces;
    }

    public void creacionExcelControlCapataces(String nombreExcel, String zona, String codLinea) {
        String nombreArchivoSalida = "EXCELS_FINALES/EXCELS_CAPATACES/"+nombreExcel+".xlsx";
        File archivoSalida = new File(nombreArchivoSalida);

        FileOutputStream fileModCapataces = null;

        if (archivoSalida.exists()) {
            // El archivo de salida ya existe, abre el libro existente
            try {
                FileInputStream file = new FileInputStream(nombreArchivoSalida);
                wbCapataces = new XSSFWorkbook(file);
                file.close();
            } catch (IOException e) {
                System.out.println("Error al abrir el archivo existente: " + e.getMessage());
                System.exit(-1);
            }
        } else {
            // El archivo de salida no existe, crea uno nuevo con el nombre proporcionado
            wbCapataces = new XSSFWorkbook();
        }

        /**
         * Comprobación de si las hojas ya existen en el excel con todos los trabajadores de la línea
         */

        listaCapatacesLinea = capatacesLinea(listaCapataces);

        for (int i = 0; i < listaCapatacesLinea.size() ;i++) {
            Sheet hoja = wbCapataces.getSheet(listaCapatacesLinea.get(i));
            if (hoja == null) {
                hoja = wbCapataces.createSheet(listaCapatacesLinea.get(i));
            }

            //Método que va a crear y rellenar mi excel de apoyos
            introducirValoresCapataz(hoja, zona, codLinea, listaCapatacesLinea.get(i));
        }
        try {
            fileModCapataces = new FileOutputStream(nombreArchivoSalida);
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

    public static ArrayList<String> capatacesLinea(ArrayList<Capataz> listaCapataces){
         ArrayList<String> listaCapatacesLinea = new ArrayList<>();

        for (Capataz capataz : listaCapataces){
            String nombreCapataz = capataz.getNombreApoyo();

            if (!listaCapatacesLinea.contains(nombreCapataz)) {
                listaCapatacesLinea.add(nombreCapataz);
            }
        }
        return listaCapatacesLinea;
    }

    public void introducirValoresCapataz(Sheet hoja, String zona, String codLinea, String nombreHoja) {
        double fecha = 0;
        LocalDate diaLocalDate = null;
        Date fechaDate = null;
        double numApoyos = 0;
        double fijoSalida = 0;
        double longMantenimiento = 0;
        double anomalia = 0;
        double longApertura = 0;
        String talasFuera = "";
        double limpiezaBase = 0;
        double km = 0;
        double importeMedios = 0;
        double importeCoeficiente = 0;
        String observaciones = "";
        int contadorApoyos = 0;
        int contadorLongMantenimiento = 0;
        int contadorAnomalia = 0;
        int contadorLongApertura = 0;
        int contadorTalasFuera = 0;
        int contadorLimpiezaBase = 0;
        int contadorKm = 0;
        int contadorImporteMedio = 0;
        int contadorImporteCoeficiente = 0;
        // ***importeCoeficiente/7***
        int importeCoeficienteSemanal = 0;


        /**
         * Dar estilo de color y alineado para el título
         */
        CellStyle estiloCeldaTitulo = wbCapataces.createCellStyle();
        //COLOR
        estiloCeldaTitulo.setFillForegroundColor(IndexedColors.LIME.getIndex());
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
        //PERMITE QUE EL TEXTO SE ENVUELVA
        estiloCeldaTitulo.setWrapText(true);

        /**
         * Dar estilo de negrita y alineado para celdas con información
         */
        CellStyle estiloCeldaInfo = wbCapataces.createCellStyle();
        //NEGRITA
        estiloCeldaInfo.setFont(font);
        //ALINEADO HORIZONTAL Y VERTICAL
        estiloCeldaInfo.setAlignment(HorizontalAlignment.CENTER);
        estiloCeldaInfo.setVerticalAlignment(VerticalAlignment.CENTER);
        //BORDE DE LA CELDA EN NEGRITA
        estiloCeldaInfo.setBorderTop(BorderStyle.THIN);
        estiloCeldaInfo.setBorderBottom(BorderStyle.THIN);
        estiloCeldaInfo.setBorderLeft(BorderStyle.THIN);
        estiloCeldaInfo.setBorderRight(BorderStyle.THIN);

        /**
         * Estilo para las fechas con fecha
         */
        //DAR FORMATO
        CellStyle estiloFecha = wbCapataces.createCellStyle();
        DataFormat formatoFecha = wbCapataces.createDataFormat();
        estiloFecha.setDataFormat(formatoFecha.getFormat("dd/mm/yyyy"));
        //NEGRITA
        estiloFecha.setFont(font);
        //ALINEADO HORIZONTAL Y VERTICAL
        estiloFecha.setAlignment(HorizontalAlignment.CENTER);
        estiloFecha.setVerticalAlignment(VerticalAlignment.CENTER);
        //BORDE DE LA CELDA EN NEGRITA
        estiloFecha.setBorderTop(BorderStyle.THIN);
        estiloFecha.setBorderBottom(BorderStyle.THIN);
        estiloFecha.setBorderLeft(BorderStyle.THIN);
        estiloFecha.setBorderRight(BorderStyle.THIN);


        for (int i = 0; i < listaCapataces.size() + 1; i++) {
            Row fila = hoja.createRow(i);

            if (i == 0) {

                Cell celdaColumnaDia = fila.createCell(0);
                celdaColumnaDia.setCellValue("DÍA");
                celdaColumnaDia.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaApoyos = fila.createCell(1);
                celdaColumnaApoyos.setCellValue("APOYOS");
                celdaColumnaApoyos.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaFijoSalida = fila.createCell(2);
                celdaColumnaFijoSalida.setCellValue("FIJO SALIDA");
                celdaColumnaFijoSalida.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLongitudMantenimineto = fila.createCell(3);
                celdaColumnaLongitudMantenimineto.setCellValue("LONG MANT");
                celdaColumnaLongitudMantenimineto.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaAnomalia = fila.createCell(4);
                celdaColumnaAnomalia.setCellValue("ANOMALIAS");
                celdaColumnaAnomalia.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLongitudApertura = fila.createCell(5);
                celdaColumnaLongitudApertura.setCellValue("LONGITUD APERTURA");
                celdaColumnaLongitudApertura.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaTalasFueraCalle = fila.createCell(6);
                celdaColumnaTalasFueraCalle.setCellValue("TALAS FUERA DE LA CALLE");
                celdaColumnaTalasFueraCalle.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaLimpiezaBase = fila.createCell(7);
                celdaColumnaLimpiezaBase.setCellValue("LIMPIEZA BASE");
                celdaColumnaLimpiezaBase.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaIdentZonasNuecas = fila.createCell(8);
                celdaColumnaIdentZonasNuecas.setCellValue("KM");
                celdaColumnaIdentZonasNuecas.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaImporte = fila.createCell(9);
                celdaColumnaImporte.setCellValue("IMPORTE MEDIOS");
                celdaColumnaImporte.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaImporteCoeficiente = fila.createCell(10);
                celdaColumnaImporteCoeficiente.setCellValue("IMPORTE COEFICIENTE");
                celdaColumnaImporteCoeficiente.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaZONA = fila.createCell(11);
                celdaColumnaZONA.setCellValue("ZONA");
                celdaColumnaZONA.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaOBSERVACIONES = fila.createCell(12);
                celdaColumnaOBSERVACIONES.setCellValue("OBSERVACIONES");
                celdaColumnaOBSERVACIONES.setCellStyle(estiloCeldaTitulo);

                Cell celdaColumnaCODLINEA = fila.createCell(13);
                celdaColumnaCODLINEA.setCellValue("COD LINEA");
                celdaColumnaCODLINEA.setCellStyle(estiloCeldaTitulo);

            } else {

                fecha = listaCapataces.get(i-1).getDia();
                diaLocalDate =  LocalDate.of(1899, 12, 30).plusDays((long) fecha);
                fechaDate = Date.from(diaLocalDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
                Cell celdaDia = fila.createCell(0);
                celdaDia.setCellValue(fechaDate);
                celdaDia.setCellStyle(estiloFecha);

                numApoyos = listaCapataces.get(i-1).getNumApoyos();
                contadorApoyos += numApoyos;
                Cell celdaNumApoyos = fila.createCell(1);
                celdaNumApoyos.setCellValue(numApoyos);
                celdaNumApoyos.setCellStyle(estiloCeldaInfo);

                fijoSalida = listaCapataces.get(i-1).getFijoSalida();
                Cell celdaFijoSalida = fila.createCell(2);
                celdaFijoSalida.setCellValue(fijoSalida);
                celdaFijoSalida.setCellStyle(estiloCeldaInfo);

                longMantenimiento = listaCapataces.get(i-1).getLongMantenimiento();
                contadorLongMantenimiento += longMantenimiento;
                Cell celdaLongMantenimiento = fila.createCell(3);
                celdaLongMantenimiento.setCellValue(longMantenimiento);
                celdaLongMantenimiento.setCellStyle(estiloCeldaInfo);

                anomalia = listaCapataces.get(i-1).getAnomalia();
                contadorAnomalia += anomalia;
                Cell celdaAnomalia = fila.createCell(4);
                celdaAnomalia.setCellValue(anomalia);
                celdaAnomalia.setCellStyle(estiloCeldaInfo);

                longApertura = listaCapataces.get(i-1).getLongApertura();
                contadorLongApertura += longApertura;
                Cell celdaLongApertura = fila.createCell(5);
                celdaLongApertura.setCellValue(longApertura);
                celdaLongApertura.setCellStyle(estiloCeldaInfo);

                Cell celdaTalasFuera = fila.createCell(6);
                celdaTalasFuera.setCellValue(talasFuera);
                celdaTalasFuera.setCellStyle(estiloCeldaInfo);

                limpiezaBase = listaCapataces.get(i-1).getLimpiezaBase();
                contadorLimpiezaBase += limpiezaBase;
                Cell celdaLimpiezaBase = fila.createCell(7);
                celdaLimpiezaBase.setCellValue(limpiezaBase);
                celdaLimpiezaBase.setCellStyle(estiloCeldaInfo);

                /**
                 * PREGUNTAR FUNCIONAMIENTO
                 */
                km = listaCapataces.get(i-1).getKm();
                contadorKm += km;
                Cell celdaKm = fila.createCell(8);
                celdaKm.setCellValue(km);
                celdaKm.setCellStyle(estiloCeldaInfo);

                /**
                 * PREGUNTAR FUNCIONAMIENTO
                 */
                importeMedios = listaCapataces.get(i-1).getImporteMedios();
                contadorImporteMedio += importeMedios;
                Cell celdaImporteMedios = fila.createCell(9);
                celdaImporteMedios.setCellValue(importeMedios);
                celdaImporteMedios.setCellStyle(estiloCeldaInfo);

                /**
                 * PREGUNTAR FUNCIONAMIENTO
                 */
                importeCoeficiente = listaCapataces.get(i-1).getImporteCoeficiente();
                contadorImporteCoeficiente += importeCoeficiente;
                Cell celdaImporteCoeficiente = fila.createCell(10);
                celdaImporteCoeficiente.setCellValue(importeCoeficiente);
                celdaImporteCoeficiente.setCellStyle(estiloCeldaInfo);

                Cell celdaZona = fila.createCell(11);
                celdaZona.setCellValue(zona);
                celdaZona.setCellStyle(estiloCeldaInfo);

                observaciones = listaCapataces.get(i-1).getObservaciones();
                Cell celdaObservaciones = fila.createCell(12);
                celdaObservaciones.setCellValue(observaciones);
                celdaObservaciones.setCellStyle(estiloCeldaInfo);

                Cell celdaCodLinea = fila.createCell(13);
                celdaCodLinea.setCellValue(codLinea);
                celdaCodLinea.setCellStyle(estiloCeldaInfo);
            }
        }

        /**
         * CELDAS DE OPERACIONES FINALES, PREGUNTAR A INÉS SI SE NECESITAN MÁS
         */
        Row filaSumas = hoja.createRow(listaCapataces.size() + 1);

        Cell celdaColumnaTotalApoyos = filaSumas.createCell(1);
        celdaColumnaTotalApoyos.setCellValue(contadorApoyos);
        celdaColumnaTotalApoyos.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalLongMantenimiento = filaSumas.createCell(3);
        celdaColumnaTotalLongMantenimiento.setCellValue(contadorLongMantenimiento);
        celdaColumnaTotalLongMantenimiento.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalAnomalia = filaSumas.createCell(4);
        celdaColumnaTotalAnomalia.setCellValue(contadorAnomalia);
        celdaColumnaTotalAnomalia.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalLongApertura = filaSumas.createCell(5);
        celdaColumnaTotalLongApertura.setCellValue(contadorLongApertura);
        celdaColumnaTotalLongApertura.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalTalasFuera = filaSumas.createCell(6);
        celdaColumnaTotalTalasFuera.setCellValue(contadorTalasFuera);
        celdaColumnaTotalTalasFuera.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalLimpiezaBase = filaSumas.createCell(7);
        celdaColumnaTotalLimpiezaBase.setCellValue(contadorLimpiezaBase);
        celdaColumnaTotalLimpiezaBase.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalKm = filaSumas.createCell(8);
        celdaColumnaTotalKm.setCellValue(contadorKm);
        celdaColumnaTotalKm.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalImporteMedio = filaSumas.createCell(9);
        celdaColumnaTotalImporteMedio.setCellValue(contadorImporteMedio);
        celdaColumnaTotalImporteMedio.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalImporteCoeficiente = filaSumas.createCell(10);
        celdaColumnaTotalImporteCoeficiente.setCellValue(contadorImporteCoeficiente);
        celdaColumnaTotalImporteCoeficiente.setCellStyle(estiloCeldaTitulo);

        /**
         * PREGUNTAR A INÉS SOBRE ESTO
         */

        Row filaImporteCoeficienteSemanal = hoja.createRow(listaCapataces.size() + 3);

        Cell celdaColumnaTextoParaCoeficienteSemanala = filaImporteCoeficienteSemanal.createCell(11);
        celdaColumnaTextoParaCoeficienteSemanala.setCellValue("IMPORTE SEMANAL:");
        celdaColumnaTextoParaCoeficienteSemanala.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTotalImporteCoeficienteSemanal = filaImporteCoeficienteSemanal.createCell(12);
        importeCoeficienteSemanal = contadorImporteCoeficiente / 7;
        celdaColumnaTotalImporteCoeficienteSemanal.setCellValue(importeCoeficienteSemanal);
        celdaColumnaTotalImporteCoeficienteSemanal.setCellStyle(estiloCeldaTitulo);
    }
}
