/**
 * @author Roberto Fernández Higuera
 */

import POJOS.Apoyo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

public class ExcelManager {

    private static FileInputStream file;
    private static XSSFWorkbook wb;
    private XSSFSheet hojaIberdrola;

    //Arraylist con los valores de todos los campos del Excel apoyos
    private ArrayList<Apoyo> listaApoyos = new ArrayList<>();

    //MAPAS A UTILIZAR
    private HashMap<Integer, Apoyo> mapaApoyos = new HashMap<>();
    private String nombreExcel;
    private String provincias;
    private String zona;
    static String linea;

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEER LAS PARTES DEL EXCEL
     */
    public ExcelManager() {

        Scanner scProvincias = new Scanner(System.in);

        System.out.print("Introduce el nombre de la carpeta de PROVINCIAS con la que quieres trabajar: \n");
        provincias = scProvincias.nextLine();

        Scanner scZona = new Scanner(System.in);

        System.out.print("Introduce el nombre de la zona de "+provincias+" con la que quieres trabajar: \n");
        zona = scZona.nextLine();

        Scanner scLinea = new Scanner(System.in);

        System.out.print("Introduce el nombre de la línea de "+zona+" con la que quieres trabajar: \n");
        linea = scLinea.nextLine();

        Scanner scanner = new Scanner(System.in);

        System.out.print("Introduce el nombre del Excel de MEDICIONES con el que quieres trabajar: \n");
        nombreExcel = scanner.nextLine();

        String rutaExcel = "src/main/2023/"+provincias+"/MEDICIONES PARTES/"+zona+"/"+linea+"/"+nombreExcel+".xlsx";
        try {
            this.file = new FileInputStream(rutaExcel);
            this.wb = new XSSFWorkbook(file);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel: "+nombreExcel);
            System.exit(-1);
        }
        hojaIberdrola = wb.getSheetAt(0);
        this.mapaApoyos = leerDatosMedicionesPartes();
    }

    /**
     * Setteo de los valores que introducimos a cada campo de las mediciones
     *
     * @return MAPA MEDICIONES
     */
    private HashMap leerDatosMedicionesPartes() {
        int filasTotales = hojaIberdrola.getLastRowNum();
        String valor = "";
        int numFilas = 0;

        //Bucle para recoger el numFilas que debemos de tomar para leer todos los apoyos
        for (int i = 0; i < filasTotales; i++) {
            Row fila = hojaIberdrola.getRow(i);
            if (fila != null) {
                Cell celda = fila.getCell(1); // Suponiendo que "TOTALES" está en la columna 1
                if (celda != null && celda.getCellType() == CellType.STRING) {
                    String valorCelda = celda.getStringCellValue();
                    if (valorCelda.equals("TOTALES")) {
                        numFilas = fila.getRowNum();
                        break; // Salir del bucle una vez que se encuentra "TOTALES"
                    }
                }
            }
        }

        for (int i = 4; i < numFilas; i++) {
            Row fila = hojaIberdrola.getRow(i);
            if (fila != null && fila.getCell(1) != null) {
                Apoyo apoyoAnyadir = new Apoyo();

                /**
                 * ID FILA APOYO
                 */
                Integer id = fila.getRowNum() - 3;
                apoyoAnyadir.setIdApoyo(id);

                /**
                 * NUM APOYO
                 */
                //En función del tipo de celda
                CellType cellType = fila.getCell(1).getCellType();
                switch (cellType) {
                    case STRING:
                        apoyoAnyadir.setNumApoyo(fila.getCell(1).getStringCellValue());
                        break;
                    case NUMERIC:
                        apoyoAnyadir.setNumApoyo(Integer.toString((int)fila.getCell(1).getNumericCellValue()));
                        break;
                }

                /**
                 * LONGITUD MANTENIMIENTO
                 */
                if (fila.getCell(2) == null) {
                    apoyoAnyadir.setLongitudMantenimiento(0);
                } else {
                    apoyoAnyadir.setLongitudMantenimiento(fila.getCell(2).getNumericCellValue());
                }

                /**
                 * LONGITUD LIMPIEZA
                 */
                if (fila.getCell(3) == null) {
                    apoyoAnyadir.setLongitudLimpieza(0);
                } else {
                    apoyoAnyadir.setLongitudLimpieza(fila.getCell(3).getNumericCellValue());
                }

                /**
                 * LONGITUD APERTURA
                 */
                if (fila.getCell(4) == null) {
                    apoyoAnyadir.setLongitudApertura(0);
                } else {
                    apoyoAnyadir.setLongitudApertura(fila.getCell(4).getNumericCellValue());
                }

                /**
                 * ANOMALÍA VEGETACIÓN
                 */
                if (fila.getCell(5) == null) {
                    apoyoAnyadir.setNumAnomalia(0);
                } else {
                    apoyoAnyadir.setNumAnomalia(fila.getCell(5).getNumericCellValue());
                }

                /**
                 * LIMPIEZA BASE
                 */
                if (fila.getCell(6) == null) {
                    apoyoAnyadir.setLimpiezaBase(0);
                } else {
                    apoyoAnyadir.setLimpiezaBase(fila.getCell(6).getNumericCellValue());
                }

                /**
                 * PODA CALLE
                 */
                if (fila.getCell(7) == null) {
                    apoyoAnyadir.setPodaCalle(0);
                } else {
                    apoyoAnyadir.setPodaCalle(fila.getCell(7).getNumericCellValue());
                }

                /**
                 * FIJO SALIDA
                 */
                if (fila.getCell(8) == null) {
                    apoyoAnyadir.setFijoSalida(0);
                } else {
                    apoyoAnyadir.setFijoSalida(fila.getCell(8).getNumericCellValue());
                }

                /**
                 * DÍA
                 */
                apoyoAnyadir.setDia(fila.getCell(9).getNumericCellValue());

                /**
                 * CAPATAZ
                 */
                apoyoAnyadir.setCapataz(fila.getCell(10).getStringCellValue());

                /**
                 * OBSERVACIONES
                 */
                if (fila.getCell(11) == null) {
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
    public void creacionExcelApoyosRealizados(String nombreHoja, String codigoHoja, String nombreExcel) {
        String nombreArchivoSalida = "src/main/2023/"+provincias+"/MEDICIONES PARTES/"+zona+"/"+nombreExcel+".xlsx";
        File archivoSalida = new File(nombreArchivoSalida);

        FileOutputStream fileMod = null;

        if (archivoSalida.exists()) {
            // El archivo de salida ya existe, abre el libro existente
            try {
                FileInputStream file = new FileInputStream(nombreArchivoSalida);
                wb = new XSSFWorkbook(file);
                file.close();
            } catch (IOException e) {
                System.out.println("Error al abrir el archivo existente: " + e.getMessage());
                System.exit(-1);
            }
        } else {
            // El archivo de salida no existe, crea uno nuevo con el nombre proporcionado
            wb = new XSSFWorkbook();
        }

        /**
         * Comprobación de si la hoja ya existe en el excel
         */
        Sheet hoja = wb.getSheet(nombreHoja);
        if (hoja == null) {
            hoja = wb.createSheet(nombreHoja);
        }

        //Método que va a crear y rellenar mi excel de apoyos
        introducirValoresApoyos(hoja, codigoHoja, nombreHoja);

        try {
            fileMod = new FileOutputStream(nombreArchivoSalida);
            wb.write(fileMod);
            fileMod.close();
        } catch (IOException e) {
            System.out.println("Error al escribir EXCEL APOYOS, probablemente lo tengas abierto.\n");
            System.exit(-1);
        } finally {
            try {
                if (file != null) {
                    file.close();
                }
            } catch (IOException e) {
                System.out.println("Error al cerrar fichero");
                System.exit(-1);
            }

            try {
                if (fileMod != null) {
                    fileMod.close();
                }
            } catch (IOException e) {
                System.out.println("Error al cerrar EXCEL");
                System.exit(-1);
            }
        }
    }

    private void introducirValoresApoyos(Sheet hoja, String codigoHoja, String nombreHoja) {
        String numApoyo = "";
        double longitudMantenimineto = 0;
        double longitudLimpieza = 0;
        double longitudApertura = 0;
        double anomaliaVegetacion = 0;
        double limpiezaBase = 0;
        double podaCalle = 0;
        double fijoSalida = 0;
        double dia = 0;
        LocalDate diaLocalDate = null;
        Date fechaDate = null;
        String capataz = "";
        int numDiasTrabajados = 0;
        String pendienteTractor = "";
        String trabajoRematado = "";
        String observaciones = "";
        double contadorLongMant = 0;
        double contadorLongLimp = 0;
        double contadorLongApertura = 0;
        double contadorAnomalia = 0;
        double contadorLongitudLimpiezaBase = 0;
        double contadorPodaCalle = 0;
        double contadorFijoSalida = 0;
        double contadorNumeroDiasTrabajados = 0;

        int filaNueva;
        int filaAntiguaSumas = 0;
        if (hoja.getLastRowNum() > 1){
            filaNueva = hoja.getLastRowNum() - 2;
            filaAntiguaSumas = filaNueva;
        } else {
            filaNueva = 2;
        }

        //Fila de la que tomamos los valores sumatorios anteriores totales
        Row filaSumasAntigua = hoja.getRow(filaAntiguaSumas + 1);
        Row filaSeparacionDias = hoja.getRow(filaAntiguaSumas + 2);

        Double[] rowData = new Double[8];
        String[] rowDataTotalDiasTrabajados = new String[1];
        Double[] rowDataCuentas = new Double[8];
        String[] rowDataSeparacionDias = new String[1];
        rowDataSeparacionDias[0] = "";

        if (filaAntiguaSumas > 0) {
            for (int i = 0; i < 11; i++) {
                if (i < 8) {
                    Cell cell = filaSumasAntigua.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData[i] = cell.getNumericCellValue();

                    Cell cellCuentas = filaSeparacionDias.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowDataCuentas[i] = cellCuentas.getNumericCellValue();
                } else if (i == 10){
                    Cell cellDias = filaSumasAntigua.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowDataTotalDiasTrabajados[0] = cellDias.getStringCellValue();

                    Cell cellDiasSeparacion = filaSeparacionDias.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowDataSeparacionDias[0] = cellDiasSeparacion.getStringCellValue();
                }
            }
        }

        /**
         * Dar estilo de color y alineado para el título
         */
        CellStyle estiloCeldaTitulo = wb.createCellStyle();
        //COLOR
        estiloCeldaTitulo.setFillForegroundColor(IndexedColors.LIME.getIndex());
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
        //PERMITE QUE EL TEXTO SE ENVUELVA
        estiloCeldaTitulo.setWrapText(true);

        /**
         * Dar estilo de negrita y alineado para celdas con información
         */
        CellStyle estiloCeldaInfo = wb.createCellStyle();
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
        CellStyle estiloFecha = wb.createCellStyle();
        DataFormat formatoFecha = wb.createDataFormat();
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

        /**
         * INFO CODLINEA
         */

        Row filaCodLinea = hoja.createRow(0);
        Cell celdaTitulo = filaCodLinea.createCell(0);
        celdaTitulo.setCellValue(codigoHoja + " " + nombreHoja);

        /**
         * TITULOS
         */

        Row filaTitulos = hoja.createRow(1);
        Cell celdaColumnaApoyo = filaTitulos.createCell(0);
        celdaColumnaApoyo.setCellValue("APOYO");
        celdaColumnaApoyo.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaLongitudMantenimineto = filaTitulos.createCell(1);
        celdaColumnaLongitudMantenimineto.setCellValue("LONG\nMANT");
        celdaColumnaLongitudMantenimineto.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaLongitudLimpieza = filaTitulos.createCell(2);
        celdaColumnaLongitudLimpieza.setCellValue("LONG\nLIMPIEZA");
        celdaColumnaLongitudLimpieza.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaLongitudApertura = filaTitulos.createCell(3);
        celdaColumnaLongitudApertura.setCellValue("LONG\nAPERTURA");
        celdaColumnaLongitudApertura.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaAnomalia = filaTitulos.createCell(4);
        celdaColumnaAnomalia.setCellValue("ANOMALIA");
        celdaColumnaAnomalia.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaLongitudLimpiezaBase = filaTitulos.createCell(5);
        celdaColumnaLongitudLimpiezaBase.setCellValue("LIMPIEZA\nBASE");
        celdaColumnaLongitudLimpiezaBase.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaPodaCalle = filaTitulos.createCell(6);
        celdaColumnaPodaCalle.setCellValue("PODA\nCALLE");
        celdaColumnaPodaCalle.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaFijoSalida = filaTitulos.createCell(7);
        celdaColumnaFijoSalida.setCellValue("FIJO\nSALIDA");
        celdaColumnaFijoSalida.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaDia = filaTitulos.createCell(8);
        celdaColumnaDia.setCellValue("FECHA");
        celdaColumnaDia.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaCapataz = filaTitulos.createCell(9);
        celdaColumnaCapataz.setCellValue("CAPATAZ");
        celdaColumnaCapataz.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaCapatazNumDiasTrabajados = filaTitulos.createCell(10);
        celdaColumnaCapatazNumDiasTrabajados.setCellValue("Nº DIAS\nTRABAJADOS");
        celdaColumnaCapatazNumDiasTrabajados.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTractor = filaTitulos.createCell(11);
        celdaColumnaTractor.setCellValue("PENDIENTE\nTRACTOR");
        celdaColumnaTractor.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaTrabajoRematado = filaTitulos.createCell(12);
        celdaColumnaTrabajoRematado.setCellValue("TRABAJO\nREMATADO");
        celdaColumnaTrabajoRematado.setCellStyle(estiloCeldaTitulo);

        Cell celdaColumnaObservaciones = filaTitulos.createCell(13);
        celdaColumnaObservaciones.setCellValue("OBSERVACIONES");
        celdaColumnaObservaciones.setCellStyle(estiloCeldaTitulo);

        int filasTotales = listaApoyos.size() + 2;
        if (filaAntiguaSumas > 0){
            filasTotales = filasTotales - 2 + filaAntiguaSumas;
        }

        for (int i = 2; i < listaApoyos.size() + 2; i++) {

            /**
             * INFO APOYOS
             */
            Row fila = hoja.createRow(filaNueva);
            numApoyo = listaApoyos.get(i - 2).getNumApoyo();
            Cell celdaNumApoyo = fila.createCell(0);
            celdaNumApoyo.setCellValue(numApoyo);
            celdaNumApoyo.setCellStyle(estiloCeldaInfo);

            longitudMantenimineto = listaApoyos.get(i - 2).getLongitudMantenimineto();
            contadorLongMant += longitudMantenimineto;
            Cell celdaLongitudMantenimiento = fila.createCell(1);
            celdaLongitudMantenimiento.setCellValue(longitudMantenimineto);
            celdaLongitudMantenimiento.setCellStyle(estiloCeldaInfo);

            longitudLimpieza = listaApoyos.get(i - 2).getLongitudLimpieza();
            contadorLongLimp += longitudLimpieza;
            Cell celdaLongitudLimpeza = fila.createCell(2);
            celdaLongitudLimpeza.setCellValue(longitudLimpieza);
            celdaLongitudLimpeza.setCellStyle(estiloCeldaInfo);

            longitudApertura = listaApoyos.get(i - 2).getLongitudApertura();
            contadorLongApertura += longitudApertura;
            Cell celdaLongitudApertura = fila.createCell(3);
            celdaLongitudApertura.setCellValue(longitudApertura);
            celdaLongitudApertura.setCellStyle(estiloCeldaInfo);

            anomaliaVegetacion = listaApoyos.get(i - 2).getNumAnomalia();
            contadorAnomalia += anomaliaVegetacion;
            Cell celdaAnomaliaVegetacion = fila.createCell(4);
            celdaAnomaliaVegetacion.setCellValue(anomaliaVegetacion);
            celdaAnomaliaVegetacion.setCellStyle(estiloCeldaInfo);

            limpiezaBase = listaApoyos.get(i - 2).getLimpiezaBase();
            contadorLongitudLimpiezaBase += limpiezaBase;
            Cell celdaLimpiezaBase = fila.createCell(5);
            celdaLimpiezaBase.setCellValue(limpiezaBase);
            celdaLimpiezaBase.setCellStyle(estiloCeldaInfo);

            podaCalle = listaApoyos.get(i - 2).getPodaCalle();
            contadorPodaCalle += podaCalle;
            Cell celdaPodaCalle = fila.createCell(6);
            celdaPodaCalle.setCellValue(podaCalle);
            celdaPodaCalle.setCellStyle(estiloCeldaInfo);

            fijoSalida = listaApoyos.get(i - 2).getFijoSalida();
            contadorFijoSalida += fijoSalida;
            Cell celdaFijoSalida = fila.createCell(7);
            celdaFijoSalida.setCellValue(fijoSalida);
            celdaFijoSalida.setCellStyle(estiloCeldaInfo);

            dia = listaApoyos.get(i - 2).getDia();
            diaLocalDate = LocalDate.of(1899, 12, 30).plusDays((long) dia);
            fechaDate = Date.from(diaLocalDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            Cell celdaDia = fila.createCell(8);
            celdaDia.setCellValue(fechaDate);
            if (dia == 0){
                celdaDia.setCellValue(01/01/0001);
            }
            celdaDia.setCellStyle(estiloFecha);

            capataz = listaApoyos.get(i - 2).getCapataz();
            Cell celdaCapataz = fila.createCell(9);
            celdaCapataz.setCellValue(capataz);
            celdaCapataz.setCellStyle(estiloCeldaInfo);

            Cell celdaNumDiasTrabajados = fila.createCell(10);
            celdaNumDiasTrabajados.setCellValue("");
            celdaNumDiasTrabajados.setCellStyle(estiloCeldaInfo);

            Cell celdaPendienteTractor = fila.createCell(11);
            celdaPendienteTractor.setCellValue(pendienteTractor);
            celdaPendienteTractor.setCellStyle(estiloCeldaInfo);

            Cell celdaTrabajoRematado = fila.createCell(12);
            celdaTrabajoRematado.setCellValue(trabajoRematado);
            celdaTrabajoRematado.setCellStyle(estiloCeldaInfo);

            observaciones = listaApoyos.get(i - 2).getObservaciones();
            Cell celdaObservaciones = fila.createCell(13);
            celdaObservaciones.setCellValue(observaciones);
            celdaObservaciones.setCellStyle(estiloCeldaInfo);
            filaNueva++;
        }

        hoja.autoSizeColumn(0);
        hoja.autoSizeColumn(1);
        hoja.autoSizeColumn(2);
        hoja.autoSizeColumn(3);
        hoja.autoSizeColumn(4);
        hoja.autoSizeColumn(5);
        hoja.autoSizeColumn(6);
        hoja.autoSizeColumn(7);
        hoja.autoSizeColumn(8);
        hoja.autoSizeColumn(9);
        hoja.autoSizeColumn(10);
        hoja.autoSizeColumn(11);
        hoja.autoSizeColumn(12);
        hoja.autoSizeColumn(13);


        /**
         * CELDAS DE OPERACIONES FINALES
         */
        Row filaSumas = hoja.createRow(filaNueva + 1);

        double totalApoyos = 0;

        if (filaAntiguaSumas != 0) {

            Cell celdaColumnaSumaTotalApoyos = filaSumas.createCell(0);
            totalApoyos = listaApoyos.size() + rowData[0];
            celdaColumnaSumaTotalApoyos.setCellValue(totalApoyos);
            celdaColumnaSumaTotalApoyos.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudMantenimiento = filaSumas.createCell(1);
            celdaColumnaSumaTotalLongitudMantenimiento.setCellValue(contadorLongMant + rowData[1]);
            celdaColumnaSumaTotalLongitudMantenimiento.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudLimpieza = filaSumas.createCell(2);
            celdaColumnaSumaTotalLongitudLimpieza.setCellValue(contadorLongLimp + rowData[2]);
            celdaColumnaSumaTotalLongitudLimpieza.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudApertura = filaSumas.createCell(3);
            celdaColumnaSumaTotalLongitudApertura.setCellValue(contadorLongApertura + rowData[3]);
            celdaColumnaSumaTotalLongitudApertura.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalAnomalia = filaSumas.createCell(4);
            celdaColumnaSumaTotalAnomalia.setCellValue(contadorAnomalia + rowData[4]);
            celdaColumnaSumaTotalAnomalia.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLimpiezaBase = filaSumas.createCell(5);
            celdaColumnaSumaTotalLimpiezaBase.setCellValue(contadorLongitudLimpiezaBase + rowData[5]);
            celdaColumnaSumaTotalLimpiezaBase.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalPodaCalle = filaSumas.createCell(6);
            celdaColumnaSumaTotalPodaCalle.setCellValue(contadorPodaCalle + rowData[6]);
            celdaColumnaSumaTotalPodaCalle.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalFijoSalida = filaSumas.createCell(7);
            celdaColumnaSumaTotalFijoSalida.setCellValue(contadorFijoSalida + rowData[7]);
            celdaColumnaSumaTotalFijoSalida.setCellStyle(estiloCeldaTitulo);

        } else {

            Cell celdaColumnaSumaTotalApoyos = filaSumas.createCell(0);
            totalApoyos = listaApoyos.size();
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

            Cell celdaColumnaSumaTotalLimpiezaBase = filaSumas.createCell(5);
            celdaColumnaSumaTotalLimpiezaBase.setCellValue(contadorLongitudLimpiezaBase);
            celdaColumnaSumaTotalLimpiezaBase.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalPodaCalle = filaSumas.createCell(6);
            celdaColumnaSumaTotalPodaCalle.setCellValue(contadorPodaCalle);
            celdaColumnaSumaTotalPodaCalle.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalFijoSalida = filaSumas.createCell(7);
            celdaColumnaSumaTotalFijoSalida.setCellValue(contadorFijoSalida);
            celdaColumnaSumaTotalFijoSalida.setCellStyle(estiloCeldaTitulo);

        }

        HashMap<String, Integer> diasTrabajadosPorCapataz = obtenerDiasTrabajadosPorCapataz(mapaApoyos);

        if (filaAntiguaSumas != 0) {

            String totalTrabajadorAntiguoString = rowDataTotalDiasTrabajados[0];
            int startIndex = totalTrabajadorAntiguoString.lastIndexOf(" ") + 1; // Encuentra el último espacio y obtiene el índice siguiente
            String numeroString = totalTrabajadorAntiguoString.substring(startIndex); // Extrae la subcadena que contiene el número como string
            int totalDiasTrabajadosAntiguo = Integer.parseInt(numeroString);

            // Ahora imprimimos la información utilizando un bucle for tradicional
            for (Map.Entry<String, Integer> entry : diasTrabajadosPorCapataz.entrySet()) {
                int totalDiasTrabajados = entry.getValue();
                contadorNumeroDiasTrabajados += totalDiasTrabajados;
            }
            int intContadorNumeroDiasTrabajados = (int) contadorNumeroDiasTrabajados;
            Cell celdaColumnaSumaTotalNumDiasTrabaj = filaSumas.createCell(10);
            celdaColumnaSumaTotalNumDiasTrabaj.setCellValue("Total días trabajados: " + (intContadorNumeroDiasTrabajados + totalDiasTrabajadosAntiguo));
            celdaColumnaSumaTotalNumDiasTrabaj.setCellStyle(estiloCeldaTitulo);

        } else {

            for (Map.Entry<String, Integer> entry : diasTrabajadosPorCapataz.entrySet()) {
                int totalDiasTrabajados = entry.getValue();
                contadorNumeroDiasTrabajados += totalDiasTrabajados;
            }
            int intContadorNumeroDiasTrabajados = (int) contadorNumeroDiasTrabajados;
            Cell celdaColumnaSumaTotalNumDiasTrabaj = filaSumas.createCell(10);
            celdaColumnaSumaTotalNumDiasTrabaj.setCellValue("Total días trabajados: " + intContadorNumeroDiasTrabajados);
            celdaColumnaSumaTotalNumDiasTrabaj.setCellStyle(estiloCeldaTitulo);

        }

        /**
         * CELDAS OPERACIONES FINALES CON RESPECTIVAS DIVISIONES
         */

        Row filaSumasDivisiones = hoja.createRow(filaNueva + 2);

        if (filaAntiguaSumas != 0) {

            Cell celdaColumnaSumaTotalApoyosDivision = filaSumasDivisiones.createCell(0);
            celdaColumnaSumaTotalApoyosDivision.setCellValue(totalApoyos);
            celdaColumnaSumaTotalApoyosDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudMantenimientoDivision = filaSumasDivisiones.createCell(1);
            celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellValue(contadorLongMant / 1000 + rowDataCuentas[1]);
            celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudLimpiezaDivision = filaSumasDivisiones.createCell(2);
            celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellValue(contadorLongLimp / 1000 + rowDataCuentas[2]);
            celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudAperturaDivision = filaSumasDivisiones.createCell(3);
            celdaColumnaSumaTotalLongitudAperturaDivision.setCellValue(contadorLongApertura / 1000 + rowDataCuentas[3]);
            celdaColumnaSumaTotalLongitudAperturaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalAnomaliaDivision = filaSumasDivisiones.createCell(4);
            celdaColumnaSumaTotalAnomaliaDivision.setCellValue(contadorAnomalia + rowDataCuentas[4]);
            celdaColumnaSumaTotalAnomaliaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLimpiezaBaseDivision = filaSumasDivisiones.createCell(5);
            celdaColumnaSumaTotalLimpiezaBaseDivision.setCellValue(contadorLongitudLimpiezaBase + rowDataCuentas[5]);
            celdaColumnaSumaTotalLimpiezaBaseDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalPodaCalleDivision = filaSumasDivisiones.createCell(6);
            celdaColumnaSumaTotalPodaCalleDivision.setCellValue(contadorPodaCalle / 1000 + rowDataCuentas[6]);
            celdaColumnaSumaTotalPodaCalleDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalFijoSalidaDivision = filaSumasDivisiones.createCell(7);
            celdaColumnaSumaTotalFijoSalidaDivision.setCellValue(contadorFijoSalida + rowDataCuentas[7]);
            celdaColumnaSumaTotalFijoSalidaDivision.setCellStyle(estiloCeldaTitulo);

        } else {

            Cell celdaColumnaSumaTotalApoyosDivision = filaSumasDivisiones.createCell(0);
            celdaColumnaSumaTotalApoyosDivision.setCellValue(totalApoyos);
            celdaColumnaSumaTotalApoyosDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudMantenimientoDivision = filaSumasDivisiones.createCell(1);
            celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellValue(contadorLongMant / 1000);
            celdaColumnaSumaTotalLongitudMantenimientoDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudLimpiezaDivision = filaSumasDivisiones.createCell(2);
            celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellValue(contadorLongLimp / 1000);
            celdaColumnaSumaTotalLongitudLimpiezaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLongitudAperturaDivision = filaSumasDivisiones.createCell(3);
            celdaColumnaSumaTotalLongitudAperturaDivision.setCellValue(contadorLongApertura / 1000);
            celdaColumnaSumaTotalLongitudAperturaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalAnomaliaDivision = filaSumasDivisiones.createCell(4);
            celdaColumnaSumaTotalAnomaliaDivision.setCellValue(contadorAnomalia);
            celdaColumnaSumaTotalAnomaliaDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalLimpiezaBaseDivision = filaSumasDivisiones.createCell(5);
            celdaColumnaSumaTotalLimpiezaBaseDivision.setCellValue(contadorLongitudLimpiezaBase);
            celdaColumnaSumaTotalLimpiezaBaseDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalPodaCalleDivision = filaSumasDivisiones.createCell(6);
            celdaColumnaSumaTotalPodaCalleDivision.setCellValue(contadorPodaCalle / 1000);
            celdaColumnaSumaTotalPodaCalleDivision.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaSumaTotalFijoSalidaDivision = filaSumasDivisiones.createCell(7);
            celdaColumnaSumaTotalFijoSalidaDivision.setCellValue(contadorFijoSalida);
            celdaColumnaSumaTotalFijoSalidaDivision.setCellStyle(estiloCeldaTitulo);
        }


        /**
         * SEPARACIÓN DÍAS TRABAJADOS
         */

        StringBuilder diasTrabajadosConcatenados = new StringBuilder();
        for (Map.Entry<String, Integer> entry : diasTrabajadosPorCapataz.entrySet()) {
            capataz = entry.getKey();
            int totalDiasTrabajados = entry.getValue();
            diasTrabajadosConcatenados.append("Capataz: ").append(capataz).append(", Días trabajados: ").append(totalDiasTrabajados).append("\n");
        }

        Cell celdaColumnaSumaTotalNumDiasTrabajDivision = filaSumasDivisiones.createCell(10);
        celdaColumnaSumaTotalNumDiasTrabajDivision.setCellValue(rowDataSeparacionDias[0] + diasTrabajadosConcatenados.toString());
        celdaColumnaSumaTotalNumDiasTrabajDivision.setCellStyle(estiloCeldaTitulo);
    }

    public static HashMap<String, Integer> obtenerDiasTrabajadosPorCapataz(HashMap<Integer, Apoyo> mapaApoyos) {
        HashMap<String, HashSet<String>> diasTrabajadosPorPersona = new HashMap<>();

        for (Apoyo apoyo : mapaApoyos.values()) {
            String capataz = apoyo.getCapataz();
            String fecha = Double.toString(apoyo.getDia());

            if (diasTrabajadosPorPersona.containsKey(capataz)) {
                HashSet<String> fechasTrabajadas = diasTrabajadosPorPersona.get(capataz);
                if (!fechasTrabajadas.contains(fecha)) {
                    fechasTrabajadas.add(fecha);
                    diasTrabajadosPorPersona.put(capataz, fechasTrabajadas);
                }
            } else {
                HashSet<String> fechasTrabajadas = new HashSet<>();
                fechasTrabajadas.add(fecha);
                diasTrabajadosPorPersona.put(capataz, fechasTrabajadas);
            }
        }

        // Calcular el total de días trabajados por cada capataz
        HashMap<String, Integer> diasTrabajadosPorCapataz = new HashMap<>();
        for (String capataz : diasTrabajadosPorPersona.keySet()) {
            HashSet<String> fechasTrabajadas = diasTrabajadosPorPersona.get(capataz);
            int totalDiasTrabajados = fechasTrabajadas.size();
            diasTrabajadosPorCapataz.put(capataz, totalDiasTrabajados);
        }

        return diasTrabajadosPorCapataz;
    }
}