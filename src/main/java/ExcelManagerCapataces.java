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
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;


public class ExcelManagerCapataces {

    private static FileInputStream fileCapataces;
    private static XSSFWorkbook wbCapataces;
    private XSSFSheet hojaApoyos;

    //Arraylist con los valores de todos los campos del Excel capataces
    private ArrayList<Capataz> listaCapataces = new ArrayList<>();

    //MAPAS A UTILIZAR
    private HashMap<String, ArrayList<Capataz>> mapaCapataces = new HashMap<>();

    private String rutaExcel;
    private String nombreExcel;

    /**
     * CONSTRUCTOR DE LA CLASE ENCARGADO DE LEER LAS PARTES DEL EXCEL
     */

    public ExcelManagerCapataces(String nombreExcel) {
        this.nombreExcel = nombreExcel;

        rutaExcel = "EXCELS_FINALES/EXCELS_APOYO/"+nombreExcel+".xlsx";

        try {
            this.fileCapataces = new FileInputStream(rutaExcel);
            this.wbCapataces = new XSSFWorkbook(fileCapataces);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel: "+nombreExcel);
            System.exit(-1);
        }

        this.mapaCapataces = leerTodosLosNombresCapataces();
    }

    /**
     * Se recorren todas las hojas del archivo Excel de apoyo y se lee el nombre del capataz de cada hoja.
     * Luego, se crea un HashMap llamado nombresCapatacesMap donde la clave es el nombre del capataz y el
     * valor es una lista de objetos Capataz que contienen los datos correspondientes a ese capataz.
     * @return
     */
    private HashMap<String, ArrayList<Capataz>> leerTodosLosNombresCapataces() {
        HashMap<String, ArrayList<Capataz>> nombresCapatacesMap = new HashMap<>();

        for (int i = 0; i < wbCapataces.getNumberOfSheets(); i++) {
            hojaApoyos = wbCapataces.getSheetAt(i);

            ArrayList<Capataz> capatacesEnHoja = leerDatosCapataces();
            if (!capatacesEnHoja.isEmpty()) {
                for (Capataz capataz : capatacesEnHoja) {
                    String nombreCapataz = capataz.getNombreApoyo().toUpperCase();
                    String claveUnica = nombreCapataz;
                    if (!nombresCapatacesMap.containsKey(claveUnica)) {
                        nombresCapatacesMap.put(claveUnica, new ArrayList<>());
                    }
                    nombresCapatacesMap.get(claveUnica).add(capataz);
                }
                listaCapataces.addAll(capatacesEnHoja);
            }
        }

        return nombresCapatacesMap;
    }

    /**
     * TODO NECESITO SABER DE DONDE SALE CADA VALOR Y SABER SOBRE QUÉ HOJA EXCEL LOS TOMO
     *
     * @return MAPA CAPATACES
     */
    private  ArrayList<Capataz> leerDatosCapataces() {
        int numFilas;
        ArrayList<Capataz> todosCapataces = new ArrayList<>();
        HashMap<String, Capataz> datosPorFechaCapataz = new HashMap<>();

        for (int i = 0; i < wbCapataces.getNumberOfSheets(); i++) {
            hojaApoyos = wbCapataces.getSheetAt(i);
            numFilas = hojaApoyos.getLastRowNum() - 1;

            for (int j = 2; j < numFilas; j++) {
                Row fila = hojaApoyos.getRow(j);
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
                    if (fila != null && fila.getCell(8) != null) {
                        Cell fechaCell = fila.getCell(8);
                        if (fechaCell.getCellType() == CellType.NUMERIC) {
                            double fechaExcel = fechaCell.getNumericCellValue();
                            capatazAnyadir.setDia(fechaExcel);
                        } else {
                            throw new RuntimeException("La celda no es de tipo numérico.");
                        }
                    } else {

                    }

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
                    if (fila.getCell(13).getStringCellValue().equals("")) {
                        capatazAnyadir.setObservaciones("");
                    } else {
                        capatazAnyadir.setObservaciones(fila.getCell(13).getStringCellValue() + " ");
                    }

                    /**
                     * COD LÍNEA Y NOMBRE LÍNEA
                     */
                    capatazAnyadir.setCodLinea(hojaApoyos.getRow(0).getCell(0).getStringCellValue() + " ");

                    String nombreCapataz = capatazAnyadir.getNombreApoyo();
                    double fecha = capatazAnyadir.getDia();
                    String claveFechaCapataz = fecha + "-" + nombreCapataz;

                    if (datosPorFechaCapataz.containsKey(claveFechaCapataz)) {
                        // Si ya existe una entrada para esta fecha y capataz, actualiza los valores
                        Capataz capatazTemporal = datosPorFechaCapataz.get(claveFechaCapataz);

                        capatazTemporal.setNumApoyos(capatazTemporal.getNumApoyos() + 1);
                        capatazTemporal.setFijoSalida(capatazTemporal.getFijoSalida() + capatazAnyadir.getFijoSalida());
                        capatazTemporal.setLongMantenimiento(capatazTemporal.getLongMantenimiento() + capatazAnyadir.getLongMantenimiento());
                        capatazTemporal.setAnomalia(capatazTemporal.getAnomalia() + capatazAnyadir.getAnomalia());
                        capatazTemporal.setLongApertura(capatazTemporal.getLongApertura() + capatazAnyadir.getLongApertura());
                        capatazTemporal.setTalasFuera(capatazTemporal.getTalasFuera() + capatazAnyadir.getTalasFuera());
                        capatazTemporal.setLongitudLimpieza(capatazTemporal.getLongitudLimpieza() + capatazAnyadir.getLimpiezaBase());
                        capatazTemporal.setKm(capatazTemporal.getKm() + capatazAnyadir.getKm());
                        capatazTemporal.setImporteMedios(capatazTemporal.getImporteMedios() + capatazAnyadir.getImporteMedios());
                        capatazTemporal.setImporteCoeficiente(capatazTemporal.getImporteCoeficiente() + capatazAnyadir.getImporteCoeficiente());
                        capatazTemporal.setZona(capatazTemporal.getZona() + capatazAnyadir.getZona());
                        capatazTemporal.setObservaciones(capatazTemporal.getObservaciones() + capatazAnyadir.getObservaciones());
                        capatazTemporal.setCodLinea(capatazTemporal.getCodLinea() + capatazAnyadir.getCodLinea());

                        datosPorFechaCapataz.put(claveFechaCapataz, capatazTemporal);
                    } else {
                        // Si no existe una entrada, crea una nueva entrada en el mapa
                        capatazAnyadir.setNumApoyos(1);
                        datosPorFechaCapataz.put(claveFechaCapataz, capatazAnyadir);
                    }

                }
            }
        }
        todosCapataces.addAll(datosPorFechaCapataz.values());

        return todosCapataces;
    }

    /**
     * Se recorre la lista de capataces y se crea una hoja en el archivo Excel de capataces para cada capataz encontrado.
     * @param nombreExcel
     * @param zona
     * @param codLinea
     */
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

        // Utilizar un HashSet para almacenar los nombres únicos
        HashSet<String> nombresUnicos = new HashSet<>();
        nombresUnicos.addAll(mapaCapataces.keySet());

        for (String nombreCapataz : nombresUnicos) {
            ArrayList<Capataz> capatacesEnHoja = mapaCapataces.get(nombreCapataz);
            for (int i = 0; i < capatacesEnHoja.size(); i++) {
                System.out.println(capatacesEnHoja.get(i).getNombreApoyo() + " - " + capatacesEnHoja.size());
            }
            if (capatacesEnHoja != null) {
                String nombreCapatazMayus = nombreCapataz.toUpperCase();
                XSSFSheet hoja = wbCapataces.getSheet(nombreCapatazMayus);
                if (hoja == null) {
                    hoja = wbCapataces.createSheet(nombreCapatazMayus);
                }
                introducirValoresCapataz(hoja, zona, codLinea, nombreCapatazMayus, capatacesEnHoja);
            }
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

    /**
     * Inserta los datos de cada capataz en su hoja correspondiente. Este método recibe la hoja, el nombre de la zona,
     * el código de la línea, el nombre del capataz y la lista de capataces correspondiente a esa hoja. Luego, utiliza
     * los datos de cada capataz para escribir en las celdas de la hoja.
     * @param hoja
     * @param zona
     * @param codLinea
     * @param nombreCapataz
     * @param capatacesEnHoja
     */
    private void introducirValoresCapataz(Sheet hoja, String zona, String codLinea, String nombreCapataz, ArrayList<Capataz> capatacesEnHoja) {
        double fecha = 0;
        LocalDate diaLocalDate = null;
        Date fechaDate = null;
        double numApoyos = 0;
        double fijoSalida = 0;
        double longMantenimiento = 0;
        double anomalia = 0;
        double longApertura = 0;
        double talasFuera = 0;
        double limpiezaBase = 0;
        double km = 0;
        double importeMedios = 0;
        double importeCoeficiente = 0;
        String observaciones = "";
        double contadorFijoSalida = 0;
        double contadorLongMantenimiento = 0;
        double contadorAnomalia = 0;
        double contadorLongApertura = 0;
        double contadorTalasFuera = 0;
        double contadorLimpiezaBase = 0;
        double contadorKm = 0;
        double contadorImporteMedio = 0;
        double contadorImporteCoeficiente = 0;
        double contadorNumApoyos = 0;
        double importeCoeficienteSemanal = 0;

        int filaNueva;
        int filaAntiguaSumas = 0;
        if (hoja.getLastRowNum() > 1){
            filaNueva = hoja.getLastRowNum() - 1;
            filaAntiguaSumas = filaNueva;
        } else {
            filaNueva = 1;
        }

        //Fila de la que tomamos los valores anteriores totales
        Row filaSumasAntigua = hoja.getRow(filaAntiguaSumas);

        Double[] rowData = new Double[14];

        if (filaAntiguaSumas > 0) {
            for (int i = 0; i < 14; i++) {
                Cell cell = filaSumasAntigua.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                rowData[i] = cell.getNumericCellValue();
            }
        }

        Row filaSumasImportesAntigua = hoja.getRow(filaAntiguaSumas+1);

        Double[] rowDataImporte = new Double[14];

        if (filaAntiguaSumas > 0) {
            for (int i = 0; i < 14; i++) {
                Cell cell = filaSumasImportesAntigua.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                rowDataImporte[i] = cell.getNumericCellValue();
            }
        }

        //TÍTULOS

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

        int filasTotales = (capatacesEnHoja.size() + 1)/2;
        if (filaAntiguaSumas > 0){
            filasTotales = capatacesEnHoja.size();
        }

        for (int k = 0; k < filasTotales; k++) {
            Capataz capataz = capatacesEnHoja.get(k);

            Row filaTitulos = hoja.createRow(0);

            Cell celdaColumnaDia = filaTitulos.createCell(0);
            celdaColumnaDia.setCellValue("DÍA");
            celdaColumnaDia.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaApoyos = filaTitulos.createCell(1);
            celdaColumnaApoyos.setCellValue("APOYOS");
            celdaColumnaApoyos.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaFijoSalida = filaTitulos.createCell(2);
            celdaColumnaFijoSalida.setCellValue("FIJO\nSALIDA");
            celdaColumnaFijoSalida.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaLongitudMantenimineto = filaTitulos.createCell(3);
            celdaColumnaLongitudMantenimineto.setCellValue("LONG\nMANT");
            celdaColumnaLongitudMantenimineto.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaAnomalia = filaTitulos.createCell(4);
            celdaColumnaAnomalia.setCellValue("ANOMALIAS");
            celdaColumnaAnomalia.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaLongitudApertura = filaTitulos.createCell(5);
            celdaColumnaLongitudApertura.setCellValue("LONGITUD\nAPERTURA");
            celdaColumnaLongitudApertura.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTalasFueraCalle = filaTitulos.createCell(6);
            celdaColumnaTalasFueraCalle.setCellValue("TALAS FUERA\nDE LA CALLE");
            celdaColumnaTalasFueraCalle.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaLimpiezaBase = filaTitulos.createCell(7);
            celdaColumnaLimpiezaBase.setCellValue("LIMPIEZA\nBASE");
            celdaColumnaLimpiezaBase.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaIdentZonasNuecas = filaTitulos.createCell(8);
            celdaColumnaIdentZonasNuecas.setCellValue("KM");
            celdaColumnaIdentZonasNuecas.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaImporte = filaTitulos.createCell(9);
            celdaColumnaImporte.setCellValue("IMPORTE\nMEDIOS");
            celdaColumnaImporte.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaImporteCoeficiente = filaTitulos.createCell(10);
            celdaColumnaImporteCoeficiente.setCellValue("IMPORTE\nCOEFICIENTE");
            celdaColumnaImporteCoeficiente.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaZONA = filaTitulos.createCell(11);
            celdaColumnaZONA.setCellValue("ZONA");
            celdaColumnaZONA.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaOBSERVACIONES = filaTitulos.createCell(12);
            celdaColumnaOBSERVACIONES.setCellValue("OBSERVACIONES");
            celdaColumnaOBSERVACIONES.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaCODLINEA = filaTitulos.createCell(13);
            celdaColumnaCODLINEA.setCellValue("COD LINEA\nY\nNOMBRE LINEA");
            celdaColumnaCODLINEA.setCellStyle(estiloCeldaTitulo);

            /**
             * INFO CAPATACES
             */
            Row fila = hoja.createRow(filaNueva);

            // DÍA
            fecha = capataz.getDia();
            diaLocalDate = LocalDate.of(1899, 12, 30).plusDays((long) fecha);
            fechaDate = Date.from(diaLocalDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            Cell celdaFecha = fila.createCell(0);
            celdaFecha.setCellValue(fechaDate);
            celdaFecha.setCellStyle(estiloFecha);

            System.out.println("Estoy en la hoja de "+ capataz.getNombreApoyo()+ " en el día " + fecha);

            // Número de apoyos
            numApoyos = capataz.getNumApoyos();
            contadorNumApoyos += numApoyos;
            Cell celdaNumApoyos = fila.createCell(1);
            celdaNumApoyos.setCellValue(numApoyos);
            celdaNumApoyos.setCellStyle(estiloCeldaInfo);

            // Fijo salida
            fijoSalida = capataz.getFijoSalida();
            contadorFijoSalida +=fijoSalida;
            Cell celdaFijoSalida = fila.createCell(2);
            celdaFijoSalida.setCellValue(fijoSalida);
            celdaFijoSalida.setCellStyle(estiloCeldaInfo);

            // Longitud mantenimiento
            longMantenimiento = capataz.getLongMantenimiento();
            contadorLongMantenimiento += longMantenimiento;
            Cell celdaLongMantenimiento = fila.createCell(3);
            celdaLongMantenimiento.setCellValue(longMantenimiento);
            celdaLongMantenimiento.setCellStyle(estiloCeldaInfo);

            // Anomalía
            anomalia = capataz.getAnomalia();
            contadorAnomalia += anomalia;
            Cell celdaAnomalia = fila.createCell(4);
            celdaAnomalia.setCellValue(anomalia);
            celdaAnomalia.setCellStyle(estiloCeldaInfo);

            // Longitud apertura
            longApertura = capataz.getLongApertura();
            contadorLongApertura += longApertura;
            Cell celdaLongApertura = fila.createCell(5);
            celdaLongApertura.setCellValue(longApertura);
            celdaLongApertura.setCellStyle(estiloCeldaInfo);

            // Talas fuera
            talasFuera = capataz.getTalasFuera();
            contadorTalasFuera += talasFuera;
            Cell celdaTalasFuera = fila.createCell(6);
            celdaTalasFuera.setCellValue(talasFuera);
            celdaTalasFuera.setCellStyle(estiloCeldaInfo);

            // Limpieza base
            limpiezaBase = capataz.getLimpiezaBase();
            contadorLimpiezaBase += limpiezaBase;
            Cell celdaLimpiezaBase = fila.createCell(7);
            celdaLimpiezaBase.setCellValue(limpiezaBase);
            celdaLimpiezaBase.setCellStyle(estiloCeldaInfo);

            // KM
            km = capataz.getKm();
            contadorKm += km;
            Cell celdaKm = fila.createCell(8);
            celdaKm.setCellValue(km);
            celdaKm.setCellStyle(estiloCeldaInfo);

            // Importe medios
            importeMedios = capataz.getImporteMedios();
            contadorImporteMedio += importeMedios;
            Cell celdaImporteMedios = fila.createCell(9);
            celdaImporteMedios.setCellValue(importeMedios);
            celdaImporteMedios.setCellStyle(estiloCeldaInfo);

            // Importe coeficiente
            importeCoeficiente = capataz.getImporteCoeficiente();
            contadorImporteCoeficiente += importeCoeficiente;
            Cell celdaImporteCoeficiente = fila.createCell(10);
            celdaImporteCoeficiente.setCellValue(importeCoeficiente);
            celdaImporteCoeficiente.setCellStyle(estiloCeldaInfo);

            // Zona
            zona = obtenerNombreCarpetaArchivo(rutaExcel);
            Cell celdaZona = fila.createCell(11);
            celdaZona.setCellValue(zona);
            celdaZona.setCellStyle(estiloCeldaInfo);

            // Observaciones
            observaciones = capataz.getObservaciones();
            Cell celdaObservaciones = fila.createCell(12);
            celdaObservaciones.setCellValue(observaciones);
            celdaObservaciones.setCellStyle(estiloCeldaInfo);

            // Cod Línea
            codLinea = capataz.getCodLinea();
            Cell celdaCodLinea = fila.createCell(13);
            celdaCodLinea.setCellValue(codLinea);
            celdaCodLinea.setCellStyle(estiloCeldaInfo);

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
        Row filaSumas = hoja.createRow(filaNueva);

        if (filaAntiguaSumas != 0) {
            Cell celdaColumnaTotalApoyos = filaSumas.createCell(1);
            celdaColumnaTotalApoyos.setCellValue(contadorNumApoyos + rowData[1]);
            celdaColumnaTotalApoyos.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalFijoSalida = filaSumas.createCell(2);
            celdaColumnaTotalFijoSalida.setCellValue(contadorFijoSalida + rowData[2]);
            celdaColumnaTotalFijoSalida.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalLongMantenimiento = filaSumas.createCell(3);
            celdaColumnaTotalLongMantenimiento.setCellValue(contadorLongMantenimiento + rowData[3]);
            celdaColumnaTotalLongMantenimiento.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalAnomalia = filaSumas.createCell(4);
            celdaColumnaTotalAnomalia.setCellValue(contadorAnomalia + rowData[4]);
            celdaColumnaTotalAnomalia.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalLongApertura = filaSumas.createCell(5);
            celdaColumnaTotalLongApertura.setCellValue(contadorLongApertura + rowData[5]);
            celdaColumnaTotalLongApertura.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalTalasFuera = filaSumas.createCell(6);
            celdaColumnaTotalTalasFuera.setCellValue(contadorTalasFuera + rowData[6]);
            celdaColumnaTotalTalasFuera.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalLimpiezaBase = filaSumas.createCell(7);
            celdaColumnaTotalLimpiezaBase.setCellValue(contadorLimpiezaBase + rowData[7]);
            celdaColumnaTotalLimpiezaBase.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalKm = filaSumas.createCell(8);
            celdaColumnaTotalKm.setCellValue(contadorKm + rowData[8]);
            celdaColumnaTotalKm.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalImporteMedio = filaSumas.createCell(9);
            celdaColumnaTotalImporteMedio.setCellValue(contadorImporteMedio + rowData[9]);
            celdaColumnaTotalImporteMedio.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalImporteCoeficiente = filaSumas.createCell(10);
            celdaColumnaTotalImporteCoeficiente.setCellValue(contadorImporteCoeficiente + rowData[10]);
            celdaColumnaTotalImporteCoeficiente.setCellStyle(estiloCeldaTitulo);

        } else {

            Cell celdaColumnaTotalApoyos = filaSumas.createCell(1);
            celdaColumnaTotalApoyos.setCellValue(contadorNumApoyos);
            celdaColumnaTotalApoyos.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalFijoSalida = filaSumas.createCell(2);
            celdaColumnaTotalFijoSalida.setCellValue(contadorFijoSalida);
            celdaColumnaTotalFijoSalida.setCellStyle(estiloCeldaTitulo);

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

        }
        /**
         * PREGUNTAR A INÉS SOBRE ESTO
         */

        Row filaImporteCoeficienteSemanal = hoja.createRow(filaNueva + 1);

        if (filaAntiguaSumas != 0) {
            Cell celdaColumnaTextoParaCoeficienteSemanala = filaImporteCoeficienteSemanal.createCell(11);
            celdaColumnaTextoParaCoeficienteSemanala.setCellValue("IMPORTE\nSEMANAL:");
            celdaColumnaTextoParaCoeficienteSemanala.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalImporteCoeficienteSemanal = filaImporteCoeficienteSemanal.createCell(12);
            importeCoeficienteSemanal = contadorImporteCoeficiente / 7;
            celdaColumnaTotalImporteCoeficienteSemanal.setCellValue(importeCoeficienteSemanal + rowDataImporte[12]);
            celdaColumnaTotalImporteCoeficienteSemanal.setCellStyle(estiloCeldaTitulo);

        } else {

            Cell celdaColumnaTextoParaCoeficienteSemanala = filaImporteCoeficienteSemanal.createCell(11);
            celdaColumnaTextoParaCoeficienteSemanala.setCellValue("IMPORTE\nSEMANAL:");
            celdaColumnaTextoParaCoeficienteSemanala.setCellStyle(estiloCeldaTitulo);

            Cell celdaColumnaTotalImporteCoeficienteSemanal = filaImporteCoeficienteSemanal.createCell(12);
            importeCoeficienteSemanal = contadorImporteCoeficiente / 7;
            celdaColumnaTotalImporteCoeficienteSemanal.setCellValue(importeCoeficienteSemanal);
            celdaColumnaTotalImporteCoeficienteSemanal.setCellStyle(estiloCeldaTitulo);
        }

    }

    public static String obtenerNombreCarpetaArchivo(String rutaExcel) {
        Path path = Paths.get(rutaExcel);
        Path carpeta1 = path.getParent().getParent();
        Path carpeta2 = path.getParent();
        String nombreCarpeta = "";

        if (carpeta1 != null && carpeta2 != null) {
            nombreCarpeta = carpeta1.getFileName().toString() + " -> " +  carpeta2.getFileName().toString();
        } else {

        }
        return nombreCarpeta;
    }
}