/**
 * @author Roberto Fernández Higuera
 */

public class ExcelManager(){

    public ExcelManager(){
        try {
            this.file = new FileInputStream("src/main/resources/RIVALAMORA-ESTADIO 5 MAYO.xlsx");
            this.wb = new XSSFWorkbook(file);
        } catch (IOException e) {
            System.out.println("Error al encontrar el fichero excel 1");
            System.exit(-1);
        }
    }

    public void creacionExcelApoyosRealizados(){

    }

    public void creacionExcelControlCapataces(){

    }

}