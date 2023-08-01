import java.io.IOException;

/**
 * @author Roberto Fernández Higuera
 */

public class Main {

    /**
     * Lanzador de la aplicación
     *
     * @param args
     */
    public static void main(String[] args) throws InterruptedException, IOException {

        String brightMagentaColor = "\u001B[95m";

        System.out.print("\n"+brightMagentaColor);

        System.out.println("*********************************************");
        System.out.println("** PROGRAMA MEDICIONES PARTES TALARTOS S.L **");
        System.out.println("*********************************************");

        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        System.out.println("\n");

        System.out.println("CONSEJO: TEN CUIDADO A LA HORA DE ESCRIBIR, AL MÍNIMO FALLO GRAMÁTICO LA APLICACIÓN NO FUNCIONARÁ CORRÉCTAMENTE :))\n");

        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }


        Consola consola = new Consola();
        consola.ejecucionPrograma();
        // Fin ejecución
        System.out.println("FIN\n");
        System.exit(0);
    }
}