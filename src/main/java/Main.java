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

        System.out.print("*********************************************\n");
        System.out.print("** PROGRAMA MEDICIONES PARTES TALARTOS S.L **\n");
        System.out.print("*********************************************\n");

        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        System.out.println(" \n");

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