package utils;


import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;

public class SystemOuputHelper {

    public static ByteArrayOutputStream baos;
    public static PrintStream old;

    public void switchSystemOutputForPowershell() throws IOException {

        // *** Create a stream to hold the output
        baos = new ByteArrayOutputStream();
        PrintStream ps = new PrintStream(baos);

        // *** IMPORTANT: Save the old System.out!
        old = System.out;

        // *** Tell Java to use special stream
        System.setOut(ps);
    }
}
