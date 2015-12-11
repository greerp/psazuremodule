package utils;

import java.io.File;
import java.io.IOException;
import java.util.Properties;
import utils.PropertiesManager;

import static org.apache.commons.io.FileUtils.copyFile;
import static org.apache.commons.io.FileUtils.deleteQuietly;
import static org.apache.commons.io.FileUtils.forceDelete;

public class FileHelpers {

    public static File destination;
    PropertiesManager props = PropertiesManager.getInstance();

    public void Azure_RM_ScriptFile_Initializer() throws IOException {

        //Get the source template file
        File source = new File(props.getValue("sourcePowershellScriptFile"));

        //Declare intended new file name and location
        destination = new File(props.getValue("destinationPowershellScriptFile"));

        copyFile(source, destination);
    }

    public static void Azure_RM_ScriptFile_Delete() throws IOException {
        File deleteThis = new File("\"E:\\programs\\bamboo\\agent01\\home\\xml-data\\build-dir\\AZ-AZ-TR\\Test\\scripts\\Azure_Test_Script_Runner.ps1\"");
        deleteQuietly(deleteThis);
    }
}
