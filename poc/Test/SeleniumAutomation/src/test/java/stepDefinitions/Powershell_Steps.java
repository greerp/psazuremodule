package stepDefinitions;

import cucumber.api.PendingException;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import utils.FileHelpers;
import utils.SystemOuputHelper;

import java.io.*;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import utils.PropertiesManager;

public class Powershell_Steps {

    /* Common Variables */
    PrintStream old;
    BufferedReader stdout;
    PropertiesManager props = PropertiesManager.getInstance();


    @Given("^I want to check azure details using the \"([^\"]*)\" command$")
    public void I_want_to_check_azure_details_using_the_command(String scriptCommand) throws Throwable {

        //Create new copy of template powershell script
        new FileHelpers().Azure_RM_ScriptFile_Initializer();

        //Append scriptCommand from feature step to end of temporary powershell script
        try
        {
            String filename= props.getValue("destinationPowershellScriptFile");
            FileWriter fw = new FileWriter(filename,true); //the true tells it to append the new data on the end of the file
            fw.write("\n" + scriptCommand);//actually appends the string to the file
            fw.close();
        }
        catch(IOException ioe)
        {
            System.err.println("IOException: " + ioe.getMessage());
        }
    }


    @When("^I check Azure using the Azure Resource Management script$")
    public void I_check_Azure_using_the_Azure_Resource_Management_script() throws Throwable {
        //Run Powershell with the temporary script file to be executed as a parameter
        String command = "powershell.exe -ExecutionPolicy Bypass -File \"" + props.getValue("destinationPowershellScriptFile") + "\"";

        //Tell the system to output from powershell in a way we can read
        new SystemOuputHelper().switchSystemOutputForPowershell();

        // Executing the command
        Process powerShellProcess = Runtime.getRuntime().exec(command);

        // Getting the results
        powerShellProcess.getOutputStream().close();
        String line;
        System.out.println("Standard Output:");
        stdout = new BufferedReader(new InputStreamReader(
                powerShellProcess.getInputStream()));
        while ((line = stdout.readLine()) != null) {
            System.out.println(line);
        }
    }


    @Then("^I expect the output to include \"([^\"]*)\"$")
    public void I_expect_the_output_to_include(String expected) throws Throwable {
        String actual = SystemOuputHelper.baos.toString();

        assertTrue(actual.contains(expected));

        // *** Return to original config
        System.setOut(old);
        stdout.close();

        //Delete the created script file
        // TODO - Move to suite tear down?
        FileHelpers.Azure_RM_ScriptFile_Delete();
    }


    @Then("^I expect the output to not include \"([^\"]*)\"$")
    public void I_expect_the_output_to_not_include(String expected) throws Throwable {

        String actual = SystemOuputHelper.baos.toString();

        assertFalse(actual.contains(expected));

        // *** Return to original config
        System.setOut(old);
        stdout.close();

        //Delete the created script file
        // TODO - Move to suite tear down?
        FileHelpers.Azure_RM_ScriptFile_Delete();
    }


    @Then("^I expect the subscription ID to be \"([^\"]*)\"$")
    public void I_expect_the_subscription_ID_to_be(String subscriptionID) throws Throwable {
        // Express the Regexp above with the code you wish you had
        String expected = subscriptionID;
        String actual = SystemOuputHelper.baos.toString();

        assertTrue(actual.contains(expected));

        // *** Return to original config
        System.setOut(old);
        stdout.close();

        //Delete the created script file
        // TODO - Move to suite tear down?
        FileHelpers.Azure_RM_ScriptFile_Delete();
    }

    @Then("^I expect the user list to include \"([^\"]*)\" and \"([^\"]*)\" and \"([^\"]*)\"$")
    public void I_expect_the_user_list_to_include_and_and(String name1, String name2, String name3) throws Throwable {
        // Express the Regexp above with the code you wish you had
        String actual = SystemOuputHelper.baos.toString();

        assertTrue(actual.contains(name1) && actual.contains(name2) && actual.contains(name3));

        // *** Return to original config
        System.setOut(old);
        stdout.close();

        //Delete the created script file
        // TODO - Move to suite tear down?
        FileHelpers.Azure_RM_ScriptFile_Delete();
    }

}


