package stepDefinitions;

import cucumber.api.PendingException;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import interaction_objects.helloWorld.HelloWorld;

import static org.bouncycastle.util.Arrays.areEqual;
import static org.junit.Assert.assertEquals;


public class UI_helloWorld_steps {

    private HelloWorld helloWorld = new HelloWorld();


    @Given("^I navigate to Hello World")
    public void I_navigate_to_hello_world() {

        helloWorld.navigateTo();

    }

    @When("^I check the body text$")
    public void I_check_the_body_text() throws Throwable {
        helloWorld.getTopText();
    }

    @Then("^I expect the hello world landing page to say \"([^\"]*)\"$")
    public void I_expect_the_hello_world_landing_page_to_say(String expected_text) throws Throwable {
        assertEquals(expected_text, helloWorld.returnTopText());
    }
}
