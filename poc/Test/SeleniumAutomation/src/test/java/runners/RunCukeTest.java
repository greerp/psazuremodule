package runners;

import cucumber.api.CucumberOptions;
import cucumber.api.junit.Cucumber;
import org.junit.runner.RunWith;

@RunWith(Cucumber.class)
@CucumberOptions(

    tags = {"~@wip"},
    features = "src/test/java/features",
    glue = "stepDefinitions",
    dryRun = false,
    monochrome = true,
    plugin = { "progress", "json:target/Cucumber.json", "junit:target/junit.xml"}

)

public class RunCukeTest {

}
