package stepDefinitions;

import cucumber.api.Scenario;
import cucumber.api.java.After;
import cucumber.api.java.Before;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.WebDriverWait;
import ru.stqa.selenium.factory.RemoteDriverProvider;
import ru.stqa.selenium.factory.WebDriverFactory;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;


import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

/**
 * Created by tejaw on 28/09/2015.
 */
public class WebHooks {
    public static WebDriver driver;
    public static WebDriverWait wait;
    static final Logger logger = LogManager.getLogger( WebHooks.class.getName() );


    @Before("@WEB")
    public void beforeScenario(){

        DesiredCapabilities caps;

        logger.debug("Running BEFORE 'WEB' tag");


        String BROWSERPROFILE = System.getProperty("BROWSERPROFILE");
        if ((BROWSERPROFILE == null) || (BROWSERPROFILE.equals(""))) {
            BROWSERPROFILE = "FIREFOX";
        }

        switch(BROWSERPROFILE){
            case "IE":
                System.setProperty("webdriver.ie.driver", "src/test/drivers/IEDriverServer.exe");
                caps = DesiredCapabilities.internetExplorer();
                break;
            case "CHROME":
                System.setProperty("webdriver.chrome.driver", "src/test/drivers/chromedriver.exe");
                caps = DesiredCapabilities.chrome();
                caps.setCapability("platform", "Windows 7");
                caps.setCapability("version", "45.0");
                break;
            case "FIREFOX":
                caps = DesiredCapabilities.firefox();
                break;
            case "HEADLESS":
                caps = DesiredCapabilities.htmlUnitWithJs();
                break;
            default:
                //FF is built into Se, so it will be the default
                caps = DesiredCapabilities.firefox();
                break;
        }


        driver = WebDriverFactory.getDriver(caps);
        wait = new WebDriverWait( driver, 5 );

    }

    @After("@WEB")
    public void afterScenario(Scenario scenario){
        //System.out.println("Running AFTER 'WEB' tag");

        /*if (scenario.isFailed()) {

            logger.error("@Web@After: take screenshot on failure for: " + scenario.getName().substring(0,50));

            File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);

            String currentDir = System.getProperty("user.dir");
            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());

            try {
                String newFilename = "\\target\\pickles\\screenshots\\" + timeStamp +"_"+ scenario.getName().substring(0,20) + ".png";
                newFilename = newFilename.replace(" ","_").replace("\"","_").replace(":","_");
                logger.error("Creating Screenshot: " + newFilename);
                FileUtils.copyFile(scrFile, new File( currentDir + newFilename ));
            } catch (IOException e1) {

                e1.printStackTrace();
            }

            final byte[] screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.BYTES);
            scenario.embed(screenshot, "image/png"); //stick it in the report

        }
*/
        WebDriverFactory.dismissDriver(driver);
    }



}
