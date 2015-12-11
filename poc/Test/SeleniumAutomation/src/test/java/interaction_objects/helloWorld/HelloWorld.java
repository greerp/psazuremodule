package interaction_objects.helloWorld;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import stepDefinitions.WebHooks;
import utils.PropertiesManager;


public class HelloWorld {

    /* Element Selectors */
    By TopText =  By.xpath("/html/body/h1");

    WebDriver driver = WebHooks.driver;
    WebDriverWait wait = WebHooks.wait;

    PropertiesManager props = PropertiesManager.getInstance();

    /* Common Variables*/
    public String hw_topText;

    public HelloWorld(){}

    public void navigateTo(){
        driver.get(props.getValue("helloWorldURL"));
    }

    public void navigateToHttp(){
        driver.get(props.getValue("helloWorldURL80"));
    }

    public void getTopText() {
        hw_topText = (driver.findElement((TopText)).getText());
    }

    public String returnTopText() {
        return hw_topText;
    }


}
