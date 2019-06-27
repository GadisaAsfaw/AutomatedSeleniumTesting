package Portal;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.stream.Collectors;

public class GenerateGradeReport {
    static Fillo fillo = new Fillo();
    static Connection connection;
    final static String TableOrsheetName="Sheet2";
    final static String excelFilePath="D:\\Book1.xlsx";
    public static void main(String[] args)throws InterruptedException, FilloException {

        System.setProperty("webdriver.chrome.driver","D:\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://portal.aait.edu.et");

        //login.
        WebElement UN = driver.findElement(By.id("UserName"));
        UN.sendKeys("atr/4136/09");
        WebElement password = driver.findElement(By.id("Password"));
        password.sendKeys("WriteYourPasswordHere");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        //login

        WebElement profileLink = driver.findElement(By.xpath("/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/p[2]/a"));
        driver.get("https://portal.aait.edu.et/Grade/GradeReport");
        Instant starttime = Instant.now(); // get current time

            //Iterate rows
            for (int i = 3; i < 37; i++) // skip header rows
            {
                 //jump header
                if(!isHeader(i)){

                WebElement Table = driver.findElement(
                        By.xpath("/html/body/div[2]/div/div[2]/div[1]/div/div/table/tbody/tr[" + i + "]"));

                //Map columns/ limit number of td to 6 to jumping assessment Button
                String tdData = Table.findElements(By.xpath(".//td")).stream().limit(6)
                        .map(e -> e.getText()) //get td data
                        .map(e -> e.replace("'", "") //format data
                                .replace("\"", ""))//format
                        .map(e -> e.trim()) //trim string
                        .collect(Collectors.joining("','"));  //format

                System.out.println(tdData);
                ExportToExcel("INSERT INTO " + TableOrsheetName +
                        "(Number,CourseTitle,Code,CreditHour,ECTS,Grade)VALUES('" + tdData + "') ");


                System.out.println();



              }
            }


        if (connection != null) {
            connection.close();
        }

        Instant endtime = Instant.now(); // get end time
        Duration totalTimeSpent = Duration.between(starttime, endtime); // calculate
        System.out.println("Total Time Spent in Seconds: " + totalTimeSpent.getSeconds());
        driver.close();

    }

    public static void ExportToExcel(String Query) throws FilloException {
        if (connection == null) {
            connection = fillo.getConnection(excelFilePath);
        }
        System.out.println(Query);
        connection.executeUpdate(Query);
    }

    // check if Header row
    public static Boolean isHeader(int i){
        Boolean isIt =false;
        int[] a ={1,2,8,9,15,16,22,23,29,30,37};
        for(int j=0;j<11;j++){
            if(a[j]==i){ isIt=true;}
        }
        return isIt;
    }
}
