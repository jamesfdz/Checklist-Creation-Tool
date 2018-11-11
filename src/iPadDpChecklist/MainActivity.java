package iPadDpChecklist;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import static java.lang.Integer.parseInt;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import javax.swing.JOptionPane;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class MainActivity {
	
	 public static final String URL = "https://login.veevavault.com/";
	 public static Logger logger = Logger.getLogger("MyLog");   
	 public static FileHandler fh;
	 public static final List<String> slideNames = new ArrayList<String>();
	 public static final List<String> slideLinks = new ArrayList<String>();
	 public static final List<String> slideBinderId = new ArrayList<String>();

	public static void main(String[] args) throws MalformedURLException, FileNotFoundException, IOException {
		
		List<String> allfileSizes = new ArrayList<String>();
		
        try{
        //open browser
        System.setProperty("webdriver.chrome.driver","chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get(URL);
        WebElement usernameField = driver.findElement(By.id("j_username"));
        
        //takes input from excel
        FileInputStream fs = new FileInputStream("input.xlsx");
        XSSFWorkbook inputWorkbook = new XSSFWorkbook(fs);
        Sheet inputSheet =  inputWorkbook.getSheet("Sheet1");
        Row row1 = inputSheet.getRow(0);
        Cell cell0 = row1.getCell(0);
        String username = cell0.getStringCellValue();
        System.out.println(username);
        usernameField.sendKeys(username);
        usernameField.sendKeys(Keys.ENTER);
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.id("j_password")));
        
        WebElement passwordField = driver.findElement(By.id("j_password"));
        
        Row row2 = inputSheet.getRow(1);
        Cell cell1 = row2.getCell(0);
        String password = cell1.getStringCellValue();
        passwordField.sendKeys(password);
        System.out.println(password);
        passwordField.sendKeys(Keys.ENTER);      
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='noItemsFound vv_no_results']")));
        
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        
        WebElement searchField = driver.findElement(By.id("search_main_box"));
        Row row3 = inputSheet.getRow(2);
        Cell cell2 = row3.getCell(0);
        String searchString = cell2.getStringCellValue();
        
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        
        searchField.sendKeys(searchString);
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        searchField.sendKeys(Keys.ENTER);
        System.out.println(searchString);
        
        inputWorkbook.close();
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@title='Detail View']")));
        WebElement detailView = driver.findElement(By.xpath("//*[@title='Detail View']"));
        detailView.click();
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@class='docLink vv_doc_title_link'])[1]")));
        
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        
        WebElement presentationTitle = driver.findElement(By.xpath("(//*[@class='docLink vv_doc_title_link'])[1]"));
        presentationTitle.click();
        
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        
        new WebDriverWait(driver, 2000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@attrkey='name']")));
        WebElement PresentationNameElement = driver.findElement(By.xpath("//*[@attrkey='name']"));
        WebElement PresentationBinderElement = driver.findElement(By.xpath("//*[@attrkey='DocumentNumber']"));
        WebElement PresentationZincIdElement = driver.findElement(By.xpath("//*[@attrkey='zincId']"));
        
        String PresentationNameText = PresentationNameElement.getText();
        String PresentationBinderText = PresentationBinderElement.getText();
        String PresentationZincIdText = PresentationZincIdElement.getText();
        
        System.out.println(PresentationNameText);
        System.out.println(PresentationBinderText);
        System.out.println(PresentationZincIdText);
        
        
        FileInputStream fis = new FileInputStream("output.xlsx");
        XSSFWorkbook outputWorkbook = new XSSFWorkbook(fis);
        Sheet outputSheet = outputWorkbook.getSheet("Checklist1");
        CreationHelper createHelper = outputWorkbook.getCreationHelper();
        XSSFCellStyle hlinkstyle = outputWorkbook.createCellStyle();
        XSSFFont hlinkfont = outputWorkbook.createFont();
        hlinkfont.setUnderline(XSSFFont.U_SINGLE);
        hlinkfont.setColor(HSSFColor.BLUE.index);
        hlinkstyle.setFont(hlinkfont);
        Cell NameCell = null;
        Cell BinderCell = null;
        Cell ZincIdCell = null;
        
        NameCell = outputSheet.getRow(12).getCell(1);
        NameCell.setCellValue(PresentationNameText);
        BinderCell = outputSheet.getRow(14).getCell(1);
        BinderCell.setCellValue(PresentationBinderText);
        ZincIdCell = outputSheet.getRow(22).getCell(1);
        ZincIdCell.setCellValue(PresentationZincIdText);
        
        String presentationUrl = driver.getCurrentUrl();
        XSSFHyperlink presentationUrlLink = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
        presentationUrlLink.setAddress(presentationUrl);
        NameCell.setHyperlink(presentationUrlLink);
        NameCell.setCellStyle(hlinkstyle);
        
        XSSFHyperlink presentationIdUrlLink = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
        presentationIdUrlLink.setAddress(presentationUrl);
        BinderCell.setHyperlink(presentationIdUrlLink);
        BinderCell.setCellStyle(hlinkstyle);
        
        
        WebElement multichannelProperties = driver.findElement(By.xpath("//*[@key='multichannelProperties']"));
        multichannelProperties.click();
        
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@attrkey='crmPresentationId_b']")));
        WebElement PresentationIdElement = driver.findElement(By.xpath("//*[@attrkey='crmPresentationId_b']"));
        String PresentationIdText = PresentationIdElement.getText();
        System.out.println(PresentationIdText);
        
        Cell PidCell = null;
        PidCell = outputSheet.getRow(13).getCell(1);
        PidCell.setCellValue(PresentationIdText);
        
        String pidUrl = driver.getCurrentUrl();
        XSSFHyperlink pidUrlLink = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
        pidUrlLink.setAddress(pidUrl);
        PidCell.setHyperlink(pidUrlLink);
        PidCell.setCellStyle(hlinkstyle);
        
        WebElement ClmProperties = driver.findElement(By.xpath("//*[@key='clmProperties']"));
        ClmProperties.click();
        
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@attrkey='crmEndDate_b']")));
        WebElement PresentationEndDateElement = driver.findElement(By.xpath("//*[@attrkey='crmEndDate_b']"));
        
        
        String PresentationEndDate = PresentationEndDateElement.getText();
        
        
        System.out.println(PresentationEndDate);
        
        
        Cell endDateCell = null;
        
        endDateCell = outputSheet.getRow(11).getCell(1);
        endDateCell.setCellValue(PresentationEndDate);
        
        WebElement firstScrolling = driver.findElement(By.id("search_main_box"));
        
        System.out.println("Going up to get page number");
        ((JavascriptExecutor)driver).executeScript("arguments[0].scrollIntoView();", firstScrolling);
        
        getPageData(driver);
        
        WebElement pages = driver.findElement(By.xpath("(//span[@class='vv_float_left'])[1]"));
	    String pagesText = pages.getText();
	    System.out.println("How many pages: "+pagesText);
	      
	    if(pagesText != "of 1") {
	    	System.out.println("More than one pages");
	     	String[] pageNumberPart = pagesText.split(" ");
	     	System.out.println("partion two:"+pageNumberPart[1]);
	     	int pageNum = parseInt(pageNumberPart[1]);
	      	
	      	for(int i = 1; i < pageNum; i++) {
	      		driver.findElement(By.xpath("//*[@class='vpage_next vv_button vv_float_left']")).click();
	      		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	      		getPageData(driver);
	      	}
	    }
	    
	    //TODO run for each loop and add slideNames in excel sheet
	    Cell docTitle = null;
	    Cell docNumber = null;
	    
	    int initialRowNumber = 23;
	      for(String slideName : slideNames){
	          docTitle = outputSheet.getRow(initialRowNumber).getCell(0);
	          docTitle.setCellValue(slideName);          
	          initialRowNumber++;
	      }
	    
	    //TODO run for each loop and add slideNums in excel sheet
	    int initialRowNumberForBinder = 23;
        for(String slideNum : slideBinderId){
            
            docNumber = outputSheet.getRow(initialRowNumberForBinder).getCell(1);
            docNumber.setCellValue(slideNum);
            initialRowNumberForBinder++;
            
        }
	    
	    //TODO run for each loop and add slideLinks to slideNames & slideNums in excel sheet
        int initialRow = 23;
        for(String slideLink : slideLinks){
            System.out.println(slideLink);
            Cell binder_cell_0 = outputSheet.getRow(initialRow).getCell(0);
            XSSFHyperlink binder_cell_0_link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
            binder_cell_0_link.setAddress(slideLink);
            binder_cell_0.setHyperlink(binder_cell_0_link);
            binder_cell_0.setCellStyle(hlinkstyle);
            
            Cell binder_cell_1 = outputSheet.getRow(initialRow).getCell(1);
            XSSFHyperlink binder_cell_1_link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
            binder_cell_1_link.setAddress(slideLink);
            binder_cell_1.setHyperlink(binder_cell_1_link);
            binder_cell_1.setCellStyle(hlinkstyle);
            
            initialRow++;
        }
	    
	    //TODO continue the rest of the code from shared resources flow
        
        WebElement firstDocElementScrolling = driver.findElement(By.id("search_main_box"));
        
        System.out.println("Going up");
        ((JavascriptExecutor)driver).executeScript("arguments[0].scrollIntoView();", firstDocElementScrolling);
        
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@class='docNameLink vv_doc_title_link veevaTooltipBound'])[1]")));
        WebElement firstDocElement = driver.findElement(By.xpath("(//*[@class='docNameLink vv_doc_title_link veevaTooltipBound'])[1]"));
        firstDocElement.click();
        
        new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@key='related_shared_resource__pm']")));
        
        WebElement relatedSharedElementCheck = driver.findElement(By.xpath("//*[@key='related_shared_resource__pm']//span[@class='count vv_section_count']"));
        String relatedSharedElementCheckString = relatedSharedElementCheck.getText();
        System.out.println(relatedSharedElementCheckString);
        
        if(relatedSharedElementCheckString.equals("(0)")){
            System.out.println("No shared Resource");
        }else{
            WebElement relatedSharedElement = driver.findElement(By.xpath("//*[@key='related_shared_resource__pm']"));
            relatedSharedElement.click();
            new WebDriverWait(driver, 3000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='vv_rd_c2']//a[@class='docName doc_link veevaTooltipBound']")));
            WebElement sharedResource = driver.findElement(By.xpath("//*[@class='vv_rd_c2']//a[@class='docName doc_link veevaTooltipBound']"));
            sharedResource.click();

            driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
            new WebDriverWait(driver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='prev_page']")));
            WebElement sharedResourceNameElement = driver.findElement(By.xpath("//*[@attrkey='name']"));
            String sharedNameText = sharedResourceNameElement.getText();
            WebElement sharedDocNumber = driver.findElement(By.xpath("//*[@attrkey='DocumentNumber']"));
            String sharedNumText = sharedDocNumber.getText();
            WebElement sharedRenditions = driver.findElement(By.xpath("//*[@key='renditions']"));
            sharedRenditions.click();

            new WebDriverWait(driver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@class='fileSize'])[1]")));
            WebElement sharedFileSizeEm = driver.findElement(By.xpath("(//*[@class='fileSize'])[1]"));
            String sharedFileSizeText = sharedFileSizeEm.getText();

            System.out.println("shared: " + sharedNameText);
            System.out.println("sharedNum: " + sharedNumText);
            System.out.println("shared File Size" + sharedFileSizeText);

            Cell sharedName = null;
            Cell sharedNum = null;

            sharedName = outputSheet.getRow(18).getCell(1);
            sharedName.setCellValue(sharedNameText);

            sharedNum = outputSheet.getRow(19).getCell(1);
            sharedNum.setCellValue(sharedNumText);

            String currentURL = driver.getCurrentUrl();
            XSSFHyperlink sharedLinkForName = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
            sharedLinkForName.setAddress(currentURL);
            sharedName.setHyperlink(sharedLinkForName);

            String currentURL1 = driver.getCurrentUrl();
            XSSFHyperlink sharedLinkForNum = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
            sharedLinkForNum.setAddress(currentURL1);
            sharedNum.setHyperlink(sharedLinkForNum);
            
            allfileSizes.add(sharedFileSizeText);
        }
        
        driver.close();
        
        
        for(String slideLink : slideLinks){
            System.setProperty("webdriver.chrome.driver","chromedriver.exe");
            WebDriver newDriver = new ChromeDriver();
            newDriver.manage().window().maximize();
            newDriver.get(slideLink);
            
            new WebDriverWait(newDriver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.id("j_username")));
            
            WebElement newusernameField = newDriver.findElement(By.id("j_username"));
            newusernameField.sendKeys(username);
            newusernameField.sendKeys(Keys.ENTER);
            
            
            new WebDriverWait(newDriver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.id("j_password")));
            WebElement newpasswordField = newDriver.findElement(By.id("j_password"));
            
            newpasswordField.sendKeys(password);
            newpasswordField.sendKeys(Keys.ENTER);
            
            new WebDriverWait(newDriver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@key='renditions']")));
            
            WebElement renditions = newDriver.findElement(By.xpath("//*[@key='renditions']"));
            renditions.click();
            
            new WebDriverWait(newDriver, 5000).until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@class='fileSize'])[1]")));
            
            WebElement fileSizeEm = newDriver.findElement(By.xpath("(//*[@class='fileSize'])[1]"));
            
            newDriver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
            
            String fileSizeTxt = fileSizeEm.getText();
            System.out.println(fileSizeTxt);
            
            allfileSizes.add(fileSizeTxt);
            
            newDriver.close();
        }
        
        String searchKB = "KB";
        List<Float> MBfileSizes = new ArrayList<Float>();
        List<Float> KBfileSizes = new ArrayList<Float>();
        
        for(String fileSize : allfileSizes){
            if(fileSize.contains(searchKB)){
                //Spilt number and kb
                String[] partKB = fileSize.split(" ");
                System.out.println(partKB[0]);
                KBfileSizes.add(Float.parseFloat(partKB[0]));
                System.out.println("Added to KB list");
            }else{
                //spilt number and mb
                String[] partMB = fileSize.split(" ");
                System.out.println(partMB[0]);
                MBfileSizes.add(Float.parseFloat(partMB[0]));
                System.out.println("Added to MB list");
            }
        }
        
        for(Float KBfileSize : KBfileSizes){
            Float convertedtoMB = KBfileSize / 1024;
            BigDecimal bd = new BigDecimal(convertedtoMB);
            bd = bd.setScale(2, BigDecimal.ROUND_HALF_UP);
            Float finalValue = bd.floatValue();
            System.out.println(finalValue);
            MBfileSizes.add(finalValue);
            System.out.println("Converted from KB to MB and added to list");
        }
        
        float sum = 0;
        for(Float MBfileSize : MBfileSizes){
            sum += MBfileSize;
        }
        
        BigDecimal bd1 = new BigDecimal(sum);
        bd1 = bd1.setScale(2, BigDecimal.ROUND_HALF_UP);
        Float finalValue1 = bd1.floatValue();
        String SumOfAll = Float.toString(finalValue1);
        System.out.println(SumOfAll);
        
        Cell fileSizeCell = null;
        
        fileSizeCell = outputSheet.getRow(21).getCell(1);
        fileSizeCell.setCellValue(SumOfAll + " MB");
        
        System.out.println("Final Answer: " + SumOfAll + " MB");

        fis.close();
        FileOutputStream output = new FileOutputStream(new File("output.xlsx"));
        outputWorkbook.write(output);
        output.close();
        
        outputWorkbook.close();
        
        JOptionPane.showMessageDialog(null, "Completed Successfully");
        
        }catch(Exception e){
             fh = new FileHandler("LogFile.log");  
             logger.addHandler(fh);
             SimpleFormatter formatter = new SimpleFormatter();  
             fh.setFormatter(formatter);  
 
            // the following statement is used to log any messages
            logger.info("" + e.getLocalizedMessage());
        }        
    }

	private static void getPageData(WebDriver driver) {
		List<WebElement> slideNamesElm = driver.findElements(By.xpath("//*[@class='docNameLink vv_doc_title_link veevaTooltipBound']"));
		int slideElmCount = slideNamesElm.size();
		for(int i = 1; i <= slideElmCount; i++) {
			WebElement slideNameElm = driver.findElement(By.xpath("(//*[@class='docNameLink vv_doc_title_link veevaTooltipBound'])["+i+"]"));
			String slideName = slideNameElm.getText();
			slideNames.add(slideName);
			String slideLink = slideNameElm.getAttribute("href");
			slideLinks.add(slideLink);
			WebElement slideBinderIdElm = driver.findElement(By.xpath("(//*[@class='docNumber vv_doc_number'])["+i+"]"));
			String slideBinderIdText = slideBinderIdElm.getText();
			slideBinderId.add(slideBinderIdText);
		}
	}
}
