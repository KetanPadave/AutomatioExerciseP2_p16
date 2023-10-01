package Tests;


import TestPages.Page;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

public class AutomationExercise extends BaseTest{
    @Test
    public void RegisterUser() throws IOException {
            Page page = new Page(driver);
            page.NewUserSignup();
            page.DeleteAccount();
    }

    @Test
    public void LoginUserCorrect() throws IOException {
            Page page = new Page(driver);
            page.LoginD();
            page.DeleteAccount();

    }
    @Test
    public void LoginUserInCorrect() throws IOException {
                Page page = new Page(driver);
                page.LoginIncorrect();
                boolean displayed1 = page.VerifyIncorrectLoginCredential.isDisplayed();
                Assert.assertEquals(displayed1, true);

    }
    @Test
    public void LogOutUser() throws IOException {

                Page page = new Page(driver);
                page.Login();
                page.ClickOnLogout.click();
                boolean displayed2 = page.VerifyLoginToYourAccount.isDisplayed();
                Assert.assertEquals(displayed2, true);
    }
    @Test
    public void RegisterWithExistingEmail() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(4);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String Name = row.getCell(1).getStringCellValue();
            String Email = row.getCell(2).getStringCellValue();
            String expectedHomepageTitle = row.getCell(3).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.Signup.click();
                boolean displayed = page.VerifyLoginToYourAccount.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.EnterName.sendKeys(Name);
                page.EnterEmail.sendKeys(Email);
                page.NewSignup.click();
                boolean displayed1 = page.VerifyEmailAlreadyExist.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }
    @Test
    public void ContactUs() throws IOException {
                Page page = new Page(driver);
                page.ContactUs();
                page.ClickOnHomeButton.click();
    }
    @Test
    public void VerifyTestCasePage() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.ClickOnTestCase.click();
                boolean displayed = page.VerifyTestCase.isDisplayed();
                Assert.assertEquals(displayed, true);
        }
    }

    @Test
    public void VerifyProductDetailPage() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.ClickOnProduct.click();
                boolean displayed = page.VerifyAllProduct.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.ClickOnViewProduct.click();
                boolean displayed1 = page.VerifyProductName.isDisplayed();
                Assert.assertEquals(displayed1, true);
                boolean displayed2 = page.VerifyProductCategory.isDisplayed();
                Assert.assertEquals(displayed2, true);
                boolean displayed3 = page.VerifyProductPrice.isDisplayed();
                Assert.assertEquals(displayed3, true);
                boolean displayed4 = page.VerifyProductAvailibility.isDisplayed();
                Assert.assertEquals(displayed4, true);
                boolean displayed5 = page.VerifyProductCondition.isDisplayed();
                Assert.assertEquals(displayed5, true);
                boolean displayed6 = page.VerifyProductBrand.isDisplayed();
                Assert.assertEquals(displayed6, true);
        }
    }

    @Test
    public void SearchProductPage() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            String searchName = row.getCell(2).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.ClickOnProduct.click();
                boolean displayed = page.VerifyAllProduct.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.ProductSearch.sendKeys(searchName);
                page.ProductSearchSubmit.click();
                boolean displayed1 = page.VerifySearchProduct.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }

    @Test
    public void VerifySubscriptionPage() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            String subscriptionEmail = row.getCell(3).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                boolean displayed = page.VerifySubscription.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.SubscriptionEmail.sendKeys(subscriptionEmail);
                page.ClickOnSubscribe.click();
                boolean displayed1 = page.VerifySuccessfulSubscribe.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }

    @Test
    public void VerifySubscriptionCartPage() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            String subscriptionEmail = row.getCell(3).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.ClickOnCart.click();
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                boolean displayed = page.VerifySubscription.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.SubscriptionEmail.sendKeys(subscriptionEmail);
                page.ClickOnSubscribe.click();
                boolean displayed1 = page.VerifySuccessfulSubscribe.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }

    @Test
    public void AddProduct() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.ClickOnProduct.click();
                boolean displayed = page.VerifyAllProduct.isDisplayed();
                Assert.assertEquals(displayed, true);
                Actions actions = new Actions(driver);
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scroll(200, 0);");
                actions.moveToElement(page.HoverOverFirstProduct).perform();
                page.ClickOnAddFirstProduct.click();
                page.ClickOnContinueShopping.click();
                actions.moveToElement(page.HoverOverSecondProduct).perform();
                page.ClickOnAddSecondProduct.click();
                page.ClickOnViewCart.click();
                boolean displayed1 = page.VerifyFirstProductInCart.isDisplayed();
                Assert.assertEquals(displayed1, true);
                boolean displayed2 = page.VerifySecondProductInCart.isDisplayed();
                Assert.assertEquals(displayed2, true);
                boolean displayed3 = page.VerifyFirstProductPrice.isDisplayed();
                Assert.assertEquals(displayed3, true);
                boolean displayed4 = page.VerifyFirstProductQuantity.isDisplayed();
                Assert.assertEquals(displayed4, true);
                boolean displayed5 = page.VerifyFirstProductTotalPrice.isDisplayed();
                Assert.assertEquals(displayed5, true);
                boolean displayed6 = page.VerifySecondProductPrice.isDisplayed();
                Assert.assertEquals(displayed6, true);
                boolean displayed7 = page.VerifySecondProductQuantity.isDisplayed();
                Assert.assertEquals(displayed7, true);
                boolean displayed8 = page.VerifySecondProductTotalPrice.isDisplayed();
                Assert.assertEquals(displayed8, true);
        }
    }

    @Test
    public void VerifyProductQuantity() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            double numericValue1 = row.getCell(4).getNumericCellValue();
            int intValue1 = (int) numericValue1;
            String quantity = String.valueOf(intValue1);
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                Actions actions = new Actions(driver);
                actions.moveToElement(page.ClickOnViewProduct).click().perform();
                boolean displayed1 = page.VerifyHomepageProductDetails.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.HomepageProductQuantity.clear();
                page.HomepageProductQuantity.sendKeys(quantity);
                page.HomepageProductAddToCart.click();
                page.ClickOnViewCart.click();
                String actualQuantity = page.VerifyHomepageProductQuantity.getText();
                Assert.assertEquals(quantity, actualQuantity);
        }
    }

    @Test
    public void PlaceOrderRegWhileCheckout() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(7);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.ClickOnCart.click();
                boolean displayed1 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.ClickOnProceedToCheckout.click();
                page.ClickOnRegLogin.click();
                page.NewUserSignup1();
                page.Payment();
                page.DeleteAccount();
        }
    }

    @Test
    public void PlaceOrderRegBeforeCheckout() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(7);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.NewUserSignup1();
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.Payment();
                page.DeleteAccount();
        }
    }

    @Test
    public void PlaceOrderLoginBeforeCheckout() throws IOException {
                Page page = new Page(driver);
                page.Login();
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.Payment();
                page.DeleteAccount();
    }

    @Test
    public void RemoveProductFromCart() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(8);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.ClickOnCart.click();
                boolean displayed1 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.RemoveFirstProductFromCart.click();
                page.RemoveSecondProductFromCart.click();
                boolean displayed2 = page.VerifyCartEmpty.isDisplayed();
                Assert.assertEquals(displayed2, true);
        }
    }

    @Test
    public void ViewCategoryProduct() throws IOException {

                Page page = new Page(driver);
                boolean displayed1 = page.VerifyCategory.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.ClickOnWomensCategory.click();
                page.ClickOnWomensDress.click();
                boolean displayed2 = page.VerifyWomenDressProductPage.isDisplayed();
                Assert.assertEquals(displayed2, true);
                page.ClickOnMensCategory.click();
                page.ClickOnMensTshirts.click();
                boolean displayed3 = page.VerifyMenTShirtProductPage.isDisplayed();
                Assert.assertEquals(displayed3, true);
    }

    @Test
    public void ViewCartBrandProduct() throws IOException {

            Page page = new Page(driver);
            page.ClickOnProduct.click();
            boolean displayed1 = page.VerifyBrand.isDisplayed();
            Assert.assertEquals(displayed1, true);
            page.ClickOnFirstBrand.click();
            boolean displayed2 = page.VerifyFirstBrandPage.isDisplayed();
            Assert.assertEquals(displayed2, true);
            boolean displayed3 = page.VerifyFirstBrandProduct.isDisplayed();
            Assert.assertEquals(displayed3, true);
            page.ClickOnSecondBrand.click();
            boolean displayed4 = page.VerifySecondBrandPage.isDisplayed();
            Assert.assertEquals(displayed4, true);
    }

    @Test
    public void VerifyCartAfterLogin() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(9);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String productToSearch = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                page.ClickOnProduct.click();
                boolean displayed1 = page.VerifyAllProductPage.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.SearchProduct.sendKeys(productToSearch);
                page.ClickOnSearch.click();
                boolean displayed2 = page.VerifySearch.isDisplayed();
                Assert.assertEquals(displayed2, true);
                boolean displayed3 = page.VerifyFirstSearchProduct.isDisplayed();
                Assert.assertEquals(displayed3, true);
                page.AddFirstSearchProduct.click();
                page.ClickOnContinueShopping.click();
                page.AddSecondSearchProduct.click();
                page.ClickOnContinueShopping.click();
                page.ClickOnCart.click();
                boolean displayed4 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed4, true);
                page.Signup.click();
                page.Login();
                page.ClickOnCart.click();
                boolean displayed5 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed5, true);
        }
    }

    @Test
    public void AddReview() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(10);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String name = row.getCell(1).getStringCellValue();
            String email = row.getCell(2).getStringCellValue();
            String review = row.getCell(3).getStringCellValue();
            Page page = new Page(driver);
            page.ClickOnProduct.click();
                boolean displayed1 = page.VerifyAllProductPage.isDisplayed();
                Assert.assertEquals(displayed1, true);
            page.ClickOnViewProduct.click();
            boolean displayed2 = page.VerifyWriteReview.isDisplayed();
            Assert.assertEquals(displayed2, true);
            page.ReviewName.sendKeys(name);
            page.ReviewEmail.sendKeys(email);
            page.AddReview.sendKeys(review);
            page.SubmitReview.click();
            boolean displayed4 = page.VerifySubmitReview.isDisplayed();
            Assert.assertEquals(displayed4, true);
        }
    }

    @Test
    public void AddRecomendedItem() throws IOException {
            Page page = new Page(driver);
            Actions actions = new Actions(driver);
            actions.moveToElement(page.VerifyRecomendedProduct).perform();
            boolean displayed1 = page.VerifyRecomendedProduct.isDisplayed();
            Assert.assertEquals(displayed1, true);
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
            wait.until(ExpectedConditions.visibilityOf(page.AddRecomendedProduct1ToCart));
            wait.until(ExpectedConditions.elementToBeClickable(page.AddRecomendedProduct1ToCart));
            page.AddRecomendedProduct1ToCart.click();
            page.ClickOnContinueShopping.click();
            page.ClickOnCart.click();
            boolean displayed2 = page.VerifyFirstProductInCart.isDisplayed();
            Assert.assertEquals(displayed2, true);
    }

    @Test
    public void VerifyDetailsCheckout() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(11);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String AddressName = row.getCell(20).getStringCellValue();
            String Companyname = row.getCell(21).getStringCellValue();
            String address = row.getCell(24).getStringCellValue();
                Page page = new Page(driver);
                page.NewUserSignup();
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.ClickOnCart.click();
                boolean displayed1 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.ClickOnProceedToCheckout.click();
                String actualDeliveryAddressName = page.VerifyDeliveryAddressName.getText();
                Assert.assertEquals(AddressName, actualDeliveryAddressName);
                String actualBillingAddressName = page.VerifyBillingAddressName.getText();
                Assert.assertEquals(AddressName, actualBillingAddressName);
                String actualDeliveryCompanyName = page.VerifyDeliveryCompanyName.getText();
                Assert.assertEquals(Companyname, actualDeliveryCompanyName);
                String actualBillingCompanyName = page.VerifyBillingCompanyName.getText();
                Assert.assertEquals(Companyname, actualBillingCompanyName);
                String actualDeliveryAddress = page.VerifyDeliveryAddress.getText();
                Assert.assertEquals(address, actualDeliveryAddress);
                String actualBillingAddress = page.VerifyBillingAddress.getText();
                Assert.assertEquals(address, actualBillingAddress);
                page.DeleteAccount();
        }
    }

    @Test
    public void DownloadInvoice() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(11);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                page.AddHomepageProduct1ToCart.click();
                page.ClickOnContinueShopping.click();
                page.AddHomepageProduct2ToCart.click();
                page.ClickOnContinueShopping.click();
                page.ClickOnCart.click();
                boolean displayed1 = page.VerifyCart.isDisplayed();
                Assert.assertEquals(displayed1, true);
                page.ClickOnProceedToCheckout.click();
                page.ClickOnRegLogIn.click();
                page.NewUserSignup1();
                page.Payment();
                page.ClickOnDownloadInvoice.click();
                String filePath = "C:\\Users\\Admin\\Downloads\\invoice.txt";
                File downloadedFile = new File(filePath);
                Assert.assertTrue(downloadedFile.exists(), "Downloaded file does not exist.");
                page.ClickOnContinue.click();
                page.DeleteAccount();
        }
    }

    @Test
    public void VerifyScrollUpUsingArrow() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                boolean displayed = page.VerifySubscription.isDisplayed();
                Assert.assertEquals(displayed, true);
                page.ClickOnRightSideUpArrow.click();
                boolean displayed1 = page.VerifyFullFledgeText.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }

    @Test
    public void VerifyScrollUpWithoutUsingArrow() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(6);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
                Page page = new Page(driver);
                String actualHomepageTitle = driver.getTitle();
                Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                boolean displayed = page.VerifySubscription.isDisplayed();
                Assert.assertEquals(displayed, true);
                js.executeScript("window.scrollTo(0, -document.body.scrollHeight);");
                boolean displayed1 = page.VerifyFullFledgeText.isDisplayed();
                Assert.assertEquals(displayed1, true);
        }
    }
}

