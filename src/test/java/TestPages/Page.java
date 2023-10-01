package TestPages;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import static Utilities.DriverManager.driver;

public class Page {
    public Page(WebDriver driver) {
        PageFactory.initElements(driver, this);
    }
    @FindBy(xpath = "//a[contains(text(),'Signup')]")
    public WebElement Signup;
    @FindBy(xpath = "//input[contains(@name, 'name')]")
    public WebElement EnterName;
    @FindBy(xpath = "(//input[contains(@name, 'email')])[2]")
    public WebElement EnterEmail;
    @FindBy(xpath = "//button[contains(text(),'Signup')]")
    public WebElement NewSignup;
    @FindBy(xpath = "//input[@value='Mr']")
    public WebElement SelectTitle;
    @FindBy(xpath = "//input[contains(@name, 'password')]")
    public WebElement EnterPassword;
    @FindBy(xpath = "//select[contains(@name, 'days')]")
    public WebElement SelectDate;
    @FindBy(xpath = "//select[contains(@name, 'months')]")
    public WebElement SelectMonth;
    @FindBy(xpath = "//select[contains(@name, 'years')]")
    public WebElement SelectYear;
    @FindBy(xpath = "//input[contains(@name, 'newsletter')]")
    public WebElement SelectNewsLetterCheckbox;
    @FindBy(xpath = "//input[contains(@name, 'optin')]")
    public WebElement SelectOfferPartnerCheckbox;
    @FindBy(xpath = "//input[contains(@name, 'first_name')]")
    public WebElement EnterFirstName;
    @FindBy(xpath = "//input[contains(@name, 'last_name')]")
    public WebElement EnterLastName;
    @FindBy(xpath = "//input[contains(@name, 'company')]")
    public WebElement EnterCompanyName;
    @FindBy(xpath = "//input[contains(@name, 'address1')]")
    public WebElement EnterAddress1;
    @FindBy(xpath = "//input[contains(@name, 'address2')]")
    public WebElement EnterAddress2;
    @FindBy(xpath = "//select[contains(@name, 'country')]")
    public WebElement SelectCountry;
    @FindBy(xpath = "//input[contains(@name, 'state')]")
    public WebElement EnterState;
    @FindBy(xpath = "//input[contains(@name, 'city')]")
    public WebElement EnterCity;
    @FindBy(xpath = "//input[contains(@name, 'zipcode')]")
    public WebElement EnterZipcode;
    @FindBy(xpath = "//input[contains(@name, 'mobile_number')]")
    public WebElement EnterMobileNo;
    @FindBy(xpath = "//button[contains(text(),'Create Account')]")
    public WebElement ClickOnCreateAccount;
    @FindBy(xpath = "//b[contains(text(),'Account Created!')]")
    public WebElement VerifyCreateAccount;
    @FindBy(xpath = "//a[contains(text(),'Continue')]")
    public WebElement ClickOnContinue;
    @FindBy(xpath = "//b[contains(text(),'')]")
    public WebElement VerifyLoginUser;
    @FindBy(xpath = "//a[contains(@href, 'delete_account')]")
    public WebElement ClickOnDeleteAccount;
    @FindBy(xpath = "//b[contains(text(),'Account Deleted!')]")
    public WebElement VerifyDeleteAccount;
    @FindBy(xpath = "//h2[contains(text(),'Login to your account')]")
    public WebElement VerifyLoginToYourAccount;
    @FindBy(xpath = "(//input[contains(@name, 'email')])[1]")
    public WebElement EnterEmailAddress;
    @FindBy(xpath = "//input[contains(@name, 'password')]")
    public WebElement EnterLoginPassword;
    @FindBy(xpath = "//button[contains(text(),'Login')]")
    public WebElement ClickOnLogin;
    @FindBy(xpath = "//a[contains(text(),' Logged in as ')]")
    public WebElement VerifyLoginAsUsername;
    @FindBy(xpath = "//p[contains(text(),'Your email or password is incorrect!')]")
    public WebElement VerifyIncorrectLoginCredential;
    @FindBy(xpath = "//a[contains(text(),'Logout')]")
    public WebElement ClickOnLogout;
    @FindBy(xpath = "//p[contains(text(),'Email Address already exist!')]")
    public WebElement VerifyEmailAlreadyExist;
    @FindBy(xpath = "//a[contains(text(),'Contact us')]")
    public WebElement ClickOnContactUs;
    @FindBy(xpath = "//h2[contains(text(),'Get In Touch')]")
    public WebElement VerifyGetInTouch;
    @FindBy(xpath = "//input[contains(@name, 'name')]")
    public WebElement EnterNameHelp;
    @FindBy(xpath = "//input[contains(@name, 'email')]")
    public WebElement EnterEmailHelp;
    @FindBy(xpath = "//input[contains(@name, 'subject')]")
    public WebElement EnterSubjectHelp;
    @FindBy(xpath = "//textarea[contains(@name, 'message')]")
    public WebElement EnterTextAreaHelp;
    @FindBy(xpath = "//input[contains(@name, 'submit')]")
    public WebElement ClickOnSubmitHelp;
    @FindBy(xpath = "(//div[contains(text(),'Success! Your details have been submitted successfully.')])[1]")
    public WebElement VerifySuccessMsg;
    @FindBy(xpath = "//a[contains(text(),'Home')]")
    public WebElement ClickOnHomeButton;
    @FindBy(xpath = "//a[contains(text(),'Test Cases')]")
    public WebElement ClickOnTestCase;
    @FindBy(xpath = "//b[contains(text(),'Test Cases')]")
    public WebElement VerifyTestCase;
    @FindBy(xpath = "//a[contains(text(),'Products')]")
    public WebElement ClickOnProduct;
    @FindBy(xpath = "//h2[contains(text(),'All Products')]")
    public WebElement VerifyAllProduct;
    @FindBy(xpath = "(//a[contains(text(),'View Product')])[1]")
    public WebElement ClickOnViewProduct;
    @FindBy(xpath = "(//h2[contains(text(),'')])[3]")
    public WebElement VerifyProductName;
    @FindBy(xpath = "//p[contains(text(),'Category')]")
    public WebElement VerifyProductCategory;
    @FindBy(xpath = "//span[contains(text(),'Rs.')]")
    public WebElement VerifyProductPrice;
    @FindBy(xpath = "//p[contains(text(),'Stock')]")
    public WebElement VerifyProductAvailibility;
    @FindBy(xpath = "(//p[contains(text(),'')])[5]")
    public WebElement VerifyProductCondition;
    @FindBy(xpath = "(//p[contains(text(),'')])[6]")
    public WebElement VerifyProductBrand;
    @FindBy(xpath = "//input[contains(@name, 'search')]")
    public WebElement ProductSearch;
    @FindBy(xpath = "//button[contains(@id, 'submit_search')]")
    public WebElement ProductSearchSubmit;
    @FindBy(xpath = "//h2[contains(text(),'Searched Products')]")
    public WebElement VerifySearchProduct;
    @FindBy(xpath = "//h2[contains(text(),'Subscription')]")
    public WebElement VerifySubscription;
    @FindBy(xpath = "//input[contains(@id, 'susbscribe_email')]")
    public WebElement SubscriptionEmail;
    @FindBy(xpath = "//button[contains(@id, 'subscribe')]")
    public WebElement ClickOnSubscribe;
    @FindBy(xpath = "//div[contains(text(),'You have been successfully subscribed!')]")
    public WebElement VerifySuccessfulSubscribe;
    @FindBy(xpath = "//a[contains(text(),'Cart')]")
    public WebElement ClickOnCart;
    @FindBy(xpath = "(//div[contains(@class, 'productinfo text-center')])[1]")
    public WebElement HoverOverFirstProduct;
    @FindBy(xpath = "(//div[contains(@class, 'productinfo text-center')])[2]")
    public WebElement HoverOverSecondProduct;
    @FindBy(xpath = "(//a[contains(@class, 'add-to-cart')])[2]")
    public WebElement ClickOnAddFirstProduct;
    @FindBy(xpath = "(//a[contains(text(),'Add to cart')])[3]")
    public WebElement ClickOnAddSecondProduct;
    @FindBy(xpath = "//button[contains(text(),'Continue Shopping')]")
    public WebElement ClickOnContinueShopping;
    @FindBy(xpath = "//u[contains(text(),'View Cart')]")
    public WebElement ClickOnViewCart;
    @FindBy(xpath = "//a[contains(@href, 'product_details/1')]")
    public WebElement VerifyFirstProductInCart;
    @FindBy(xpath = "//a[contains(@href, 'product_details/2')]")
    public WebElement VerifySecondProductInCart;
    @FindBy(xpath = "(//p[contains(text(),'Rs.')])[1]")
    public WebElement VerifyFirstProductPrice;
    @FindBy(xpath = "(//button[contains(text(),'')])[2]")
    public WebElement VerifyFirstProductQuantity;
    @FindBy(xpath = "(//p[contains(text(),'Rs.')])[2]")
    public WebElement VerifyFirstProductTotalPrice;
    @FindBy(xpath = "(//p[contains(text(),'Rs.')])[3]")
    public WebElement VerifySecondProductPrice;
    @FindBy(xpath = "(//button[contains(text(),'')])[3]")
    public WebElement VerifySecondProductQuantity;
    @FindBy(xpath = "(//p[contains(text(),'Rs.')])[4]")
    public WebElement VerifySecondProductTotalPrice;
    @FindBy(xpath = "//div[contains(@class, 'product-information')]")
    public WebElement VerifyHomepageProductDetails;
    @FindBy(xpath = "//input[contains(@name, 'quantity')]")
    public WebElement HomepageProductQuantity;
    @FindBy(xpath = "//button[contains(@type, 'button')]")
    public WebElement HomepageProductAddToCart;
    @FindBy(xpath = "(//button[contains(text(),'')])[2]")
    public WebElement VerifyHomepageProductQuantity;
    @FindBy(xpath = "(//a[contains(@class, 'add-to-cart')])[1]")
    public WebElement AddHomepageProduct1ToCart;
    @FindBy(xpath = "(//a[contains(@class, 'add-to-cart')])[3]")
    public WebElement AddHomepageProduct2ToCart;
    @FindBy(xpath = "//li[contains(text(),'Shopping Cart')]")
    public WebElement VerifyCart;
    @FindBy(xpath = "//a[contains(text(),'Proceed To Checkout')]")
    public WebElement ClickOnProceedToCheckout;
    @FindBy(xpath = "//u[contains(text(),'Register / Login')]")
    public WebElement ClickOnRegLogin;
    @FindBy(xpath = "//h2[contains(text(),'Address Details')]")
    public WebElement VerifyAddressDetails;
    @FindBy(xpath = "//h2[contains(text(),'Review Your Order')]")
    public WebElement VerifyReviewOrder;
    @FindBy(xpath = "//textarea[contains(@name, 'message')]")
    public WebElement CommentText;
    @FindBy(xpath = "//a[contains(text(),'Place Order')]")
    public WebElement ClickOnPlaceOrder;
    @FindBy(xpath = "//input[contains(@name, 'name_on_card')]")
    public WebElement PaymentName;
    @FindBy(xpath = "//input[contains(@name, 'card_number')]")
    public WebElement PaymentCardNo;
    @FindBy(xpath = "//input[contains(@name, 'cvc')]")
    public WebElement PaymentCardCVV;
    @FindBy(xpath = "//input[contains(@name, 'expiry_month')]")
    public WebElement PaymentCardExpiryMonth;
    @FindBy(xpath = "//input[contains(@name, 'expiry_year')]")
    public WebElement PaymentCardExpiryYear;
    @FindBy(xpath = "//button[contains(text(),'Pay and Confirm Order')]")
    public WebElement ClickOnPayAndConfirmOrder;
    @FindBy(xpath = "//p[contains(text(),'Congratulations! Your order has been confirmed!')]")
    public WebElement VerifySuccessMessage;
    @FindBy(xpath = "//span[contains(text(),'Close')]")
    public WebElement ClickOnAdvertiseClose;
    @FindBy(xpath = "(//a[contains(@class, 'cart_quantity_delete')])[1]")
    public WebElement RemoveFirstProductFromCart;
    @FindBy(xpath = "(//a[contains(@class, 'cart_quantity_delete')])[2]")
    public WebElement RemoveSecondProductFromCart;
    @FindBy(xpath = "//b[contains(text(),'Cart is empty!')]")
    public WebElement VerifyCartEmpty;
    @FindBy(xpath = "//h2[contains(text(),'Category')]")
    public WebElement VerifyCategory;
    @FindBy(xpath = "(//i[contains(@class, 'plus')])[1]")
    public WebElement ClickOnWomensCategory;
    @FindBy(xpath = "(//a[contains(text(),'Dress')])[1]")
    public WebElement ClickOnWomensDress;
    @FindBy(xpath = "//h2[contains(text(),'Women - Dress Products')]")
    public WebElement VerifyWomenDressProductPage;
    @FindBy(xpath = "(//i[contains(@class, 'plus')])[2]")
    public WebElement ClickOnMensCategory;
    @FindBy(xpath = "//a[contains(text(),'Tshirts')]")
    public WebElement ClickOnMensTshirts;
    @FindBy(xpath = "//h2[contains(text(),'Men - Tshirts Products')]")
    public WebElement VerifyMenTShirtProductPage;
    @FindBy(xpath = "//h2[contains(text(),'Brands')]")
    public WebElement VerifyBrand;
    @FindBy(xpath = "(//a[contains(@href, 'brand_products')])[1]")
    public WebElement ClickOnFirstBrand;
    @FindBy(xpath = "//h2[contains(text(),'Brand -')]")
    public WebElement VerifyFirstBrandPage;
    @FindBy(xpath = "(//div[contains(@class, 'product-overlay')])[1]")
    public WebElement VerifyFirstBrandProduct;
    @FindBy(xpath = "(//a[contains(@href, 'brand_products')])[2]")
    public WebElement ClickOnSecondBrand;
    @FindBy(xpath = "//h2[contains(text(),'Brand -')]")
    public WebElement VerifySecondBrandPage;
    @FindBy(xpath = "//h2[contains(text(),'All Products')]")
    public WebElement VerifyAllProductPage;
    @FindBy(xpath = "//input[contains(@name, 'search')]")
    public WebElement SearchProduct;
    @FindBy(xpath = "//button[contains(@type, 'button')]")
    public WebElement ClickOnSearch;
    @FindBy(xpath = "//h2[contains(text(),'Searched Products')]")
    public WebElement VerifySearch;
    @FindBy(xpath = "(//div[contains(@class, 'product-overlay')])[1]")
    public WebElement VerifyFirstSearchProduct;
    @FindBy(xpath = "(//a[contains(@class, 'add-to-cart')])[1]")
    public WebElement AddFirstSearchProduct;
    @FindBy(xpath = "(//a[contains(@class, 'add-to-cart')])[3]")
    public WebElement AddSecondSearchProduct;
    @FindBy(xpath = "//a[contains(text(),'Write Your Review')]")
    public WebElement VerifyWriteReview;
    @FindBy(xpath = "//input[contains(@id, 'name')]")
    public WebElement ReviewName;
    @FindBy(xpath = "(//input[contains(@id, 'email')])[1]")
    public WebElement ReviewEmail;
    @FindBy(xpath = "//textarea[contains(@id, 'review')]")
    public WebElement AddReview;
    @FindBy(xpath = "//button[contains(text(),'Submit')]")
    public WebElement SubmitReview;
    @FindBy(xpath = "//span[contains(text(),'Thank you for your review.')]")
    public WebElement VerifySubmitReview;
    @FindBy(xpath = "//h2[contains(text(),'recommended items')]")
    public WebElement VerifyRecomendedProduct;
    @FindBy(xpath = "(//a[contains(text(),'Add to cart')])[69]")
    public WebElement AddRecomendedProduct1ToCart;
    @FindBy(xpath = "(//a[contains(text(),'Add to cart')])[72]")
    public WebElement AddRecomendedProduct2ToCart;
    @FindBy(xpath = "(//li[contains(@class, 'address_firstname')])[1]")
    public WebElement VerifyDeliveryAddressName;
    @FindBy(xpath = "(//li[contains(@class, 'address_firstname')])[2]")
    public WebElement VerifyBillingAddressName;
    @FindBy(xpath = "(//li[contains(@class, 'address_address1')])[1]")
    public WebElement VerifyDeliveryCompanyName;
    @FindBy(xpath = "(//li[contains(@class, 'address_firstname')])[4]")
    public WebElement VerifyBillingCompanyName;
    @FindBy(xpath = "(//li[contains(@class, 'address_city')])[1]")
    public WebElement VerifyDeliveryAddress;
    @FindBy(xpath = "(//li[contains(@class, 'address_city')])[2]")
    public WebElement VerifyBillingAddress;
    @FindBy(xpath = "//u[contains(text(),'Register / Login')]")
    public WebElement ClickOnRegLogIn;
    @FindBy(xpath = "//a[contains(text(),'Download Invoice')]")
    public WebElement ClickOnDownloadInvoice;
    @FindBy(xpath = "//i[contains(@class, 'fa fa-angle-up')]")
    public WebElement ClickOnRightSideUpArrow;
    @FindBy(xpath = "//h2[contains(text(),'Full-Fledged practice website for Automation Engineers')]")
    public WebElement VerifyFullFledgeText;
    @FindBy(xpath = "(//i[contains(@class, 'fa fa-angle-left')])[2]")
    public WebElement ClickOnLeftArrow;
    @FindBy(xpath = "//iframe[contains(@id, 'aswift_1')]")
    public WebElement SwitchToIframe1;
    @FindBy(xpath = "//iframe[contains(@id, 'ad_iframe')]")
    public WebElement SwitchToIframe2;

    @FindBy(xpath = "//div[contains(@class, 'close-button')]")
    public WebElement ClickOnCloseAdv;

//*[@id="ad_iframe"]


    public void LoginD() throws IOException {

        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(1);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String email = row.getCell(1).getStringCellValue();
            String password = row.getCell(2).getStringCellValue();
            String expectedHomepageTitle = row.getCell(3).getStringCellValue();
            String actualHomepageTitle = driver.getTitle();
            Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
            Signup.click();
            boolean displayed = VerifyLoginToYourAccount.isDisplayed();
            Assert.assertEquals(displayed, true);
            EnterEmailAddress.sendKeys(email);
            EnterLoginPassword.sendKeys(password);
            ClickOnLogin.click();
            boolean displayed1 = VerifyLoginAsUsername.isDisplayed();
            Assert.assertEquals(displayed1, true);
        }
    }
    public void Login() throws IOException {

        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(3);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String email = row.getCell(1).getStringCellValue();
            String password = row.getCell(2).getStringCellValue();
            String expectedHomepageTitle = row.getCell(3).getStringCellValue();
            String actualHomepageTitle = driver.getTitle();
            Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
            Signup.click();
            boolean displayed = VerifyLoginToYourAccount.isDisplayed();
            Assert.assertEquals(displayed, true);
            EnterEmailAddress.sendKeys(email);
            EnterLoginPassword.sendKeys(password);
            ClickOnLogin.click();
            boolean displayed1 = VerifyLoginAsUsername.isDisplayed();
            Assert.assertEquals(displayed1, true);
        }
    }
    public void LoginIncorrect() throws IOException {

        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(2);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String email = row.getCell(1).getStringCellValue();
            String password = row.getCell(2).getStringCellValue();
            String expectedHomepageTitle = row.getCell(3).getStringCellValue();
            String actualHomepageTitle = driver.getTitle();
            Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
            Signup.click();
            boolean displayed = VerifyLoginToYourAccount.isDisplayed();
            Assert.assertEquals(displayed, true);
            EnterEmailAddress.sendKeys(email);
            EnterLoginPassword.sendKeys(password);
            ClickOnLogin.click();
        }
    }
    public void NewUserSignup() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            String expectedSignupTitle = row.getCell(2).getStringCellValue();
            String Name = row.getCell(3).getStringCellValue();
            String Email = row.getCell(4).getStringCellValue();
            String ExpectedAccountInfoTitle = row.getCell(5).getStringCellValue();
            String password = row.getCell(6).getStringCellValue();
            double numericValue1 = row.getCell(7).getNumericCellValue();
            int intValue1 = (int) numericValue1;
            String date = String.valueOf(intValue1);
            String month = row.getCell(8).getStringCellValue();
            double numericValue2 = row.getCell(9).getNumericCellValue();
            int intValue2 = (int) numericValue2;
            String year = String.valueOf(intValue2);
            String firstName = row.getCell(10).getStringCellValue();
            String lastName = row.getCell(11).getStringCellValue();
            String CompanyName = row.getCell(12).getStringCellValue();
            String address1 = row.getCell(13).getStringCellValue();
            String address2 = row.getCell(14).getStringCellValue();
            String country = row.getCell(15).getStringCellValue();
            String state = row.getCell(16).getStringCellValue();
            String city = row.getCell(17).getStringCellValue();
            double numericValue3 = row.getCell(18).getNumericCellValue();
            int intValue3 = (int) numericValue3;
            String zipcode = String.valueOf(intValue3);
            double numericValue4 = row.getCell(19).getNumericCellValue();
            int intValue4 = (int) numericValue4;
            String mobileNo = String.valueOf(intValue4);
            String actualHomepageTitle = driver.getTitle();
            Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
            Signup.click();
            EnterName.sendKeys(Name);
            EnterEmail.sendKeys(Email);
            NewSignup.click();
            String ActualEnterAccountInfoTitle = driver.getTitle();
            Assert.assertEquals(ExpectedAccountInfoTitle, ActualEnterAccountInfoTitle);
            SelectTitle.click();
            EnterPassword.sendKeys(password);
            Select select1 = new Select(SelectDate);
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("window.scroll(300, 0);");
            select1.selectByVisibleText(date);
            Select select2 = new Select(SelectMonth);
            select2.selectByVisibleText(month);
            Select select3 = new Select(SelectYear);
            select3.selectByVisibleText(year);
            SelectNewsLetterCheckbox.click();
            SelectOfferPartnerCheckbox.click();
            EnterFirstName.sendKeys(firstName);
            js.executeScript("window.scroll(300, 0);");
            EnterLastName.sendKeys(lastName);
            EnterCompanyName.sendKeys(CompanyName);
            EnterAddress1.sendKeys(address1);
            EnterAddress2.sendKeys(address2);
            js.executeScript("window.scroll(300, 0);");
            Select select4 = new Select(SelectCountry);
            select4.selectByVisibleText(country);
            EnterState.sendKeys(state);
            EnterCity.sendKeys(city);
            js.executeScript("window.scroll(300, 0);");
            EnterZipcode.sendKeys(zipcode);
            EnterMobileNo.sendKeys(mobileNo);
            ClickOnCreateAccount.click();
            boolean displayed = VerifyCreateAccount.isDisplayed();
            Assert.assertEquals(displayed, true);
            ClickOnContinue.click();
            boolean displayed1 = VerifyLoginAsUsername.isDisplayed();
            Assert.assertEquals(displayed1, true);
        }
    }
    public void NewUserSignup1() throws IOException {
        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String Name = row.getCell(3).getStringCellValue();
            String Email = row.getCell(4).getStringCellValue();
            String ExpectedAccountInfoTitle = row.getCell(5).getStringCellValue();
            String password = row.getCell(6).getStringCellValue();
            double numericValue1 = row.getCell(7).getNumericCellValue();
            int intValue1 = (int) numericValue1;
            String date = String.valueOf(intValue1);
            String month = row.getCell(8).getStringCellValue();
            double numericValue2 = row.getCell(9).getNumericCellValue();
            int intValue2 = (int) numericValue2;
            String year = String.valueOf(intValue2);
            String firstName = row.getCell(10).getStringCellValue();
            String lastName = row.getCell(11).getStringCellValue();
            String CompanyName = row.getCell(12).getStringCellValue();
            String address1 = row.getCell(13).getStringCellValue();
            String address2 = row.getCell(14).getStringCellValue();
            String country = row.getCell(15).getStringCellValue();
            String state = row.getCell(16).getStringCellValue();
            String city = row.getCell(17).getStringCellValue();
            double numericValue3 = row.getCell(18).getNumericCellValue();
            int intValue3 = (int) numericValue3;
            String zipcode = String.valueOf(intValue3);
            double numericValue4 = row.getCell(19).getNumericCellValue();
            int intValue4 = (int) numericValue4;
            String mobileNo = String.valueOf(intValue4);
            EnterName.sendKeys(Name);
            EnterEmail.sendKeys(Email);
            NewSignup.click();
            SelectTitle.click();
            EnterPassword.sendKeys(password);
            Select select1 = new Select(SelectDate);
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("window.scroll(300, 0);");
            select1.selectByVisibleText(date);
            Select select2 = new Select(SelectMonth);
            select2.selectByVisibleText(month);
            Select select3 = new Select(SelectYear);
            select3.selectByVisibleText(year);
            SelectNewsLetterCheckbox.click();
            SelectOfferPartnerCheckbox.click();
            EnterFirstName.sendKeys(firstName);
            js.executeScript("window.scroll(300, 0);");
            EnterLastName.sendKeys(lastName);
            EnterCompanyName.sendKeys(CompanyName);
            EnterAddress1.sendKeys(address1);
            EnterAddress2.sendKeys(address2);
            js.executeScript("window.scroll(300, 0);");
            Select select4 = new Select(SelectCountry);
            select4.selectByVisibleText(country);
            EnterState.sendKeys(state);
            EnterCity.sendKeys(city);
            js.executeScript("window.scroll(300, 0);");
            EnterZipcode.sendKeys(zipcode);
            EnterMobileNo.sendKeys(mobileNo);
            ClickOnCreateAccount.click();
            boolean displayed = VerifyCreateAccount.isDisplayed();
            Assert.assertEquals(displayed, true);
            ClickOnContinue.click();
            boolean displayed1 = VerifyLoginAsUsername.isDisplayed();
            Assert.assertEquals(displayed1, true);
        }
    }
    public void ContactUs() throws IOException {

        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(5);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String expectedHomepageTitle = row.getCell(1).getStringCellValue();
            String Name = row.getCell(2).getStringCellValue();
            String Email = row.getCell(3).getStringCellValue();
            String Subject = row.getCell(4).getStringCellValue();
            String TextArea = row.getCell(5).getStringCellValue();
            String actualHomepageTitle = driver.getTitle();
            Assert.assertEquals(expectedHomepageTitle, actualHomepageTitle);
            ClickOnContactUs.click();
            boolean displayed = VerifyGetInTouch.isDisplayed();
            Assert.assertEquals(displayed, true);
            EnterNameHelp.sendKeys(Name);
            EnterEmailHelp.sendKeys(Email);
            EnterSubjectHelp.sendKeys(Subject);
            EnterTextAreaHelp.sendKeys(TextArea);
            ClickOnSubmitHelp.click();
            driver.switchTo().alert().accept();
            boolean displayed1 = VerifySuccessMsg.isDisplayed();
            Assert.assertEquals(displayed1, true);
        }
    }

    public void Payment() throws IOException {

        String excelFilePath = "C:\\Users\\Admin\\IdeaProjects\\AutomatioExerciseP2\\src\\test\\resources\\Project.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(7);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String comment = row.getCell(18).getStringCellValue();
            String cardName = row.getCell(19).getStringCellValue();
            double numericValue5 = row.getCell(20).getNumericCellValue();
            int intValue5 = (int) numericValue5;
            String cardNo = String.valueOf(intValue5);
            double numericValue6 = row.getCell(21).getNumericCellValue();
            int intValue6 = (int) numericValue6;
            String cvc = String.valueOf(intValue6);
            double numericValue7 = row.getCell(22).getNumericCellValue();
            int intValue7 = (int) numericValue7;
            String expiryMonth = String.valueOf(intValue7);
            double numericValue8 = row.getCell(23).getNumericCellValue();
            int intValue8 = (int) numericValue8;
            String expiryYear = String.valueOf(intValue8);
            ClickOnCart.click();
            ClickOnProceedToCheckout.click();
            boolean displayed2 = VerifyAddressDetails.isDisplayed();
            Assert.assertEquals(displayed2, true);
            boolean displayed3 = VerifyReviewOrder.isDisplayed();
            Assert.assertEquals(displayed3, true);
            CommentText.sendKeys(comment);
            ClickOnPlaceOrder.click();
            PaymentName.sendKeys(cardName);
            PaymentCardNo.sendKeys(cardNo);
            PaymentCardCVV.sendKeys(cvc);
            PaymentCardExpiryMonth.sendKeys(expiryMonth);
            PaymentCardExpiryYear.sendKeys(expiryYear);
            ClickOnPayAndConfirmOrder.click();
            boolean displayed4 = VerifySuccessMessage.isDisplayed();
            Assert.assertEquals(displayed4, true);
        }
    }

    public void DeleteAccount() throws IOException {

        ClickOnDeleteAccount.click();
        boolean displayed5 = VerifyDeleteAccount.isDisplayed();
        Assert.assertEquals(displayed5, true);
        ClickOnContinue.click();
    }
}