const { Builder, By, Key, until } = require('selenium-webdriver');
const assert = require('assert');
const { StaleElementReferenceError } = require('selenium-webdriver/lib/error');
const ExcelJS = require('exceljs');
const edge = require('selenium-webdriver/edge');

const edgeDriverPath = 'C:/dedge/msedgedriver.exe';
const edgeOptions = new edge.Options();
let driver;
let vars = {};


async function clickElement(driver, locator) {
    const element = await driver.wait(until.elementLocated(locator), 10000);
    await driver.wait(until.elementIsEnabled(element), 10000);  // Modificación aquí
    await element.click();
  }

async function sendKeysToElement(driver, locator, keysToSend) {
  const element = await driver.wait(until.elementLocated(locator), 10000);
  await element.sendKeys(keysToSend);
}
async function clickElementByXpath(driver, xpath) {
    let elementLocator = By.xpath(xpath);
    await driver.wait(until.elementLocated(elementLocator), 10000);
    let element = await driver.findElement(elementLocator);
    await element.click();
  }
  async function leerDatosDesdeExcel() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/Andre/Downloads/RegistroDeDesconexiones_Completo.xlsx');
    const worksheet = workbook.getWorksheet('DatosDePrueba');

    const datos = [];


    for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const dato = {
            columna1: row.getCell(1).value,
            columna2: row.getCell(2).value,
            columna3: row.getCell(3).value,
            columna4: row.getCell(4).value,
            columna5: row.getCell(5).value,
            columna6: row.getCell(6).value,
            columna7: row.getCell(7).value,
            columna8: row.getCell(8).value,
            columna9: row.getCell(9).value,
            columna10: row.getCell(10).value,
            columna12: row.getCell(12).value,
            columna13: row.getCell(13).value,
            columna14: row.getCell(14).value,
            columna15: row.getCell(15).value,
            columna16: row.getCell(16).value,
            // Agrega más columnas según sea necesario
          };
          datos.push(dato);
      }
  
      return datos;
}

describe('Prueba_1', function() {
  this.timeout(30000);

  beforeEach(async function() {
    driver = await new Builder()
      .forBrowser('MicrosoftEdge')
      .setEdgeOptions(edgeOptions)
      .setEdgeService(new edge.ServiceBuilder(edgeDriverPath))
      .build();
    vars = {};
  });

  afterEach(async function() {
    if (driver) {
      await driver.quit();
    }
  });
  

  it('Prueba_1 con datos desde Excel', async function() { 
    try{
    await driver.get("https://appsdev.itaipu.gov.br/SAM");

    await sendKeysToElement(driver, By.id("OutSystemsUIWeb_wt15_block_wtLogin_wt18_wtUsername_wtUserNameInput"), "ocampos");
    await sendKeysToElement(driver, By.id("OutSystemsUIWeb_wt15_block_wtLogin_wt18_wtPassword_wtPasswordInput"), "mJVuRbE6");
    await clickElement(driver, By.id("OutSystemsUIWeb_wt15_block_wtLogin_wt18_wtAction_wt12"));
    await clickElementByXpath(driver, "//div[@id='SAMTemplateAssets_wt9_block_IguazuTheme_wt28_block_wtMainContent_wtMainContent_SAMTemplateAssets_wt6_block_wt4_wtListRecordsLaunchpadApps_ctl18_IguazuTheme_wt7_block_wtAppImage']/img");
    await clickElement(driver, By.id("b6-b7-Menu"));
    await clickElementByXpath(driver,"//div[@id=\'b6-b7-Items\']/a[1]");
    await driver.get("https://appsdev.itaipu.gov.br/SAMPTD/DisconnectionRegistrationList");
    await driver.wait(until.elementLocated(By.css(".fa-eraser")), 10000).click();
    await driver.findElement(By.id("Dropdown_SystemType")).click();
    
    const datos = await leerDatosDesdeExcel();

    

    for (const dato of datos) {
    
      // Continuar llenando los campos específicos de la prueba
      await driver.findElement(By.id("Input_DisconnectionNumber")).click();
      await driver.findElement(By.id("Input_DisconnectionNumber")).sendKeys(dato.columna1); 
      await driver.findElement(By.id("Dropdown_SystemType")).click();
      await driver.findElement(By.id("Dropdown_SystemType")).sendKeys(dato.columna2);
      await driver.findElement(By.id("Input_InitialLocalization")).sendKeys(dato.columna3); 
      await driver.findElement(By.id("Input_FinalLocalization")).sendKeys(dato.columna4);
      await driver.findElement(By.id("Dropdown_LocalizationType")).click();
      await driver.findElement(By.id("Dropdown_LocalizationType")).sendKeys(dato.columna5);
      await driver.findElement(By.id("Dropdown_ImportanceType")).click();
      await driver.findElement(By.id("Dropdown_ImportanceType")).sendKeys(dato.columna6);
      await driver.findElement(By.id("Input_InitialDate")).click(); 
      await driver.findElement(By.id("Input_InitialDate")).clear();
      await driver.findElement(By.id("Input_InitialDate")).sendKeys(dato.columna7.toString());
      await driver.findElement(By.id("Input_FinalDate")).click(); 
      await driver.findElement(By.id("Input_FinalDate")).sendKeys(dato.columna8.toString());
      await driver.findElement(By.id("Input_InitialYearWeek")).click();  
      await driver.findElement(By.id("Input_InitialYearWeek")).sendKeys(dato.columna9); 
      await driver.findElement(By.id("Input_FinalYearWeek")).sendKeys(dato.columna10);
      await driver.findElement(By.id("Dropdown_Division")).click();
      await driver.findElement(By.id("Dropdown_Division")).sendKeys(dato.columna12);
      await driver.findElement(By.id("Dropdown_Sector")).click();
      await driver.findElement(By.id("Dropdown_Sector")).sendKeys(dato.columna13);
      await driver.findElement(By.id("Dropdown_DisconnectionType")).click();
      await driver.findElement(By.id("Dropdown_DisconnectionType")).sendKeys(dato.columna14);
      await driver.findElement(By.id("Dropdown_Periodicity")).click();
      await driver.findElement(By.id("Dropdown_Periodicity")).sendKeys(dato.columna15);
      await driver.findElement(By.id("Input_WorkDescription")).sendKeys(dato.columna16);
      await driver.sleep(1000);
    
    await driver.findElement(By.css(".fa-search")).click();
    {
      const element = await driver.findElement(By.css(".fa-search"));
      await driver.actions({ bridge: true }).move(element).perform();
    }

    const element = await driver.findElement(By.css("body"));
    await driver.actions({ bridge: true }).move(element, { x: 0, y: 0 }).perform();
    await driver.findElement(By.css(".fa-eraser")).click();
    }
  } catch (error) {
    console.error('Error durante la ejecución:', error);
}
  })
})
