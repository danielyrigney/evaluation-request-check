const {By,Key,Builder, until} = require("selenium-webdriver");
require("chromedriver");
const reader = require('xlsx');


// code to read the xlsx data
const workbook = reader.readFile('./fiies.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
let data = reader.utils.sheet_to_json(worksheet);

// function to slow the program down to website speed
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function checkEvalConcerns(){
      // login and direct to website
      let driver = await new Builder().forBrowser("chrome").build();
      await driver.get("https://portal.austinisd.org");
      await sleep(500);

      await driver.wait(until.elementLocated(By.id("identification"))).then(el => el.sendKeys("XXXXXXXX",Key.RETURN));
      await driver.wait(until.elementLocated(By.id("ember533"))).then(el => el.sendKeys("XXXXXXXX",Key.RETURN));
      await sleep(500);
      await driver.get("https://austinisd.us001-rapididentity.com/idp/profile/SAML2/Unsolicited/SSO?providerId=https://Austin.acceliplan.com");
      await driver.get("https://austin.acceliplan.com/plan/Students/Landing");
      await sleep(500);

      for (let i=0; i < 1500; i++) {
        // loads the permnum to be checked
        var searchString = data[i].permnum;
        console.log("****************");
        console.log(data[i].name);
        console.log(data[i].permnum);
        // goes to the webpage to check
        await driver.wait(until.elementLocated(By.id("UniqueId"))).then(el => el.sendKeys(searchString,Key.RETURN));
        await sleep(5000);
        await driver.wait(until.elementLocated(By.linkText(data[i].name))).then(el => el.click());
        // await driver.wait(until.elementLocated(By.xpath("//a[@onclick='AjaxManager.showProgress()']"))).then(el => el.click());
        // await driver.findElement(By.xpath("//a[@onclick='AjaxManager.showProgress()']")).click();
        await driver.wait(until.elementLocated(By.linkText("Events"))).then(el => el.click());
        await driver.wait(until.elementLocated(By.linkText("IEP Referral Decision"))).then(el => el.click());
        await driver.wait(until.elementLocated(By.xpath("//a[contains(text(),'Notice of Referral Decision')]"))).then(el => el.click());

        //AU
        await driver.wait(until.elementLocated(By.xpath('//input[contains(@data-bind,"checked: EmotionalBehavioralDisorderSpecifiesCheckBox[1]")]')));

        let auIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: EmotionalBehavioralDisorderSpecifiesCheckBox[1]")]')).isSelected();
        if (auIsChecked) {
          data[i].au = "yes";
          data[i].si = "yes";
          data[i].ot = "yes";
          console.log("AU checked");
        };

        //SI
        let siIsChecked = await driver.wait(until.elementLocated(By.xpath('//input[contains(@data-bind,"checked: SpeechLanguageDisorderCheckbox")]'))).isSelected();
        if (siIsChecked) {
          data[i].si = "yes";
          console.log("SI has been checked");
        };

        // VI
        let viIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: OtherIDEAEligibleDisabilitySpecifiesCkeckBoxes[2]")]')).isSelected();
        if (viIsChecked) {
          data[i].vi = "yes";
          console.log("VI checked");
        };

        // AI
        let aiIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: OtherIDEAEligibleDisabilitySpecifiesCkeckBoxes[3]")]')).isSelected();
        if (aiIsChecked) {
          data[i].ai = "yes";
          console.log("AI checked");
        };

        // ID
        let idIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: IntellectualDisabilityCheckbox")]')).isSelected();
        if (idIsChecked) {
          data[i].id = "yes";
          console.log("ID checked");
        };

        // SLD Reading
        let sldrIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: SpecificLearningDisability[0]")]')).isSelected();
        if (sldrIsChecked) {
          data[i].sldr = "yes";
          console.log("Reading/Dyslexia checked");
        };

        // SLD Math
        let sldmIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: SpecificLearningDisability[2]")]')).isSelected();
        if (sldmIsChecked) {
          data[i].sldm = "yes";
          console.log("Math checked");
        };

        // SLD Writing
        let sldwIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: SpecificLearningDisability[1]")]')).isSelected();
        if (sldwIsChecked) {
          data[i].sldw = "yes";
          console.log("Wrting checked");
        };

        // ED
        let edIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: EmotionalBehavioralDisorderSpecifiesCheckBox[0]")]')).isSelected();
        if (edIsChecked) {
          data[i].ed = "yes";
          console.log("ED checked");
        };

        // OHI
        let ohiIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: OtherIDEAEligibleDisabilitySpecifiesCkeckBoxes[0]")]')).isSelected();
        if (ohiIsChecked) {
          data[i].ohi = "yes";
          console.log("OHI checked");
        };

        // TBI
        let tbiIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: OtherIDEAEligibleDisabilitySpecifiesCkeckBoxes[1]")]')).isSelected();
        if (tbiIsChecked) {
          data[i].tbi = "yes";
          console.log("TBI checked");
        };

        // OI
        let oiIsChecked = await driver.findElement(By.xpath('//input[contains(@data-bind,"checked: OtherIDEAEligibleDisabilitySpecifiesCkeckBoxes[4]")]')).isSelected();
        if (oiIsChecked) {
          data[i].oi = "yes";
          console.log("OI checked");
        };

        await driver.wait(until.elementLocated(By.xpath("//span[contains(text(),'Notice of Proposal to Evaluate')]"))).then(el => el.click());
        //await sleep(7000);

        await driver.wait(until.elementLocated(By.xpath('//div[contains(@data-bind,"value: MotorAbilitiesTextbox")]')));

        let motorAbilitiesTextbox = await driver.wait(until.elementLocated(By.xpath('//div[contains(@data-bind,"value: MotorAbilitiesTextbox")]')));
        await motorAbilitiesTextbox.getText().then(function (text) {
          text = text.toLowerCase();
          if (text.search("occupational") >= 0 || text.search(" ot ") >= 0){
            data[i].ot = "yes";
            console.log("OT has been checked");
          };

          if (text.search("physical therapy") >= 0 || text.search(" pt ") >= 0 || text.search("physical and occupational") >= 0 ){
            data[i].pt = "yes";
            console.log("PT has been checked");
          };

          if (text.search("adaptive") >= 0 || text.search(" ape ") >= 0) {
            data[i].ape = "yes";
            console.log("APE has been checked");
          };

        });

        let assistiveTechnologyTextbox = await driver.findElement(By.xpath('//div[contains(@data-bind,"value: AssistiveTechnologyText")]'))
        await assistiveTechnologyTextbox.getText().then(function (text) {
          text = text.toLowerCase();
          if (text.search("formal assistive technology") >= 0 || text.search("at evaluation") >= 0) {
            data[i].at = "yes";
            console.log("AT has been checked");
          };
        });
        await sleep(1000);
        await driver.get("https://austin.acceliplan.com/plan/Students/Landing");
        await sleep(4000);
      };

      let writeToWorkBook = reader.utils.book_new();
      const writeToWorkSheet = reader.utils.json_to_sheet(data);
      let exportFileName = `results.xlsx`;
      reader.utils.book_append_sheet(writeToWorkBook, writeToWorkSheet, `response`);
      reader.writeFile(writeToWorkBook, exportFileName);

      await driver.quit();
}

checkEvalConcerns();
