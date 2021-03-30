// Global
const { chromium } = require("playwright-chromium");
const spauth = require("node-sp-auth");
const Cpass = require('cpass').Cpass;

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger beginning Update-DocSetWelcomePageView request.');

    // Note that this is for Azure Function so there is a request body to be sent along.
    const groupAlias = req.body.GroupAlias;
    const contentTypeID = req.body.ContentTypeID;
    const listID = req.body.ListID;
    const welcomePageViewName = req.body.WelcomePageViewName;

    // Doc Set Setting page to go to     
    const pageUrl = `https://${process.env.TenantName}.sharepoint.com/sites/${groupAlias}/_layouts/15/docsetsettings.aspx?ctype=${contentTypeID}&List=${listID}`;
    

    // Authentication setup
    const cpass = new Cpass();
    const authObject = await spauth.getAuth(pageUrl, {
        username: cpass.decode(process.env.SPAdminUsername),
        password: cpass.decode(process.env.SPAdminPassword)
    });

    // Do the magic
    // Playwright Automation
    const browser = await chromium.launch();
    const page = await browser.newPage();
    await setDocSetWelcomePageView(authObject, page, pageUrl, welcomePageViewName);
    const responseBody = `Document library is now set. Visit <a href='${pageUrl}'>Document Set Settings Page</a> to verify`
    await browser.close();

    // Return response
    context.res = {
        status: 200,
        body: responseBody,
        headers: {
            "content-type": "text/html"
        }
    };
}

// Custom Helper Functions
async function scrollOnElement(page, selector) {
    await page.$eval(selector, (element) => {
        element.scrollIntoView();
    });
}

async function setDocSetWelcomePageView(authObject, page, pageUrl, welcomePageViewName) {

    // Add the authentication headers
    await page.setExtraHTTPHeaders(authObject.headers);
    await page.goto(pageUrl);

    // Scroll to where the Welcome Page View Dropdown is so that the automation tool can find it
    await page.waitForSelector("#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");
    await scrollOnElement(page, "#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");

    // Update the Welcome Page View
    await page.click("#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");
    await page.selectOption("[name='ctl00$PlaceHolderMain$idWelcomePageView$ctl01$DropDownListViews']", [{ 'label': welcomePageViewName }]);

    // Submit and leave
    await page.click("[type='submit']");
}