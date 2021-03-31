// Global
const { chromium } = require("playwright-chromium");
const spauth = require("node-sp-auth");
const bootstrap = require("pnp-auth").bootstrap;
const sp = require("@pnp/sp-commonjs").sp;

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger beginning Update-DocSetWelcomePageView request.');
    // Note that this is for Azure Function so there is a request body to be sent along.
    const groupAlias = req.body.GroupAlias;
    const listName = req.body.ListName;
    const contentTypeName = req.body.ContentTypeName;
    const welcomePageViewName = req.body.WelcomePageViewName;

    const siteCollectionUrl = `https://${process.env.TenantName}.sharepoint.com/sites/${groupAlias}`;

    bootstrap(sp, { username: process.env.SPAdminUsername, password: process.env.SPAdminPassword }, siteCollectionUrl);

    const list = await sp.web.lists.getByTitle(listName).get();
    let listContentTypeId = "";
    const listContentTypes = await sp.web.lists.getByTitle(listName).contentTypes.get().then(function (result) {
        result.forEach(element => {
            if (element.Name = contentTypeName) {
                listContentTypeId = element.Id.StringValue;
            }
        });
    });

    // Doc Set Setting page to go to     
    const docSetSettingsPageUrl = `${siteCollectionUrl}/_layouts/15/docsetsettings.aspx?ctype=${listContentTypeId}&List=${list.Id}`;

    // Playwright Automation
    let responseBody = null;
    let statuscode = 500;
    const browser = await chromium.launch();
    const page = await browser.newPage();
    // Authentication setup
    const authObject = await spauth.getAuth(docSetSettingsPageUrl, {
        username: process.env.SPAdminUsername,
        password: process.env.SPAdminPassword,
        online: true
    });

    // Add the authentication headers
    await page.setExtraHTTPHeaders(authObject.headers);

    // Do the magic
    await setDocSetWelcomePageView(page, docSetSettingsPageUrl, welcomePageViewName, responseBody, statuscode);

    // Cleanup and close browser
    await browser.close();

    // Return response
    context.res = {
        status: statuscode,
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

async function setDocSetWelcomePageView(page, pageUrl, welcomePageViewName, responseBody, statuscode) {

    // Go to the library settings page
    await page.goto(pageUrl);
    const listContentTypeTitle = await page.goto(pageUrl);
    page.waitForSelector("#ctl00_PlaceHolderPageTitleInTitleArea_lblBreadcrumb");

    if (listContentTypeTitle) {
        // Scroll to where the Welcome Page View Dropdown is so that the automation tool can find it
        await page.waitForSelector("#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");
        await scrollOnElement(page, "#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");

        // Update the Welcome Page View
        await page.click("#ctl00_PlaceHolderMain_idWelcomePageView_ctl01_DropDownListViews");
        await page.selectOption("[name='ctl00$PlaceHolderMain$idWelcomePageView$ctl01$DropDownListViews']", [{ 'label': welcomePageViewName }]);

        // Submit and leave
        await page.click("[type='submit']");

        responseBody = `Document library is now set. Visit <a href='${pageUrl}'>Document Set Settings Page</a> to verify`;
        statuscode = 200;
    }
    else {
        responseBody = `Something is wrong with the page at ${pageUrl}`;
        statuscode = 404;
    }
}