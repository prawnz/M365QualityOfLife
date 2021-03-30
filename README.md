# Update Library Document Set Welcome Page View using Playwright and Azure Function

1. This is using [Playright](https://playwright.dev) which is pretty new and maintained by Microsoft.
2. This uses a Linux NodeJs Azure Function so you will have to use Visual Studio or Visual Studio code to deploy.
3. This was made to faciliate https://github.com/pnp/pnpframework/issues/105 till there is a proper solution from the PnP Framework.

## Things to note
1. If you are using VS Code ensure your ```.vscode/settings.json``` has ```"azureFunctions.scmDoBuildDuringDeployment": true```
2. Ensure Azure Function has ```"PLAYWRIGHT_BROWSERS_PATH" :0``` in its function app settings. Remove this if you are testing locally.
3. Navigate to /home/site/wwwroot and run npm install and npm install playwright-chromium. I had it failed on me when playwright-chromium didn't get installed as part of the package.json.
4. This require a SharePoint Admin account providing username and password to Playwright to get access to the required page as SharePoint isn't an anonymous access system. I've followed Elio's method for the username and password. File is stored in the Auth folder
<pre>To make it a bit “safer” I added a dependency called cpass in order to encode and decode the passwords. In the project, you find an encoder.js file which you can use to encode your username and password. Add your username and password in it, and run node encoder.js, after this you can remove the file.Username and Password is encoded using CPASS.</pre>

## References
1. https://anthonychu.ca/post/azure-functions-headless-chromium-puppeteer-playwright/
2. https://dotnetthoughts.net/running-playwright-on-azure-functions/
3. https://www.eliostruyf.com/testing-the-ui-of-your-spfx-solution-with-puppeteer-and-jest/

## TODO
1. Use PNPJS to get the Content Type and List ID based on the title.    


## Deployment
1. Create Linux Consumption Function App that is a 
2. func azure functionapp publish <Function App Name> --build remote
3. Using kudu run npm install
