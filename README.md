# Update Library Document Set Welcome Page View using Playwright and Azure Function

1. This is using [Playright](https://playwright.dev) which is pretty new and maintained by Microsoft.
2. This uses a Linux NodeJs Azure Function so you will have to use Visual Studio or Visual Studio code to deploy.
3. This was made to faciliate https://github.com/pnp/pnpframework/issues/105 till there is a proper solution from the PnP Framework.

## Things to note
1. If you are using VS Code ensure your ```.vscode/settings.json``` has ```"azureFunctions.scmDoBuildDuringDeployment": true```
2. Ensure Azure Function has ```"PLAYWRIGHT_BROWSERS_PATH" :0``` in its function app settings. Remove this if you are testing locally.
3. Navigate to /home/site/wwwroot and run npm install and npm install playwright-chromium. I had it failed on me when playwright-chromium didn't get installed as part of the package.json.
4. This require a SharePoint Admin account providing username and password to Playwright to get access to the required page as SharePoint isn't an anonymous access system. So it is advisable to create a separate account or update to use Azure KeyVault

## References
1. https://anthonychu.ca/post/azure-functions-headless-chromium-puppeteer-playwright/
2. https://dotnetthoughts.net/running-playwright-on-azure-functions/

## TODO
1. Link with KeyVault

## Deployment
1. Create Linux NodeJs Azure Function App that is a running on an app service plan.
2. func azure functionapp publish <Function App Name> --build remote
3. Using kudu run npm install and npm install playwright-chromium