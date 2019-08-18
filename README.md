# automailer
A simple Node.js application that reads students from a list of excel sheet (.xlxs) and automatically sends them email regarding their status.

# setting up application

Clone the application and then run npm install to install all the required node modules. After that, create a config.json file which will contain: email, clientId, clientSecret, refreshToken and accessToken.

Use [Google App Console](https://support.google.com/googleapi/answer/6158862?hl=en) to set up a Google project. Generate clientId and clientSecret. Ensure you enable GMail APIs in the project.

Use [Google Outh2 Playground](https://developers.google.com/oauthplayground/) to get accessToken and refreshToken.

# running the application

In order to run the application, type the command 'node index'.
