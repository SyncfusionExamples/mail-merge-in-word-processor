## Steps to run angular app and web services 

To work on this demo, run the web services and angular app simultaneously.

### Run the web services

To Run the webservices in "Web Services/ASP.NET Core" folder, run the following commands
* **dotnet restore** to restore the nuget packages.
* **dotnet run** to run the web services.

### Run the angular app

To run the angular app,
* Install the package using **npm install** command.
* Then, set the **serviceURL** in Document editor by running web services URL.

 For example, if your running URL is "http://localhost:62869/", change the service URL in app.component.ts file(line number: 140).
* Run the app, using **npm start** command. Now, the document editor will launch in browser.