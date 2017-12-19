# NgTaxSharePointWebServicesModule

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 1.4.2.

[NPM repository](https://www.npmjs.com/package/ng-tax-share-point-web-services-module)

Usage ( IN your app.module.ts:) 
import {NgTaxServices} from 'ng-tax-share-point-web-services-module';
import {HttpClientModule} from '@angular/common/http'; //Utilizes the new HTTP Client


then in your imports array of your main moduel: NgTaxServices.forRoot()

It encapsulates some SharePoint ASMX web services.
Tested against a SharePoint 2010, but in theory should work at least on 2013, and hopefully on 2016 and online but have no idea.


Dependencies: Please add JQuery to the HTML of your project, i tested it as this:
<script src="/SiteAssets/jquery-3.1.1.min.js"></script>

JQuery is used solely for manipulating the XML Results.

If you consume lists.asmx web service (SharepointListsWebService), you will likely need to extend the basic SharepointListItem object:

[This](https://github.com/isaacchacon/NgVehicleRegistrationForm/blob/master/src/VehicleRegistrationReactive/vehicle-registration-list-entry.ts) is one example of a child class that properly consumes the service.

Put the SharePoint column names into the getItemProperties.



## Development server

Run `ng serve` for a dev server. Navigate to `http://localhost:4200/`. The app will automatically reload if you change any of the source files.

## Code scaffolding

Run `ng generate component component-name` to generate a new component. You can also use `ng generate directive|pipe|service|class|guard|interface|enum|module`.

## Build

Run `ng build` to build the project. The build artifacts will be stored in the `dist/` directory. Use the `-prod` flag for a production build.

## Running unit tests

Run `ng test` to execute the unit tests via [Karma](https://karma-runner.github.io).

## Running end-to-end tests

Run `ng e2e` to execute the end-to-end tests via [Protractor](http://www.protractortest.org/).
Before running the tests make sure you are serving the app via `ng serve`.

## Further help

To get more help on the Angular CLI use `ng help` or go check out the [Angular CLI README](https://github.com/angular/angular-cli/blob/master/README.md).
