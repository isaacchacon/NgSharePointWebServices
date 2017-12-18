import { CommonModule } from '@angular/common';
import {  ModuleWithProviders, NgModule, Optional, SkipSelf } from  '@angular/core';

import {SharepointListsWebService} from './sharepoint-lists-web.service';
import {UrlService} from './url-service';
import {SharePointUserProfileWebService} from './sharepoint-user-profile-web.service';
import {SharepointUserGroupWebService} from './sharepoint-usergroup-web.service';


@NgModule({
  imports:      [ CommonModule ],
  declarations: [],
  exports: []
})
export class NgTaxServices { 

	constructor (@Optional() @SkipSelf() parentModule: NgTaxServices) {
		  if (parentModule) {
			throw new Error(
			  'NgTaxServices / NgTaxSharePointWebServicesModule is already loaded. Import it in the AppModule only');
		  }
	}

  public static forRoot(): ModuleWithProviders {
		return {
		  ngModule: NgTaxServices,
		  providers: [SharepointListsWebService,UrlService,SharePointUserProfileWebService,SharepointUserGroupWebService]
		};
	}
}
