import { Injectable } from '@angular/core';
declare var GetUrlKeyValue: any;
declare var window:any;
declare var _spPageContextInfo:any;

@Injectable()
/** Dummy version of an authenticated user service */
export class UrlService {

/*Takes advantage of pre-existing method in SharePoint's init.js*/
  navigateToDefaultSource(): void{
	if(GetUrlKeyValue('Source', false, window.location.href)){
		window.location.href= GetUrlKeyValue('Source', false, window.location.href);
	}
	else{
		window.location.href = _spPageContextInfo.webServerRelativeUrl;
	}
  }
  
  getItemId():string{
  try{
	return GetUrlKeyValue('ID', false, window.location.href);
	}catch(x){
		return '';
	}
	
	
  }
}