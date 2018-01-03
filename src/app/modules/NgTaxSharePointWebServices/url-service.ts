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
  
  getCurrentSiteUrl():string{
	  if(typeof _spPageContextInfo  !=='undefined'){
		  //preference for the SharePoint site, if hosted inside a webpart
			return _spPageContextInfo.webServerRelativeUrl+"/";  
	  }else{
			return this.InternalGetUrlKeyValue("JUrl", null);
	  }
  }
  
  InternalGetUrlKeyValue(d:string,a:string):string{
	  let c="";
	  if(a==null){
		  a=window.location.href+""
	  }
	  let b=a.indexOf("#");
	  if(b>=0){
		  a=a.substr(0,b);
	  }
	  b=a.indexOf("&"+d+"=");
	  if(b==-1){
		  b=a.indexOf("?"+d+"=");
	  }
	  if(b!=-1){
		  let ndx2=a.indexOf("&",b+1);
		  if(ndx2==-1){
			  ndx2=a.length;
		  }
		  c=a.substring(b+d.length+2,ndx2)
		}
		return c;
	}
}