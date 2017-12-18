import {Injectable} from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import 'rxjs/add/operator/toPromise';
declare var $:any;

import {SharePointUserProfile} from './sharepoint-user-profile';

@Injectable()
export class SharePointUserProfileWebService{

	constructor(private http: HttpClient) { }
	private serviceUrl = '/_vti_bin/UserProfileService.asmx';//url can be absolute for the user profile service.
	private xmlPayloadWrapperStart = `<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>`;
	private xmlPayloadWrapperEnd = '</soap12:Body></soap12:Envelope>';
	private getUserProfileByNamePayload = `
    <GetUserProfileByName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">
      <AccountName>accountNameString</AccountName>
    </GetUserProfileByName>`;

	private headers = new HttpHeaders().set('Content-Type','application/soap+xml; charset=utf-8');
							 
	getUserProfileByName(accountName?:string):Promise<SharePointUserProfile>{
		let currentPayload = this.xmlPayloadWrapperStart+this.getUserProfileByNamePayload+this.xmlPayloadWrapperEnd;
		if(accountName){
			currentPayload = currentPayload.replace("accountNameString", accountName);
		}
		else{
			currentPayload = currentPayload.replace("accountNameString", "");
		}
		return this.http.post(this.serviceUrl, currentPayload,{headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res)
			{
				let filledResponse:SharePointUserProfile = new SharePointUserProfile();
				let jQueryValue:string = $(res).find("PropertyData:has(name:contains('WorkEmail'))").find('value').text();
				if(jQueryValue){
					filledResponse.workEmail = jQueryValue;
				}
				jQueryValue = $(res).find("PropertyData:has(name:contains('WorkPhone'))").find('value').text();
				if(jQueryValue){
					filledResponse.workPhone = jQueryValue;
				}
				
				$(res).find("z\\:row, row").each(function( index:any ) {
				let result:SharePointUserProfile = new SharePointUserProfile();
				//$(this)[0].attributes)
					//do some fancy jquery coding here.
				});
				return filledResponse;
		})
		.catch(this.handleError);
	}
	
	private handleError(error: any): Promise<any> {	
		console.error('An error occurred', error); // for demo purposes only
		return Promise.reject(error.message || error);
	}

		/* return this.http.post(dummyInstance.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
		.toPromise()
		.then(function(res)
			{
				let filledResponse:SharepointListItem[] = [];
				$(res.text()).find("z\\:row, row").each(function( index:any ) {
					filledResponse.push(new ctor($(this)[0].attributes)); 
				});
				return filledResponse;
		})
		.catch(this.handleError);
	}
	
	

		.then(function (res){
			return +$(res.text()).find("z\\:row:first, row:first").attr("ows_ID");
		})
	.catch(this.handleError);
	} */
	
	 
}