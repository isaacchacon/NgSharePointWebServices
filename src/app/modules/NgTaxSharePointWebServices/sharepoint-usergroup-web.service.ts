import {Injectable} from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import 'rxjs/add/operator/toPromise';

declare var $:any;

import {TaxSpUser} from './tax-sp-user';


@Injectable()
export class SharepointUserGroupWebService{
	constructor(private http: HttpClient) { }
	private serviceUrl = '_vti_bin/UserGroup.asmx';
	private xmlPayloadWrapperStart = `<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>`;
	private xmlPayloadWrapperEnd = '</soap12:Body></soap12:Envelope>';
	private headers = new HttpHeaders().set('Content-Type','application/soap+xml; charset=utf-8');
	private getUserLoginFromEmailPayload = `<GetUserLoginFromEmail xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/"><emailXml><Users>payload</Users></emailXml></GetUserLoginFromEmail>`;
	private addUserCollectionToGroupPayload=`<AddUserCollectionToGroup xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">
      <groupName>groupNamePayload</groupName>
      <usersInfoXml><Users>usersPayload</Users></usersInfoXml>
    </AddUserCollectionToGroup>`;
	private removeUserCollectionFromGroupPayload=`<RemoveUserCollectionFromGroup xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">
      <groupName>groupNamePayload</groupName>
      <userLoginNamesXml><Users>usersPayload</Users></userLoginNamesXml>
    </RemoveUserCollectionFromGroup>`;
	private getUserCollectionFromGroupPayload=`<GetUserCollectionFromGroup xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">
      <groupName>groupNamePayload</groupName></GetUserCollectionFromGroup>`;
	
	
	getUserLoginFromEmail(emailStrings:string[], siteurl:string):Promise<TaxSpUser[]>{
		let requestBody=this.xmlPayloadWrapperStart+this.getUserLoginFromEmailPayload+this.xmlPayloadWrapperEnd;
		let userString = "<User Email=\"emailString\"/>";
		let finalPayload = "";
		for(let x:number = 0;x<emailStrings.length;x++){
			finalPayload+=userString.replace("emailString", emailStrings[x]);
		}
		requestBody = requestBody.replace("payload", finalPayload);
		return this.http.post(siteurl+this.serviceUrl, requestBody, {headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res){
			let result:TaxSpUser[]=[];
			$(res).find("User").each(function(index:any){				
				result.push( {id:$(this).attr('SiteUser'),displayName:$(this).attr('DisplayName'), login:$(this).attr('Login'), email:$(this).attr('Email')});
			});
			return result;
		})
		.catch(this.handleError);
	}
	
	addUserCollectionToGroup(users:TaxSpUser[],groupName:string, siteurl:string):Promise<number>{
		let requestBody=this.xmlPayloadWrapperStart+this.addUserCollectionToGroupPayload+this.xmlPayloadWrapperEnd;
		requestBody = requestBody.replace("groupNamePayload", groupName);
		let internalpayload:string="";
		for(let x = 0 ;x<users.length;x++){
			/*internalpayload+="<User LoginName=\""+users[x].login+"\" Email=\""
			+users[x].email+"\" Name=\""+ users[x].displayName +"\"/>";*/
			internalpayload+="<User LoginName=\""+users[x].login+"\"/>";
		}
		requestBody = requestBody.replace("usersPayload", internalpayload);
		return this.http.post(siteurl+this.serviceUrl, requestBody, {headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res){
			return 0;
		})
		.catch(this.handleError);
		
	}
	
	
	removeUserCollectionFromGroup(users:TaxSpUser[],groupName:string, siteurl:string):Promise<number>{
		let requestBody=this.xmlPayloadWrapperStart+this.removeUserCollectionFromGroupPayload+this.xmlPayloadWrapperEnd;
		requestBody = requestBody.replace("groupNamePayload", groupName);
		let internalpayload:string="";
		for(let x = 0 ;x<users.length;x++){
			internalpayload+="<User LoginName=\""+users[x].login+"\"/>";
		}
		requestBody = requestBody.replace("usersPayload", internalpayload);
		return this.http.post(siteurl+this.serviceUrl, requestBody, {headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res){
			return 0;
		})
		.catch(this.handleError);
		
	}
	
	
	getUserCollectionFromGroup(groupName:string, siteUrl:string):Promise<TaxSpUser[]>{
		let requestBody=this.xmlPayloadWrapperStart+this.getUserCollectionFromGroupPayload+this.xmlPayloadWrapperEnd;
		requestBody = requestBody.replace("groupNamePayload" ,groupName);
		return this.http.post(siteUrl+this.serviceUrl, requestBody, {headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res){
			let result:TaxSpUser[]=[];
			$(res).find("User").each(function(index:any){				
				result.push( {id:$(this).attr('ID'),displayName:$(this).attr('Name'), login:$(this).attr('LoginName'), email:$(this).attr('Email')});
			});
			return result;
		})
		.catch(this.handleError);
	}
	
	private handleError(error: any): Promise<any> {
	
    console.error('An error occurred', error); // for demo purposes only
    return Promise.reject(error.message || error);
  }

}