import {Injectable} from '@angular/core';
import {Headers, Http} from '@angular/http';
import 'rxjs/add/operator/toPromise';
declare var $:any;

import {SharepointListItem} from './sharepoint-list-item';
import {SharepointListItemConstructor} from './sharepoint-list-item-constructor';

@Injectable()
export class SharepointListsWebService{

	constructor(private http: Http) { }
	private serviceUrl = '_vti_bin/Lists.asmx';
	private xmlPayloadWrapperStart = `<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>`;
	private xmlPayloadWrapperEnd = '</soap12:Body></soap12:Envelope>';
	private getListItemsPayload = `
    <GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <listName>listNamePayLoad</listName>
      <viewName>viewNamePayLoad</viewName>
      <query>queryPayload</query>
      <viewFields>viewFieldsPayload</viewFields>
      <rowLimit>rowLimitPayload</rowLimit>
      <QueryOptions>queryOptionsPayload</QueryOptions>
    </GetListItems>`;
	
	private updateListItemsPayload = `<UpdateListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>listNamePayLoad</listName><updates>updatesPayLoad</updates></UpdateListItems>`;

	private headers = new Headers({
								 'Content-Type': 'application/soap+xml; charset=utf-8',
							 });

	///If Filter query is not empty, we will use it to filter the CAML QUery.						 
	//Filter QUery : ifr provided, need to be a field Equals to a Value.
	//camlQuery: If provided, the caml query is applied.
	// If no filter nor caml query, an empty query is passed (likely to get all the items).
    getListItems(ctor:SharepointListItemConstructor, filterQuery:[string,string], camlQuery:string, orderByField:string):Promise<SharepointListItem[]>{
		let dummyInstance = new ctor();
		let currentPayload = this.xmlPayloadWrapperStart+this.getListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad",dummyInstance.getListName());
		currentPayload = currentPayload.replace("viewNamePayLoad", "");
		if(filterQuery||camlQuery){
		
			if(filterQuery){
				let internalCamlQuery = "<Query><Where><Eq><FieldRef Name='FieldName' /><Value Type='Text'>ValueName</Value></Eq></Where></Query>"
				internalCamlQuery = internalCamlQuery.replace("FieldName", filterQuery[0]);
				internalCamlQuery = internalCamlQuery.replace("ValueName", filterQuery[1]);
				currentPayload = currentPayload.replace("queryPayload", internalCamlQuery );
			}
			else{
				currentPayload = currentPayload.replace("queryPayload", camlQuery );
			}
			
		}else{
			currentPayload = currentPayload.replace("queryPayload", "");
		}
		currentPayload = currentPayload.replace("viewFieldsPayload", "");
		currentPayload = currentPayload.replace("rowLimitPayload", "");
		currentPayload = currentPayload.replace("queryOptionsPayload", "");		
		return this.http.post(dummyInstance.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
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
	
	
	/*Use this method for updating only one column in the list.*/
	updateListItem(itemToUpdate:SharepointListItem,newValue:string):Promise<void>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad", itemToUpdate.getListName());
		currentPayload = currentPayload.replace("updatesPayLoad", `<Batch OnError="Continue"><Method ID="1" Cmd="Update"><Field Name="ID">`+itemToUpdate.ID+`</Field><Field Name="`+itemToUpdate.getFieldToUpdate()+`">`+newValue+`</Field></Method></Batch>`);
	return this.http.post(itemToUpdate.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
		.toPromise()
		.then(()=> null)
	.catch(this.handleError);
	}
	
	 private handleError(error: any): Promise<any> {
	
    console.error('An error occurred', error); // for demo purposes only
    return Promise.reject(error.message || error);
  }
  
  
  /*Use this method to add or update a list item with multiple columns.*/
  addOrUpdateListItem(itemToUpdate:SharepointListItem,keyValuePairs: [string, string][]):Promise<number>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		let perColumnString = `<Field Name="fieldName">fieldValue</Field>`;
		let allColumnsToUpdate = '';
		let command="New";
		//if there is an existing 'ID' property with a value, then this is not new but update.
		for (let i = 0; i < keyValuePairs.length; i++) {
			if (keyValuePairs[i][0] == "ID" && keyValuePairs[i][1]) {
				command = "Update"
				break;
			}
		}
		keyValuePairs.forEach(entry =>
			allColumnsToUpdate= allColumnsToUpdate+perColumnString.replace("fieldName", entry[0]).replace("fieldValue",entry[1])
		);
		currentPayload = currentPayload.replace("listNamePayLoad", itemToUpdate.getListName());
		currentPayload = currentPayload.replace("updatesPayLoad", `<Batch OnError="Continue"><Method ID="1" Cmd="`+command+`">`+allColumnsToUpdate+`</Method></Batch>`);
	return this.http.post(itemToUpdate.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
		.toPromise()
		.then(function (res){
			return +$(res.text()).find("z\\:row:first, row:first").attr("ows_ID");
		})
	.catch(this.handleError);
	}
	
	 
}