import {Injectable} from '@angular/core';
import 'rxjs/add/operator/toPromise';
import {XmlUtilities} from './xml-utilities';
//new http client 
import { HttpClient, HttpHeaders } from '@angular/common/http';


declare var $:any;

import {SharepointListItem} from './sharepoint-list-item';
import {SharepointListItemConstructor} from './sharepoint-list-item-constructor';

@Injectable()
export class SharepointListsWebService{

	constructor(private http: HttpClient) { }
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

	private headers = new HttpHeaders().set('Content-Type','application/soap+xml; charset=utf-8');

	///If Filter query is not empty, we will use it to filter the CAML QUery.						 
	//Filter QUery : ifr provided, need to be a field Equals to a Value.
	//camlQuery: If provided, the caml query is applied.
	// If no filter nor caml query, an empty query is passed (likely to get all the items).
    getListItems(ctor:SharepointListItemConstructor, filterQuery:[string,string], camlQuery:string, orderByField:string):Promise<SharepointListItem[]>{
		let dummyInstance = new ctor();
		let currentPayload = this.xmlPayloadWrapperStart+this.getListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad",XmlUtilities.escapeString(dummyInstance.getListName()));
		currentPayload = currentPayload.replace("viewNamePayLoad", "");
		if(filterQuery||camlQuery){
		
			if(filterQuery){
				let internalCamlQuery = "<Query><Where><Eq><FieldRef Name='FieldName' /><Value Type='Text'>ValueName</Value></Eq></Where></Query>";
				internalCamlQuery = internalCamlQuery.replace("FieldName", filterQuery[0]);
				internalCamlQuery = internalCamlQuery.replace("ValueName", XmlUtilities.escapeString(filterQuery[1]));
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
		return this.http.post(dummyInstance.getSiteUrl()+this.serviceUrl,currentPayload,{headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function(res)
			{
				let filledResponse:SharepointListItem[] = [];
				$(res).find("z\\:row, row").each(function( index:any ) {
					filledResponse.push(new ctor($(this)[0].attributes)); 
				});
				return filledResponse;
		})
		.catch(this.handleError);
	}
	
	
	/*Use this method for updating only one column in the list.*/
	updateListItem(itemToUpdate:SharepointListItem,newValue:string):Promise<void>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad",XmlUtilities.escapeString( itemToUpdate.getListName()));
		currentPayload = currentPayload.replace("updatesPayLoad", `<Batch OnError="Continue"><Method ID="1" Cmd="Update"><Field Name="ID">`+itemToUpdate.ID+`</Field><Field Name="`+itemToUpdate.getFieldToUpdate()+`">`+XmlUtilities.escapeString(newValue)+`</Field></Method></Batch>`);
	return this.http.post(itemToUpdate.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,responseType:'text'})
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
			allColumnsToUpdate= allColumnsToUpdate+perColumnString.replace("fieldName", entry[0]).replace("fieldValue",XmlUtilities.escapeString(entry[1]))
		);
		currentPayload = currentPayload.replace("listNamePayLoad", XmlUtilities.escapeString(itemToUpdate.getListName()));
		currentPayload = currentPayload.replace("updatesPayLoad", `<Batch OnError="Continue"><Method ID="1" Cmd="`+command+`">`+allColumnsToUpdate+`</Method></Batch>`);
	return this.http.post(itemToUpdate.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,responseType:'text'})
		.toPromise()
		.then(function (res){
			return +$(res).find("z\\:row:first, row:first").attr("ows_ID");
		})
	.catch(this.handleError);
	}
	
	
	//Removes multiple list items at the same time ( In a batch operation)
	//parameter array should not be null NOR empty.
	//items should belong to the exact same list.
	//returns an array containing the errors or an empty array.
	removeListItems(itemsToRemove:SharepointListItem[]):Promise<string[]>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad", XmlUtilities.escapeString(itemsToRemove[0].getListName()));
		let updatesPayload = '<Batch OnError="Continue">';		
		for(let methodId=1;methodId<=itemsToRemove.length; methodId++){
			updatesPayload+='<Method ID="'+methodId+'" Cmd="Delete"><Field Name="ID">'+itemsToRemove[methodId-1].ID+'</Field></Method>';
		}
		updatesPayload+="</Batch>";
		currentPayload = currentPayload.replace("updatesPayLoad",updatesPayload);
		return this.http.post(itemsToRemove[0].getSiteUrl()+this.serviceUrl,currentPayload ,{headers:this.headers,responseType:'text'})
		.toPromise().then(function(res){
			let errorList = [];
			$(res).find("result:not(:contains('0x00000000'))").each(function( index:any ) {
					errorList.push($(this).find('ErrorCode').text()); 
				});
			return Promise.resolve(errorList);
		})
		.catch(this.handleError);
	}
	
	//Should add or update the list items that are given to it.
	//items should belong to the same list.
	//returns array of items, if added, it will update the ID on them.
	//if an error occurs, it will return a non empty string on the second part of the tuple.
	addOrUpdateListItems(itemsToAddOrUpdate:SharepointListItem[]):Promise<[SharepointListItem,string][]>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad",XmlUtilities.escapeString( itemsToAddOrUpdate[0].getListName()));
		let updatesPayload = '<Batch OnError="Continue">';	
		let idFieldXml:string;
		for(let i=0;i<itemsToAddOrUpdate.length;i++){
			idFieldXml= "";
			if(itemsToAddOrUpdate[i].ID && itemsToAddOrUpdate[i].ID>0)
			idFieldXml ='<Field Name="ID">'+itemsToAddOrUpdate[i].ID+'</Field>';
			updatesPayload+='<Method ID="'+(i+1)+'" Cmd="'+(idFieldXml?"Update":"New")+'">';
			if(idFieldXml)
			updatesPayload+=idFieldXml;
			for(let property of itemsToAddOrUpdate[i].getItemProperties()){
				if(property!="ID"){
					updatesPayload+='<Field Name="'+property+'">'+XmlUtilities.escapeString(itemsToAddOrUpdate[i][property])+'</Field>';
				}
			}
			updatesPayload+="</Method>";
		}
		updatesPayload+="</Batch>";
		currentPayload = currentPayload.replace("updatesPayLoad",updatesPayload);
		
		return this.http.post(itemsToAddOrUpdate[0].getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,responseType:'text'})
		.toPromise().then((res)=>{
			let finalResult:[SharepointListItem,string][] = [];
			let counter = 0;
			let currentItem :SharepointListItem;
			$(res).find("Result").each(function(index:any){
				// first, check for errors...
				currentItem = itemsToAddOrUpdate[counter];
				if($(this).find('ErrorCode').text()!='0x00000000'){
					finalResult.push([currentItem,$(this).find('ErrorCode').text()]);
				}
				else
				{
					if((!currentItem.ID )|| currentItem.ID==0){
						currentItem.ID = + $(this).find("z\\:row, row").attr("ows_ID");
					}
					finalResult.push([currentItem,""]);
				}
				counter++;
			});
			return finalResult;
		}).catch(this.handleError);
	}
	
}