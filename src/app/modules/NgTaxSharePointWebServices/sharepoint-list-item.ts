

export abstract class SharepointListItem{
/*itemId:number;*/
ID:number;
title?:string;	

constructor(rawResponse?:any){
	if(rawResponse){
		this.ID = +this.findString('ows_id',rawResponse);
		for(let x = 0; x<this.getItemProperties().length; x++){
				this[this.getItemProperties()[x]] = this.findString('ows_'+this.getItemProperties()[x].toLowerCase(),rawResponse);
			}			
		}
}
	
	protected findString(searchTerm:string, myArray:any):string{
		for(let x:number = 0;x<myArray.length;x++)
			{
				if(myArray[x].name ==searchTerm)
				{
					return myArray[x].value;
				}
			}
			return '';
	}
	
	abstract getSiteUrl():string;
	abstract getListName():string;
	abstract getFieldToUpdate():string;
	
	/*New method that needs to return an array of the properties that will be fetched.*/
	getItemProperties():string[]{	
		return [];
	}
	
	toBoolean(aString:string):void{
		if(this[aString]){
			this[aString]= this[aString] && this[aString]=="true";
		}
	}
	
	toInteger(aString:string):void{
		if(this[aString] && this[aString].includes(".")){
				this[aString] = this[aString].split(".")[0];
			}
	}

}