import {SharepointListItem} from './sharepoint-list-item';
declare var $:any;

export class UserInfoListEntry extends SharepointListItem{

constructor(rawResponse?:any){
		super(rawResponse);
		if(rawResponse){
			/*custom code goes here - nothing for the moment.*/
		}
	}
	
	/*They don't have to match to the internal column name because we are not writing  on this list.
	Otherwise, we would have to do that*/
	getItemProperties():string[]{	
		return ["title", "jobTitle",  "name", "email" ]
	}
	
	static urlOfSubsite = "/";
	
	///user profile service always run on root web .
	getSiteUrl():string{
		return UserInfoListEntry.urlOfSubsite;
	}
	getListName():string{
		return 'UserInfo';
	}
	getFieldToUpdate():string{
		return 'Not implemented';
	}

}