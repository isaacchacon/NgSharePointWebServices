import {SharepointListItem} from './sharepoint-list-item';

export interface SharepointListItemConstructor{
	new (rawResponse?:any): SharepointListItem;
}