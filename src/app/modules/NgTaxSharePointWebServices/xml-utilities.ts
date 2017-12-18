export class XmlUtilities{
	static escapeString(stringToEscape:string):string{
		if(stringToEscape)
		return (stringToEscape+"")
			   .replace(/&(?!(amp;)|(lt;)|(gt;)|(quot;)|(#39;)|(apos;))/g, "&amp;")
               .replace(/</g, '&lt;')
               .replace(/>/g, '&gt;')
               .replace(/"/g, '&quot;')
               .replace(/'/g, '&apos;');
		return stringToEscape;
	}
}