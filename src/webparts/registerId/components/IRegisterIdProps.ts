import { SPHttpClient } from "@microsoft/sp-http";

export interface IRegisterIdProps {
	description: string;
	listName: string;
	spHttpClient: SPHttpClient;
	siteUrl: string;
	context: any;
}
