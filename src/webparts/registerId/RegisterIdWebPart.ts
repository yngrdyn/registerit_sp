import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "RegisterIdWebPartStrings";
import RegisterId from "./components/RegisterId";
import { IRegisterIdProps } from "./components/IRegisterIdProps";

export interface IRegisterIdWebPartProps {
	description: string;
}

export default class RegisterIdWebPart extends BaseClientSideWebPart<IRegisterIdWebPartProps> {
	public render(): void {
		const element: React.ReactElement<IRegisterIdProps> = React.createElement(
			RegisterId,
			{
				description: this.properties.description,
				listName: "Teams%20ID%205",
				spHttpClient: this.context.spHttpClient,
				siteUrl: this.context.pageContext.web.absoluteUrl,
				context: this.context,
			}
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse("1.0");
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField("description", {
									label: strings.DescriptionFieldLabel,
								}),
								PropertyPaneTextField("listName", {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
