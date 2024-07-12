import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "GroupMembersWebPartStrings";
import GroupMembers from "./components/GroupMembers";
import { IGroupMembersProps } from "./components/IGroupMembersProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGroup } from "./models/IGroup";

export interface IGroupMembersWebPartProps {
  description: string;
}

export default class GroupMembersWebPart extends BaseClientSideWebPart<IGroupMembersWebPartProps> {
  private _groups: IGroup[] = [];

  public render(): void {
    const element: React.ReactElement<IGroupMembersProps> = React.createElement(
      GroupMembers,
      {
        groups: this._groups,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this.getGroups();
    return Promise.resolve();
  }

  private async getGroups(): Promise<void> {
    await this.context.msGraphClientFactory
      .getClient("3")
      .then(async (client: MSGraphClientV3) => {
        await client
          .api("groups")
          .version("v1.0")
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }

            // Map the JSON response to the output array
            this._groups = res.value;
          });
      });
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
              groupFields: [],
            },
          ],
        },
      ],
    };
  }
}
