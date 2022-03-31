import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import * as strings from "ZogenixDocumentListWebPartStrings";
import DynamicTabs from "./components/ZogenixDocumentList";
import { IDynamicTabsProps } from "./components/IZogenixDocumentListProps";
import spservices from "../../services/spservices";

export interface IDynamicTabsWebPartProps {
  description: string;
  context: any;
  list: any;
  multiSelect: any;
  ShowNameModified: boolean;
  ShowSubFolders: boolean;
  SiteUrl: any;
  spweb: any;
}

export default class DynamicTabsWebPart extends BaseClientSideWebPart<IDynamicTabsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = [];
  private spService: spservices = null;

  public async onInit(): Promise<void> {
    this.spService = new spservices(this.context, this.properties.SiteUrl);
    this.loadLists();
  }

  public render(): void {
    const element: React.ReactElement<IDynamicTabsProps> = React.createElement(
      DynamicTabs,
      {
        description: this.properties.description,
        context: this.context,
        list: this.properties.list,
        multiSelect: this.properties.multiSelect,
        ShowNameModified: this.properties.ShowNameModified,
        ShowSubFolders: this.properties.ShowSubFolders,
        SiteUrl: this.properties.SiteUrl,
        spweb: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onPropertyPaneConfigurationStart() {
    try {
      const _lists = await this.loadLists();
      this.lists = _lists;
      this.context.propertyPane.refresh();
    } catch (error) {}
  }

  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    const _lists: IPropertyPaneDropdownOption[] = [];
    try {
      const results = await this.spService.getSiteLists(
        this.properties.SiteUrl
      );
      for (const list of results) {
        _lists.push({ key: list.Id, text: list.Title });
      }
      // push new item value
    } catch (error) {
      this.context.propertyPane.refresh();
    }
    return _lists;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("SiteUrl", {
                  label: "Site Url",
                }),
                PropertyFieldMultiSelect("multiSelect", {
                  key: "multiSelect",
                  label: "Select Document Library",
                  options: this.lists,
                  selectedKeys: this.properties.multiSelect,
                }),
                PropertyPaneToggle("ShowNameModified", {
                  key: "ShowNameModified",
                  label: "",
                  checked: false,
                  onText: "Hide Created By Name and Date",
                  offText: "Show Created By Name and Date",
                }),
                PropertyPaneToggle("ShowSubFolders", {
                  label: "",
                  checked: false,
                  onText: "Hide Sub Folders as Tabs",
                  offText: "Show Sub Folders as Tabs",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
