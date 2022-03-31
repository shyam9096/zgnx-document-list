import * as React from "react";
import styles from "./ZogenixDocumentList.module.scss";
import { IDynamicTabsProps } from "./IZogenixDocumentListProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  PivotItem,
  IPivotItemProps,
  Pivot,
} from "office-ui-fabric-react/lib/Pivot";
import {
  Panel,
  PanelType,
  TextField,
  PrimaryButton,
} from "office-ui-fabric-react";
import { Label } from "office-ui-fabric-react/lib";
import spservices from "../../../services/spservices";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/search";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files/folder";
import { Web } from "@pnp/sp/webs";

let ListName: any = [];
let SubFolder: any = [];
let ShowNameModified: any;
let ShowSubFolders: boolean;
var SubFolderList: any = [];

interface IDynamicTabsState {
  listData: any;
  SubFolderName: any;
  SubFolderList: any;
  subFolders: any;
}
const pivotStyles = {
  linkIsSelected: {
    selectors: {
      ":before": {
        height: "0px",
      },
    },
  },
  link: {
    ":hover": {
      backgroundColor: "none",
    },
  },
};

export default class DynamicTabs extends React.Component<
  IDynamicTabsProps,
  IDynamicTabsState
> {
  private spService: spservices = null;
  public constructor(props) {
    super(props);
    ShowNameModified = this.props.ShowNameModified;
    this.state = {
      listData: [],
      SubFolderName: [],
      SubFolderList: [],
      subFolders: [],
    };
    const Siteurl = this.props.SiteUrl;
    this.spService = new spservices(this.props.context, Siteurl);
  }

  public async componentDidMount() {
    if (this.props.ShowSubFolders == true) {
      this.Urlsubfolder();
    } else {
      this.getTabData();
    }
  }

  public async getTabData() {
    var listData: any = [];
    if (this.props.multiSelect) {
      if (this.props.multiSelect.length > 0) {
        if (this.props.SiteUrl != "" && this.props.SiteUrl != undefined) {
          // this.Urlsubfolder();

          for (var i = 0; i < this.props.multiSelect.length; i++) {
            var listDetails = await this.spService.getListDetails(
              this.props.multiSelect[i],
              this.props.SiteUrl
            );

            let List = listDetails["EntityTypeName"];
            let Libraryname: any = List.replace("_x0020_", "%20");
            var tabJson: any = [];

            var ListFolders = await this.spService.getListItem(
              this.props.SiteUrl,
              Libraryname,
              "LinkFilename,Modified,Editor/Title,FileRef,EncodedAbsUrl,Title,LinkURL,Files",
              "Editor"
            );

            let tabData = await Web(this.props.SiteUrl)
              .getFolderByServerRelativeUrl(listDetails["Title"])
              .files.expand("ListItemAllFields")
              .get();

            tabJson["Title"] = listDetails["Title"];
            tabJson["TabData"] = tabData;
            listData.push(tabJson);
            ListName.push(tabJson["Title"]);
            console.log(listData);
            listData.sort();
          }
        } else {
          // this.subfolder();
          for (var i = 0; i < this.props.multiSelect.length; i++) {
            var listDetails = await this.spService.getListDetails(
              this.props.multiSelect[i],
              this.props.SiteUrl
            );
            let List = listDetails["EntityTypeName"];
            let Libraryname: any = List.replace("_x0020_", "%20");
            var tabJson: any = [];

            var ListFolders = await this.spService.getListItem(
              this.props.SiteUrl,
              Libraryname,
              "LinkFilename,Modified,Editor/Title,FileRef,EncodedAbsUrl,Title,LinkURL",
              "Editor"
            );

            var tabData = await sp.web
              .getFolderByServerRelativeUrl(listDetails["Title"])
              .files.expand("ListItemAllFields")
              .get();

            tabJson["Title"] = listDetails["Title"];
            tabJson["TabData"] = tabData;
            listData.push(tabJson);
            ListName.push(tabJson["Title"]);
            console.log(listData);
          }
        }
        this.setState({ listData: listData });
        console.log(SubFolder);
      }
    }
  }

  public async Urlsubfolder() {
    if (this.props.multiSelect.length > 0) {
      if (this.props.SiteUrl != "" && this.props.SiteUrl != undefined) {
        this.props.multiSelect.map(async (FolderID) => {
          var listDetails = await this.spService.getListDetails(
            FolderID,
            this.props.SiteUrl
          );
          let List = listDetails["EntityTypeName"];
          let Libraryname: any = List.replace("_x0020_", "%20");
          var tabJson: any = [];
          var ListFolders = await this.spService.getListItem(
            this.props.SiteUrl,
            Libraryname,
            "LinkFilename,Modified,Editor/Title,FileRef,EncodedAbsUrl,Title,LinkURL,Files",
            "Editor"
          );
          ListFolders["Folders"].map(async (FolderItems) => {
            var SubFolders: any = [];
            const ListPath = Libraryname + "/" + FolderItems.Name;
            const files = await Web(this.props.SiteUrl)
              .getFolderByServerRelativeUrl(ListPath)
              .files.expand("ListItemAllFields")
              .get();
            console.log(files);

            SubFolders["Title"] = FolderItems.Name;
            SubFolders["Files"] = files;
            if (SubFolders["Title"] != "Forms") {
              this.state.SubFolderList.push(SubFolders);
              console.log(SubFolderList);
            }
            this.getTabData();
          });
        });
        this.setState({ subFolders: this.state.SubFolderList });
        // this.getTabData();
      } else {
        this.props.multiSelect.map(async (FolderID) => {
          var listDetails = await this.spService.getListDetails(
            FolderID,
            this.props.SiteUrl
          );
          let List = listDetails["EntityTypeName"];
          let Libraryname: any = List.replace("_x0020_", "%20");

          var ListFolders = await this.spService.getListItem(
            this.props.SiteUrl,
            Libraryname,
            "LinkFilename,Modified,Editor/Title,FileRef,EncodedAbsUrl,Title,LinkURL",
            "Editor"
          );

          ListFolders["Folders"].map(async (FolderItems) => {
            var SubFolders: any = [];
            const ListPath = Libraryname + "/" + FolderItems.Name;
            const files = await sp.web
              .getFolderByServerRelativeUrl(ListPath)
              .files.expand("ListItemAllFields")
              .get();
            console.log(files);

            SubFolders["Title"] = FolderItems.Name;
            SubFolders["Files"] = files;
            if (SubFolders["Title"] != "Forms") {
              this.state.SubFolderList.push(SubFolders);
              console.log(SubFolderList);
            }
            this.getTabData();
          });
        });
        this.setState({ subFolders: this.state.SubFolderList });
      }
    }
  }

  // public async subfolder() {
  //   this.props.multiSelect.map(async (FolderID) => {
  //     var listDetails = await this.spService.getListDetails(
  //       FolderID,
  //       this.props.SiteUrl
  //     );
  //     let List = listDetails["EntityTypeName"];
  //     let Libraryname: any = List.replace("_x0020_", "%20");

  //     var ListFolders = await this.spService.getListItem(
  //       this.props.SiteUrl,
  //       Libraryname,
  //       "LinkFilename,Modified,Editor/Title,FileRef,EncodedAbsUrl,Title,LinkURL",
  //       "Editor"
  //     );

  //     ListFolders["Folders"].map(async (FolderItems) => {
  //       var SubFolders: any = [];
  //       const ListPath = Libraryname + "/" + FolderItems.Name;
  //       const files = await sp.web
  //         .getFolderByServerRelativeUrl(ListPath)
  //         .files.expand("ListItemAllFields")
  //         .get();
  //       console.log(files);

  //       SubFolders["Title"] = FolderItems.Name;
  //       SubFolders["Files"] = files;
  //       if (SubFolders["Title"] != "Forms") {
  //         this.state.SubFolderList.push(SubFolders);
  //         console.log(SubFolderList);
  //       }
  //     });
  //   });
  //   this.setState({ subFolders: this.state.SubFolderList });
  // }

  public render(): React.ReactElement<IDynamicTabsProps> {
    return (
      <div className={styles.dynamicTabs}>
        <div className={styles.container}>
          {this.props.multiSelect ? (
            <>
              {this.props.ShowSubFolders != true && (
                <Pivot className={styles.pivot} styles={pivotStyles}>
                  {this.state.listData
                    .sort((a, b) => a.Title.localeCompare(b.Title))
                    .map((data, index) => {
                      return (
                        <PivotItem linkText={data.Title}>
                          <div className={styles.headerrow}>
                            <div>Name</div>
                            {this.props.ShowNameModified == true && (
                              <>
                                <div>Date Modified</div>
                                {/* <div>Modified By</div> */}
                              </>
                            )}
                          </div>
                          {data.TabData.sort((a, b) =>
                            a.Name.localeCompare(b.Name)
                          ).map((item) => {
                            console.log(item);

                            var filePath = item.ServerRelativeUrl;
                            if (filePath.slice(-4) == ".url") {
                              filePath = item.ListItemAllFields.LinkURL.Url;
                            }
                            return (
                              <div className={styles.row}>
                                <div>
                                  <a
                                    href={filePath}
                                    target="_blank"
                                    data-interception="off"
                                  >
                                    {item.Name}
                                  </a>
                                </div>{" "}
                                {this.props.ShowNameModified == true && (
                                  <>
                                    <div>{item.TimeLastModified}</div>
                                  </>
                                )}
                              </div>
                            );
                          })}
                        </PivotItem>
                      );
                    })}
                </Pivot>
              )}
              {this.props.ShowSubFolders == true && (
                <Pivot className={styles.pivot} styles={pivotStyles}>
                  {this.state.subFolders
                    .sort((a, b) => a.Title.localeCompare(b.Title))
                    .map((data, index) => {
                      return (
                        <PivotItem linkText={data.Title}>
                          <div className={styles.headerrow}>
                            <div>Name</div>
                            {this.props.ShowNameModified == true && (
                              <>
                                <div>Date Modified</div>
                                {/* <div>Modified By</div> */}
                              </>
                            )}
                          </div>
                          {data.Files.sort((a, b) =>
                            a.Name.localeCompare(b.Name)
                          ).map((item) => {
                            console.log(item);
                            var filePath = item.ServerRelativeUrl;
                            if (filePath.slice(-4) == ".url") {
                              filePath = item.ListItemAllFields.LinkURL.Url;
                            }
                            return (
                              <div className={styles.row}>
                                <div>
                                  <a
                                    href={filePath}
                                    target="_blank"
                                    data-interception="off"
                                  >
                                    {item.Name}
                                  </a>
                                </div>{" "}
                                {this.props.ShowNameModified == true && (
                                  <>
                                    <div>{item.TimeLastModified}</div>
                                    {/* <div>Name</div> */}
                                  </>
                                )}
                              </div>
                            );
                          })}
                          {/*  */}
                        </PivotItem>
                      );
                    })}
                </Pivot>
              )}
            </>
          ) : (
            <h2>Configure the webpart property to show data on page</h2>
          )}
        </div>
      </div>
    );
  }
}
